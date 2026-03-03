<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-005957

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_P3115-Create-ExchangeOnlineMailEnabledSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'Name',
    'Alias',
    'DisplayName',
    'PrimarySmtpAddress',
    'ManagedBy',
    'Notes',
    'RequireSenderAuthenticationEnabled',
    'HiddenFromAddressListsEnabled',
    'ModerationEnabled',
    'ModeratedBy',
    'AcceptMessagesOnlyFrom',
    'AcceptMessagesOnlyFromDLMembers',
    'RejectMessagesFrom',
    'RejectMessagesFromDLMembers',
    'BypassModerationFromSendersOrMembers',
    'SendModerationNotifications'
)

Write-Status -Message 'Starting Exchange Online mail-enabled security group creation script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$setDistributionGroupCommand = Get-Command -Name Set-DistributionGroup -ErrorAction Stop
$supportsSendModerationNotifications = $setDistributionGroupCommand.Parameters.ContainsKey('SendModerationNotifications')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $name = ([string]$row.Name).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Name is required.'
        }

        $alias = ([string]$row.Alias).Trim()
        if ([string]::IsNullOrWhiteSpace($alias)) {
            $alias = ($name -replace '[^a-zA-Z0-9]', '')
            if ([string]::IsNullOrWhiteSpace($alias)) {
                throw 'Alias is empty after sanitization. Provide an Alias value in the CSV.'
            }
        }

        $primarySmtpAddress = ([string]$row.PrimarySmtpAddress).Trim()
        $identityToCheck = if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) { $primarySmtpAddress } else { $name }

        $existingGroup = Invoke-WithRetry -OperationName "Lookup mail-enabled security group $identityToCheck" -ScriptBlock {
            Get-DistributionGroup -Identity $identityToCheck -ErrorAction SilentlyContinue
        }

        if ($existingGroup) {
            $recipientTypeDetails = ([string]$existingGroup.RecipientTypeDetails).Trim()
            if ($recipientTypeDetails -eq 'MailUniversalSecurityGroup') {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateMailEnabledSecurityGroup' -Status 'Skipped' -Message 'Mail-enabled security group already exists.'))
                $rowNumber++
                continue
            }

            throw "Recipient '$identityToCheck' already exists with type '$recipientTypeDetails'."
        }

        $createParams = @{
            Name  = $name
            Alias = $alias
            Type  = 'Security'
        }

        $displayName = ([string]$row.DisplayName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $createParams.DisplayName = $displayName
        }

        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $createParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $managedBy = ConvertTo-Array -Value ([string]$row.ManagedBy)
        if ($managedBy.Count -gt 0) {
            $createParams.ManagedBy = $managedBy
        }

        if ($PSCmdlet.ShouldProcess($identityToCheck, 'Create Exchange Online mail-enabled security group')) {
            $createdGroup = Invoke-WithRetry -OperationName "Create mail-enabled security group $identityToCheck" -ScriptBlock {
                New-DistributionGroup @createParams -ErrorAction Stop
            }

            $setParams = @{
                Identity = $createdGroup.Identity
            }

            $notes = ([string]$row.Notes).Trim()
            if (-not [string]::IsNullOrWhiteSpace($notes)) {
                $setParams.Notes = $notes
            }

            $requireAuthRaw = ([string]$row.RequireSenderAuthenticationEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($requireAuthRaw)) {
                $setParams.RequireSenderAuthenticationEnabled = ConvertTo-Bool -Value $requireAuthRaw
            }

            $hiddenRaw = ([string]$row.HiddenFromAddressListsEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
                $setParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
            }

            $moderationEnabledRaw = ([string]$row.ModerationEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($moderationEnabledRaw)) {
                $setParams.ModerationEnabled = ConvertTo-Bool -Value $moderationEnabledRaw
            }

            $moderatedBy = ConvertTo-Array -Value ([string]$row.ModeratedBy)
            if ($moderatedBy.Count -gt 0) {
                $setParams.ModeratedBy = $moderatedBy
            }

            $acceptMessagesOnlyFrom = ConvertTo-Array -Value ([string]$row.AcceptMessagesOnlyFrom)
            if ($acceptMessagesOnlyFrom.Count -gt 0) {
                $setParams.AcceptMessagesOnlyFrom = $acceptMessagesOnlyFrom
            }

            $acceptMessagesOnlyFromDLMembers = ConvertTo-Array -Value ([string]$row.AcceptMessagesOnlyFromDLMembers)
            if ($acceptMessagesOnlyFromDLMembers.Count -gt 0) {
                $setParams.AcceptMessagesOnlyFromDLMembers = $acceptMessagesOnlyFromDLMembers
            }

            $rejectMessagesFrom = ConvertTo-Array -Value ([string]$row.RejectMessagesFrom)
            if ($rejectMessagesFrom.Count -gt 0) {
                $setParams.RejectMessagesFrom = $rejectMessagesFrom
            }

            $rejectMessagesFromDLMembers = ConvertTo-Array -Value ([string]$row.RejectMessagesFromDLMembers)
            if ($rejectMessagesFromDLMembers.Count -gt 0) {
                $setParams.RejectMessagesFromDLMembers = $rejectMessagesFromDLMembers
            }

            $bypassModeration = ConvertTo-Array -Value ([string]$row.BypassModerationFromSendersOrMembers)
            if ($bypassModeration.Count -gt 0) {
                $setParams.BypassModerationFromSendersOrMembers = $bypassModeration
            }

            $sendModerationNotifications = ([string]$row.SendModerationNotifications).Trim()
            if (-not [string]::IsNullOrWhiteSpace($sendModerationNotifications)) {
                if ($supportsSendModerationNotifications) {
                    $setParams.SendModerationNotifications = $sendModerationNotifications
                }
                else {
                    Write-Status -Message 'Set-DistributionGroup does not support -SendModerationNotifications in this session. Value ignored.' -Level WARN
                }
            }

            if ($setParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Set mail-enabled security group options $identityToCheck" -ScriptBlock {
                    Set-DistributionGroup @setParams -ErrorAction Stop
                }
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateMailEnabledSecurityGroup' -Status 'Created' -Message 'Mail-enabled security group created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateMailEnabledSecurityGroup' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($name) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $name -Action 'CreateMailEnabledSecurityGroup' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mail-enabled security group creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





