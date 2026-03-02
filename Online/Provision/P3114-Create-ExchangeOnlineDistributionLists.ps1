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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_P3114-Create-ExchangeOnlineDistributionLists_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'MemberJoinRestriction',
    'MemberDepartRestriction',
    'ModerationEnabled',
    'ModeratedBy',
    'RequireSenderAuthenticationEnabled',
    'HiddenFromAddressListsEnabled',
    'AcceptMessagesOnlyFrom',
    'AcceptMessagesOnlyFromDLMembers',
    'RejectMessagesFrom',
    'RejectMessagesFromDLMembers',
    'BypassModerationFromSendersOrMembers',
    'SendModerationNotifications'
)

Write-Status -Message 'Starting Exchange Online distribution list creation script.'
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

        $existingGroup = Invoke-WithRetry -OperationName "Lookup distribution list $identityToCheck" -ScriptBlock {
            Get-DistributionGroup -Identity $identityToCheck -ErrorAction SilentlyContinue
        }

        if ($existingGroup) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateDistributionList' -Status 'Skipped' -Message 'Distribution list already exists.'))
            $rowNumber++
            continue
        }

        $createParams = @{
            Name  = $name
            Alias = $alias
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

        if ($PSCmdlet.ShouldProcess($identityToCheck, 'Create Exchange Online distribution list')) {
            $createdGroup = Invoke-WithRetry -OperationName "Create distribution list $identityToCheck" -ScriptBlock {
                New-DistributionGroup @createParams -ErrorAction Stop
            }

            $setParams = @{
                Identity = $createdGroup.Identity
            }

            $notes = ([string]$row.Notes).Trim()
            if (-not [string]::IsNullOrWhiteSpace($notes)) {
                $setParams.Notes = $notes
            }

            $memberJoinRestriction = ([string]$row.MemberJoinRestriction).Trim()
            if (-not [string]::IsNullOrWhiteSpace($memberJoinRestriction)) {
                $setParams.MemberJoinRestriction = $memberJoinRestriction
            }

            $memberDepartRestriction = ([string]$row.MemberDepartRestriction).Trim()
            if (-not [string]::IsNullOrWhiteSpace($memberDepartRestriction)) {
                $setParams.MemberDepartRestriction = $memberDepartRestriction
            }

            $moderationEnabledRaw = ([string]$row.ModerationEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($moderationEnabledRaw)) {
                $setParams.ModerationEnabled = ConvertTo-Bool -Value $moderationEnabledRaw
            }

            $moderatedBy = ConvertTo-Array -Value ([string]$row.ModeratedBy)
            if ($moderatedBy.Count -gt 0) {
                $setParams.ModeratedBy = $moderatedBy
            }

            $requireAuthRaw = ([string]$row.RequireSenderAuthenticationEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($requireAuthRaw)) {
                $setParams.RequireSenderAuthenticationEnabled = ConvertTo-Bool -Value $requireAuthRaw
            }

            $hiddenRaw = ([string]$row.HiddenFromAddressListsEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
                $setParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
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
                Invoke-WithRetry -OperationName "Set distribution list options $identityToCheck" -ScriptBlock {
                    Set-DistributionGroup @setParams -ErrorAction Stop
                }
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateDistributionList' -Status 'Created' -Message 'Distribution list created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateDistributionList' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($name) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $name -Action 'CreateDistributionList' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online distribution list creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





