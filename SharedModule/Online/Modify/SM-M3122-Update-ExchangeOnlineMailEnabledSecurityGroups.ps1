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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3122-Update-ExchangeOnlineMailEnabledSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'SecurityGroupIdentity',
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

Write-Status -Message 'Starting Exchange Online mail-enabled security group update script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$setDistributionGroupCommand = Get-Command -Name Set-DistributionGroup -ErrorAction Stop
$supportsSendModerationNotifications = $setDistributionGroupCommand.Parameters.ContainsKey('SendModerationNotifications')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $securityGroupIdentity = ([string]$row.SecurityGroupIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($securityGroupIdentity)) {
            throw 'SecurityGroupIdentity is required.'
        }

        $group = Invoke-WithRetry -OperationName "Lookup mail-enabled security group $securityGroupIdentity" -ScriptBlock {
            Get-DistributionGroup -Identity $securityGroupIdentity -ErrorAction SilentlyContinue
        }

        if (-not $group) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $securityGroupIdentity -Action 'UpdateMailEnabledSecurityGroup' -Status 'NotFound' -Message 'Mail-enabled security group not found.'))
            $rowNumber++
            continue
        }

        $recipientTypeDetails = ([string]$group.RecipientTypeDetails).Trim()
        if ($recipientTypeDetails -ne 'MailUniversalSecurityGroup') {
            throw "Recipient '$securityGroupIdentity' is '$recipientTypeDetails'. Expected MailUniversalSecurityGroup."
        }

        $setParams = @{
            Identity = $group.Identity
        }

        $displayName = ([string]$row.DisplayName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $setParams.DisplayName = $displayName
        }

        $primarySmtpAddress = ([string]$row.PrimarySmtpAddress).Trim()
        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $setParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $managedBy = ConvertTo-Array -Value ([string]$row.ManagedBy)
        if ($managedBy.Count -gt 0) {
            $setParams.ManagedBy = $managedBy
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

        if ($setParams.Count -eq 1) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $securityGroupIdentity -Action 'UpdateMailEnabledSecurityGroup' -Status 'Skipped' -Message 'No updates specified.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($securityGroupIdentity, 'Update Exchange Online mail-enabled security group')) {
            Invoke-WithRetry -OperationName "Update mail-enabled security group $securityGroupIdentity" -ScriptBlock {
                Set-DistributionGroup @setParams -ErrorAction Stop
            }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $securityGroupIdentity -Action 'UpdateMailEnabledSecurityGroup' -Status 'Updated' -Message 'Mail-enabled security group updated successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $securityGroupIdentity -Action 'UpdateMailEnabledSecurityGroup' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($securityGroupIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $securityGroupIdentity -Action 'UpdateMailEnabledSecurityGroup' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mail-enabled security group update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





