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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3114-Update-ExchangeOnlineDistributionLists_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'DistributionGroupIdentity',
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

Write-Status -Message 'Starting Exchange Online distribution list update script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$setDistributionGroupCommand = Get-Command -Name Set-DistributionGroup -ErrorAction Stop
$supportsSendModerationNotifications = $setDistributionGroupCommand.Parameters.ContainsKey('SendModerationNotifications')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $distributionGroupIdentity = ([string]$row.DistributionGroupIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($distributionGroupIdentity)) {
            throw 'DistributionGroupIdentity is required.'
        }

        $group = Invoke-WithRetry -OperationName "Lookup distribution list $distributionGroupIdentity" -ScriptBlock {
            Get-DistributionGroup -Identity $distributionGroupIdentity -ErrorAction SilentlyContinue
        }
        if (-not $group) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'UpdateDistributionList' -Status 'NotFound' -Message 'Distribution list not found.'))
            $rowNumber++
            continue
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

        if ($setParams.Count -eq 1) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'UpdateDistributionList' -Status 'Skipped' -Message 'No updates specified.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($distributionGroupIdentity, 'Update Exchange Online distribution list')) {
            Invoke-WithRetry -OperationName "Update distribution list $distributionGroupIdentity" -ScriptBlock {
                Set-DistributionGroup @setParams -ErrorAction Stop
            }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'UpdateDistributionList' -Status 'Updated' -Message 'Distribution list updated successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'UpdateDistributionList' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($distributionGroupIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'UpdateDistributionList' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online distribution list update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





