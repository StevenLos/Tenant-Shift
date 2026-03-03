<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-000100

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M0214-Update-ExchangeOnPremDistributionLists_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
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

Write-Status -Message 'Starting Exchange on-prem distribution list update script.'
Ensure-ExchangeOnPremConnection

$setDistributionGroupCommand = Get-Command -Name Set-DistributionGroup -ErrorAction Stop
$supportsSendModerationNotifications = $setDistributionGroupCommand.Parameters.ContainsKey('SendModerationNotifications')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $distributionGroupIdentity = Get-TrimmedValue -Value $row.DistributionGroupIdentity

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

        $recipientTypeDetails = Get-TrimmedValue -Value $group.RecipientTypeDetails
        if ($recipientTypeDetails -ne 'MailUniversalDistributionGroup' -and $recipientTypeDetails -ne 'MailNonUniversalGroup') {
            throw "Recipient '$distributionGroupIdentity' is '$recipientTypeDetails'. Expected a distribution list recipient type."
        }

        $setParams = @{
            Identity = $group.Identity
        }

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $setParams.DisplayName = $displayName
        }

        $primarySmtpAddress = Get-TrimmedValue -Value $row.PrimarySmtpAddress
        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $setParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $managedBy = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ManagedBy)
        if ($managedBy.Count -gt 0) {
            $setParams.ManagedBy = $managedBy
        }

        $notes = Get-TrimmedValue -Value $row.Notes
        if (-not [string]::IsNullOrWhiteSpace($notes)) {
            $setParams.Notes = $notes
        }

        $memberJoinRestriction = Get-TrimmedValue -Value $row.MemberJoinRestriction
        if (-not [string]::IsNullOrWhiteSpace($memberJoinRestriction)) {
            $setParams.MemberJoinRestriction = $memberJoinRestriction
        }

        $memberDepartRestriction = Get-TrimmedValue -Value $row.MemberDepartRestriction
        if (-not [string]::IsNullOrWhiteSpace($memberDepartRestriction)) {
            $setParams.MemberDepartRestriction = $memberDepartRestriction
        }

        $moderationEnabledRaw = Get-TrimmedValue -Value $row.ModerationEnabled
        if (-not [string]::IsNullOrWhiteSpace($moderationEnabledRaw)) {
            $setParams.ModerationEnabled = ConvertTo-Bool -Value $moderationEnabledRaw
        }

        $moderatedBy = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ModeratedBy)
        if ($moderatedBy.Count -gt 0) {
            $setParams.ModeratedBy = $moderatedBy
        }

        $requireAuthRaw = Get-TrimmedValue -Value $row.RequireSenderAuthenticationEnabled
        if (-not [string]::IsNullOrWhiteSpace($requireAuthRaw)) {
            $setParams.RequireSenderAuthenticationEnabled = ConvertTo-Bool -Value $requireAuthRaw
        }

        $hiddenRaw = Get-TrimmedValue -Value $row.HiddenFromAddressListsEnabled
        if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
            $setParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
        }

        $acceptMessagesOnlyFrom = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.AcceptMessagesOnlyFrom)
        if ($acceptMessagesOnlyFrom.Count -gt 0) {
            $setParams.AcceptMessagesOnlyFrom = $acceptMessagesOnlyFrom
        }

        $acceptMessagesOnlyFromDLMembers = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.AcceptMessagesOnlyFromDLMembers)
        if ($acceptMessagesOnlyFromDLMembers.Count -gt 0) {
            $setParams.AcceptMessagesOnlyFromDLMembers = $acceptMessagesOnlyFromDLMembers
        }

        $rejectMessagesFrom = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.RejectMessagesFrom)
        if ($rejectMessagesFrom.Count -gt 0) {
            $setParams.RejectMessagesFrom = $rejectMessagesFrom
        }

        $rejectMessagesFromDLMembers = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.RejectMessagesFromDLMembers)
        if ($rejectMessagesFromDLMembers.Count -gt 0) {
            $setParams.RejectMessagesFromDLMembers = $rejectMessagesFromDLMembers
        }

        $bypassModeration = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.BypassModerationFromSendersOrMembers)
        if ($bypassModeration.Count -gt 0) {
            $setParams.BypassModerationFromSendersOrMembers = $bypassModeration
        }

        $sendModerationNotifications = Get-TrimmedValue -Value $row.SendModerationNotifications
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

        if ($PSCmdlet.ShouldProcess($distributionGroupIdentity, 'Update Exchange on-prem distribution list')) {
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
Write-Status -Message 'Exchange on-prem distribution list update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
