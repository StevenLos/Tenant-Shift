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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M0222-Update-ExchangeOnPremMailEnabledSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
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

Write-Status -Message 'Starting Exchange on-prem mail-enabled security group update script.'
Ensure-ExchangeOnPremConnection

$setDistributionGroupCommand = Get-Command -Name Set-DistributionGroup -ErrorAction Stop
$supportsSendModerationNotifications = $setDistributionGroupCommand.Parameters.ContainsKey('SendModerationNotifications')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $securityGroupIdentity = Get-TrimmedValue -Value $row.SecurityGroupIdentity

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

        $recipientTypeDetails = Get-TrimmedValue -Value $group.RecipientTypeDetails
        if ($recipientTypeDetails -ne 'MailUniversalSecurityGroup') {
            throw "Recipient '$securityGroupIdentity' is '$recipientTypeDetails'. Expected MailUniversalSecurityGroup."
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

        $requireAuthRaw = Get-TrimmedValue -Value $row.RequireSenderAuthenticationEnabled
        if (-not [string]::IsNullOrWhiteSpace($requireAuthRaw)) {
            $setParams.RequireSenderAuthenticationEnabled = ConvertTo-Bool -Value $requireAuthRaw
        }

        $hiddenRaw = Get-TrimmedValue -Value $row.HiddenFromAddressListsEnabled
        if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
            $setParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
        }

        $moderationEnabledRaw = Get-TrimmedValue -Value $row.ModerationEnabled
        if (-not [string]::IsNullOrWhiteSpace($moderationEnabledRaw)) {
            $setParams.ModerationEnabled = ConvertTo-Bool -Value $moderationEnabledRaw
        }

        $moderatedBy = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ModeratedBy)
        if ($moderatedBy.Count -gt 0) {
            $setParams.ModeratedBy = $moderatedBy
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
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $securityGroupIdentity -Action 'UpdateMailEnabledSecurityGroup' -Status 'Skipped' -Message 'No updates specified.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($securityGroupIdentity, 'Update Exchange on-prem mail-enabled security group')) {
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
Write-Status -Message 'Exchange on-prem mail-enabled security group update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
