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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M3116-Update-ExchangeOnlineSharedMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'SharedMailboxIdentity',
    'DisplayName',
    'PrimarySmtpAddress',
    'HiddenFromAddressListsEnabled',
    'GrantSendOnBehalfTo',
    'MessageCopyForSentAsEnabled',
    'MessageCopyForSendOnBehalfEnabled',
    'ForwardingSmtpAddress',
    'DeliverToMailboxAndForward',
    'AuditEnabled',
    'LitigationHoldEnabled'
)

Write-Status -Message 'Starting Exchange Online shared mailbox update script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop
$supports = @{
    DisplayName                       = $setMailboxCommand.Parameters.ContainsKey('DisplayName')
    PrimarySmtpAddress                = $setMailboxCommand.Parameters.ContainsKey('PrimarySmtpAddress')
    HiddenFromAddressListsEnabled     = $setMailboxCommand.Parameters.ContainsKey('HiddenFromAddressListsEnabled')
    GrantSendOnBehalfTo               = $setMailboxCommand.Parameters.ContainsKey('GrantSendOnBehalfTo')
    MessageCopyForSentAsEnabled       = $setMailboxCommand.Parameters.ContainsKey('MessageCopyForSentAsEnabled')
    MessageCopyForSendOnBehalfEnabled = $setMailboxCommand.Parameters.ContainsKey('MessageCopyForSendOnBehalfEnabled')
    ForwardingSmtpAddress             = $setMailboxCommand.Parameters.ContainsKey('ForwardingSmtpAddress')
    DeliverToMailboxAndForward        = $setMailboxCommand.Parameters.ContainsKey('DeliverToMailboxAndForward')
    AuditEnabled                      = $setMailboxCommand.Parameters.ContainsKey('AuditEnabled')
    LitigationHoldEnabled             = $setMailboxCommand.Parameters.ContainsKey('LitigationHoldEnabled')
}

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $sharedMailboxIdentity = ([string]$row.SharedMailboxIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($sharedMailboxIdentity)) {
            throw 'SharedMailboxIdentity is required.'
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup shared mailbox $sharedMailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $sharedMailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox -or ([string]$mailbox.RecipientTypeDetails).Trim() -ne 'SharedMailbox') {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentity -Action 'UpdateSharedMailbox' -Status 'NotFound' -Message 'Shared mailbox not found.'))
            $rowNumber++
            continue
        }

        $setParams = @{
            Identity = $mailbox.Identity
        }
        $warnings = [System.Collections.Generic.List[string]]::new()

        $displayName = ([string]$row.DisplayName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            if ($supports.DisplayName) {
                $setParams.DisplayName = $displayName
            }
            else {
                $warnings.Add('DisplayName ignored (unsupported parameter).')
            }
        }

        $primarySmtpAddress = ([string]$row.PrimarySmtpAddress).Trim()
        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            if ($supports.PrimarySmtpAddress) {
                $setParams.PrimarySmtpAddress = $primarySmtpAddress
            }
            else {
                $warnings.Add('PrimarySmtpAddress ignored (unsupported parameter).')
            }
        }

        $hiddenRaw = ([string]$row.HiddenFromAddressListsEnabled).Trim()
        if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
            if ($supports.HiddenFromAddressListsEnabled) {
                $setParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
            }
            else {
                $warnings.Add('HiddenFromAddressListsEnabled ignored (unsupported parameter).')
            }
        }

        $grantSendOnBehalfRaw = ([string]$row.GrantSendOnBehalfTo).Trim()
        if (-not [string]::IsNullOrWhiteSpace($grantSendOnBehalfRaw)) {
            if ($supports.GrantSendOnBehalfTo) {
                if ($grantSendOnBehalfRaw -eq '-') {
                    $setParams.GrantSendOnBehalfTo = @()
                }
                else {
                    $setParams.GrantSendOnBehalfTo = ConvertTo-Array -Value $grantSendOnBehalfRaw
                }
            }
            else {
                $warnings.Add('GrantSendOnBehalfTo ignored (unsupported parameter).')
            }
        }

        $sentAsCopyRaw = ([string]$row.MessageCopyForSentAsEnabled).Trim()
        if (-not [string]::IsNullOrWhiteSpace($sentAsCopyRaw)) {
            if ($supports.MessageCopyForSentAsEnabled) {
                $setParams.MessageCopyForSentAsEnabled = ConvertTo-Bool -Value $sentAsCopyRaw
            }
            else {
                $warnings.Add('MessageCopyForSentAsEnabled ignored (unsupported parameter).')
            }
        }

        $sendOnBehalfCopyRaw = ([string]$row.MessageCopyForSendOnBehalfEnabled).Trim()
        if (-not [string]::IsNullOrWhiteSpace($sendOnBehalfCopyRaw)) {
            if ($supports.MessageCopyForSendOnBehalfEnabled) {
                $setParams.MessageCopyForSendOnBehalfEnabled = ConvertTo-Bool -Value $sendOnBehalfCopyRaw
            }
            else {
                $warnings.Add('MessageCopyForSendOnBehalfEnabled ignored (unsupported parameter).')
            }
        }

        $forwardingSmtpAddress = ([string]$row.ForwardingSmtpAddress).Trim()
        if (-not [string]::IsNullOrWhiteSpace($forwardingSmtpAddress)) {
            if ($supports.ForwardingSmtpAddress) {
                $setParams.ForwardingSmtpAddress = $forwardingSmtpAddress
            }
            else {
                $warnings.Add('ForwardingSmtpAddress ignored (unsupported parameter).')
            }
        }

        $deliverAndForwardRaw = ([string]$row.DeliverToMailboxAndForward).Trim()
        if (-not [string]::IsNullOrWhiteSpace($deliverAndForwardRaw)) {
            if ($supports.DeliverToMailboxAndForward) {
                $setParams.DeliverToMailboxAndForward = ConvertTo-Bool -Value $deliverAndForwardRaw
            }
            else {
                $warnings.Add('DeliverToMailboxAndForward ignored (unsupported parameter).')
            }
        }

        $auditEnabledRaw = ([string]$row.AuditEnabled).Trim()
        if (-not [string]::IsNullOrWhiteSpace($auditEnabledRaw)) {
            if ($supports.AuditEnabled) {
                $setParams.AuditEnabled = ConvertTo-Bool -Value $auditEnabledRaw
            }
            else {
                $warnings.Add('AuditEnabled ignored (unsupported parameter).')
            }
        }

        $litigationHoldRaw = ([string]$row.LitigationHoldEnabled).Trim()
        if (-not [string]::IsNullOrWhiteSpace($litigationHoldRaw)) {
            if ($supports.LitigationHoldEnabled) {
                $setParams.LitigationHoldEnabled = ConvertTo-Bool -Value $litigationHoldRaw
            }
            else {
                $warnings.Add('LitigationHoldEnabled ignored (unsupported parameter).')
            }
        }

        if ($setParams.Count -eq 1) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentity -Action 'UpdateSharedMailbox' -Status 'Skipped' -Message 'No updates specified.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($sharedMailboxIdentity, 'Update Exchange Online shared mailbox')) {
            Invoke-WithRetry -OperationName "Update shared mailbox $sharedMailboxIdentity" -ScriptBlock {
                Set-Mailbox @setParams -ErrorAction Stop
            }
            $message = 'Shared mailbox updated successfully.'
            if ($warnings.Count -gt 0) {
                $message = "$message $($warnings -join ' ')"
            }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentity -Action 'UpdateSharedMailbox' -Status 'Updated' -Message $message))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentity -Action 'UpdateSharedMailbox' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($sharedMailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentity -Action 'UpdateSharedMailbox' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online shared mailbox update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





