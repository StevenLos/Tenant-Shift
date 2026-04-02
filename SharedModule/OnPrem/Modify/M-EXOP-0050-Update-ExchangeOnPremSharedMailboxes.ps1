<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-220000

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)

.SYNOPSIS
    Modifies ExchangeOnPremSharedMailboxes in Active Directory.

.DESCRIPTION
    Updates ExchangeOnPremSharedMailboxes in Active Directory based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M0216-Update-ExchangeOnPremSharedMailboxes.ps1 -InputCsvPath .\0216.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M0216-Update-ExchangeOnPremSharedMailboxes.ps1 -InputCsvPath .\0216.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Exchange Management Shell cmdlets (session)
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    Column                             Type      Required  Description
    ---------------------------------  ----      --------  -----------
    SharedMailboxIdentity              String    Yes       <fill in description>
    DisplayName                        String    Yes       <fill in description>
    PrimarySmtpAddress                 String    Yes       <fill in description>
    HiddenFromAddressListsEnabled      String    Yes       <fill in description>
    MessageCopyForSentAsEnabled        String    Yes       <fill in description>
    MessageCopyForSendOnBehalfEnabled  String    Yes       <fill in description>
    MailTip                            String    Yes       <fill in description>
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M0216-Update-ExchangeOnPremSharedMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'SharedMailboxIdentity',
    'DisplayName',
    'PrimarySmtpAddress',
    'HiddenFromAddressListsEnabled',
    'MessageCopyForSentAsEnabled',
    'MessageCopyForSendOnBehalfEnabled',
    'MailTip'
)

Write-Status -Message 'Starting Exchange on-prem shared mailbox update script.'
Ensure-ExchangeOnPremConnection

$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop
$supports = @{
    HiddenFromAddressListsEnabled       = $setMailboxCommand.Parameters.ContainsKey('HiddenFromAddressListsEnabled')
    MessageCopyForSentAsEnabled         = $setMailboxCommand.Parameters.ContainsKey('MessageCopyForSentAsEnabled')
    MessageCopyForSendOnBehalfEnabled   = $setMailboxCommand.Parameters.ContainsKey('MessageCopyForSendOnBehalfEnabled')
    MailTip                             = $setMailboxCommand.Parameters.ContainsKey('MailTip')
}

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $sharedMailboxIdentity = Get-TrimmedValue -Value $row.SharedMailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($sharedMailboxIdentity)) {
            throw 'SharedMailboxIdentity is required.'
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup shared mailbox $sharedMailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $sharedMailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentity -Action 'UpdateSharedMailbox' -Status 'NotFound' -Message 'Shared mailbox not found.'))
            $rowNumber++
            continue
        }

        if ((Get-TrimmedValue -Value $mailbox.RecipientTypeDetails) -ne 'SharedMailbox') {
            throw "Recipient '$sharedMailboxIdentity' is '$($mailbox.RecipientTypeDetails)'. Expected SharedMailbox."
        }

        $setParams = @{ Identity = $mailbox.Identity }
        $warnings = [System.Collections.Generic.List[string]]::new()

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $setParams.DisplayName = $displayName
        }

        $primarySmtpAddress = Get-TrimmedValue -Value $row.PrimarySmtpAddress
        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $setParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $hiddenRaw = Get-TrimmedValue -Value $row.HiddenFromAddressListsEnabled
        if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
            if ($supports.HiddenFromAddressListsEnabled) {
                $setParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
            }
            else {
                $warnings.Add('HiddenFromAddressListsEnabled ignored (unsupported parameter).')
            }
        }

        $copySentAsRaw = Get-TrimmedValue -Value $row.MessageCopyForSentAsEnabled
        if (-not [string]::IsNullOrWhiteSpace($copySentAsRaw)) {
            if ($supports.MessageCopyForSentAsEnabled) {
                $setParams.MessageCopyForSentAsEnabled = ConvertTo-Bool -Value $copySentAsRaw
            }
            else {
                $warnings.Add('MessageCopyForSentAsEnabled ignored (unsupported parameter).')
            }
        }

        $copySobRaw = Get-TrimmedValue -Value $row.MessageCopyForSendOnBehalfEnabled
        if (-not [string]::IsNullOrWhiteSpace($copySobRaw)) {
            if ($supports.MessageCopyForSendOnBehalfEnabled) {
                $setParams.MessageCopyForSendOnBehalfEnabled = ConvertTo-Bool -Value $copySobRaw
            }
            else {
                $warnings.Add('MessageCopyForSendOnBehalfEnabled ignored (unsupported parameter).')
            }
        }

        $mailTip = Get-TrimmedValue -Value $row.MailTip
        if (-not [string]::IsNullOrWhiteSpace($mailTip)) {
            if ($supports.MailTip) {
                $setParams.MailTip = $mailTip
            }
            else {
                $warnings.Add('MailTip ignored (unsupported parameter).')
            }
        }

        if ($setParams.Count -eq 1) {
            $message = 'No updates specified.'
            if ($warnings.Count -gt 0) {
                $message = "$message $($warnings -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentity -Action 'UpdateSharedMailbox' -Status 'Skipped' -Message $message))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($sharedMailboxIdentity, 'Update Exchange on-prem shared mailbox')) {
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
Write-Status -Message 'Exchange on-prem shared mailbox update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
