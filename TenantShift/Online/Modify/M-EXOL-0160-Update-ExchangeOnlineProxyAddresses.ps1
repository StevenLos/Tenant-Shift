<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-151500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Modifies ExchangeOnlineProxyAddresses in Microsoft 365.

.DESCRIPTION
    Updates ExchangeOnlineProxyAddresses in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M3125-Update-ExchangeOnlineProxyAddresses.ps1 -InputCsvPath .\3125.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M3125-Update-ExchangeOnlineProxyAddresses.ps1 -InputCsvPath .\3125.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator
    Limitations:      None known.

    CSV Fields:
    Column                        Type      Required  Description
    ----------------------------  ----      --------  -----------
    MailboxIdentity               String    Yes       <fill in description>
    PrimarySmtpAddress            String    Yes       <fill in description>
    AddProxyAddresses             String    Yes       <fill in description>
    RemoveProxyAddresses          String    Yes       <fill in description>
    ReplaceAllProxyAddresses      String    Yes       <fill in description>
    ClearSecondaryProxyAddresses  String    Yes       <fill in description>
    Notes                         String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3125-Update-ExchangeOnlineProxyAddresses_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function ConvertTo-NormalizedSmtpAddress {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    $trimmed = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return ''
    }

    if ($trimmed.StartsWith('SMTP:', [System.StringComparison]::OrdinalIgnoreCase)) {
        $trimmed = $trimmed.Substring(5)
    }

    return $trimmed.ToLowerInvariant()
}

function ConvertTo-ProxyAddressArray {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    $items = ConvertTo-Array -Value $Value
    $deduped = [System.Collections.Generic.List[string]]::new()

    foreach ($item in $items) {
        $trimmed = Get-TrimmedValue -Value $item
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }

        if ($trimmed -notmatch ':') {
            if ($trimmed -match '@') {
                $trimmed = "smtp:$trimmed"
            }
        }

        if (-not ($deduped.Contains($trimmed))) {
            $deduped.Add($trimmed)
        }
    }

    return $deduped.ToArray()
}

function ConvertTo-CanonicalProxyAddressSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PrimarySmtpAddress,

        [Parameter(Mandatory)]
        [string[]]$InputAddresses
    )

    $normalizedPrimary = ConvertTo-NormalizedSmtpAddress -Value $PrimarySmtpAddress
    if ([string]::IsNullOrWhiteSpace($normalizedPrimary)) {
        throw 'Primary SMTP address cannot be empty when building canonical proxy address set.'
    }

    $finalList = [System.Collections.Generic.List[string]]::new()
    $finalList.Add("SMTP:$normalizedPrimary")

    foreach ($entry in @($InputAddresses)) {
        $trimmedEntry = Get-TrimmedValue -Value $entry
        if ([string]::IsNullOrWhiteSpace($trimmedEntry)) {
            continue
        }

        $candidate = ''
        if ($trimmedEntry -match '^(?<prefix>[^:]+):(?<value>.+)$') {
            $prefix = $matches['prefix']
            $value = Get-TrimmedValue -Value $matches['value']

            if ($prefix.Equals('smtp', [System.StringComparison]::OrdinalIgnoreCase)) {
                $smtp = ConvertTo-NormalizedSmtpAddress -Value $value
                if ($smtp -eq $normalizedPrimary) {
                    continue
                }

                $candidate = "smtp:$smtp"
            }
            else {
                $candidate = "{0}:{1}" -f $prefix, $value
            }
        }
        else {
            $smtp = ConvertTo-NormalizedSmtpAddress -Value $trimmedEntry
            if ([string]::IsNullOrWhiteSpace($smtp) -or $smtp -eq $normalizedPrimary) {
                continue
            }

            $candidate = "smtp:$smtp"
        }

        if (-not [string]::IsNullOrWhiteSpace($candidate) -and -not ($finalList.Contains($candidate))) {
            $finalList.Add($candidate)
        }
    }

    return $finalList.ToArray()
}

$requiredHeaders = @(
    'MailboxIdentity',
    'PrimarySmtpAddress',
    'AddProxyAddresses',
    'RemoveProxyAddresses',
    'ReplaceAllProxyAddresses',
    'ClearSecondaryProxyAddresses',
    'Notes'
)

Write-Status -Message 'Starting Exchange Online proxy address update script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = Get-TrimmedValue -Value $row.MailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
            throw 'MailboxIdentity is required.'
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentity" -ScriptBlock {
            Get-ExchangeOnlineMailbox -Identity $mailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'UpdateMailboxProxyAddresses' -Status 'NotFound' -Message 'Mailbox not found.'))
            $rowNumber++
            continue
        }

        $primarySmtpAddress = Get-TrimmedValue -Value $row.PrimarySmtpAddress
        $addAddresses = ConvertTo-ProxyAddressArray -Value ([string]$row.AddProxyAddresses)
        $removeAddresses = ConvertTo-ProxyAddressArray -Value ([string]$row.RemoveProxyAddresses)
        $replaceAddresses = ConvertTo-ProxyAddressArray -Value ([string]$row.ReplaceAllProxyAddresses)

        $clearSecondaryRaw = Get-TrimmedValue -Value $row.ClearSecondaryProxyAddresses
        $clearSecondary = $false
        if (-not [string]::IsNullOrWhiteSpace($clearSecondaryRaw)) {
            $clearSecondary = ConvertTo-Bool -Value $clearSecondaryRaw
        }

        if ($replaceAddresses.Count -gt 0 -and ($addAddresses.Count -gt 0 -or $removeAddresses.Count -gt 0 -or $clearSecondary)) {
            throw 'ReplaceAllProxyAddresses cannot be combined with AddProxyAddresses, RemoveProxyAddresses, or ClearSecondaryProxyAddresses.'
        }

        if ($clearSecondary -and $removeAddresses.Count -gt 0) {
            throw 'ClearSecondaryProxyAddresses cannot be combined with RemoveProxyAddresses.'
        }

        $requestedChange = (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) -or ($addAddresses.Count -gt 0) -or ($removeAddresses.Count -gt 0) -or ($replaceAddresses.Count -gt 0) -or $clearSecondary
        if (-not $requestedChange) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'UpdateMailboxProxyAddresses' -Status 'Skipped' -Message 'No proxy address updates were requested.'))
            $rowNumber++
            continue
        }

        $setParams = @{
            Identity = $mailbox.Identity
        }

        if ($replaceAddresses.Count -gt 0) {
            $targetPrimary = ConvertTo-NormalizedSmtpAddress -Value $primarySmtpAddress
            if ([string]::IsNullOrWhiteSpace($targetPrimary)) {
                $primaryInList = @($replaceAddresses | Where-Object { $_.StartsWith('SMTP:', [System.StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1)
                if ($primaryInList.Count -gt 0) {
                    $targetPrimary = ConvertTo-NormalizedSmtpAddress -Value $primaryInList[0]
                }
            }
            if ([string]::IsNullOrWhiteSpace($targetPrimary) -and $replaceAddresses.Count -gt 0) {
                $targetPrimary = ConvertTo-NormalizedSmtpAddress -Value $replaceAddresses[0]
            }
            if ([string]::IsNullOrWhiteSpace($targetPrimary)) {
                $targetPrimary = ConvertTo-NormalizedSmtpAddress -Value $mailbox.PrimarySmtpAddress
            }

            $setParams.EmailAddresses = ConvertTo-CanonicalProxyAddressSet -PrimarySmtpAddress $targetPrimary -InputAddresses $replaceAddresses

            if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
                $setParams.PrimarySmtpAddress = $primarySmtpAddress
            }
        }
        elseif ($clearSecondary) {
            $targetPrimary = if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) { $primarySmtpAddress } else { ([string]$mailbox.PrimarySmtpAddress).Trim() }
            if ([string]::IsNullOrWhiteSpace($targetPrimary)) {
                throw 'Unable to determine primary SMTP address while clearing secondary proxy addresses.'
            }

            $setParams.EmailAddresses = ConvertTo-CanonicalProxyAddressSet -PrimarySmtpAddress $targetPrimary -InputAddresses $addAddresses

            if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
                $setParams.PrimarySmtpAddress = $primarySmtpAddress
            }
        }
        else {
            if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
                $setParams.PrimarySmtpAddress = $primarySmtpAddress
            }

            $emailAddressOps = @{}
            if ($addAddresses.Count -gt 0) {
                $emailAddressOps['Add'] = $addAddresses
            }
            if ($removeAddresses.Count -gt 0) {
                $emailAddressOps['Remove'] = $removeAddresses
            }

            if ($emailAddressOps.Count -gt 0) {
                $setParams.EmailAddresses = $emailAddressOps
            }
        }

        if ($setParams.Count -eq 1) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'UpdateMailboxProxyAddresses' -Status 'Skipped' -Message 'No effective proxy address updates were generated.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($mailboxIdentity, 'Update mailbox proxy addresses')) {
            Invoke-WithRetry -OperationName "Update mailbox proxy addresses $mailboxIdentity" -ScriptBlock {
                Set-Mailbox @setParams -ErrorAction Stop
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'UpdateMailboxProxyAddresses' -Status 'Updated' -Message 'Mailbox proxy addresses updated successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'UpdateMailboxProxyAddresses' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'UpdateMailboxProxyAddresses' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online proxy address update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
