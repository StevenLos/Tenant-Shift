<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-170000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Modifies ExchangeOnlineSafeSenderDomains in Microsoft 365.

.DESCRIPTION
    Updates ExchangeOnlineSafeSenderDomains in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M3128-Set-ExchangeOnlineSafeSenderDomains.ps1 -InputCsvPath .\3128.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M3128-Set-ExchangeOnlineSafeSenderDomains.ps1 -InputCsvPath .\3128.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator
    Limitations:      None known.

    CSV Fields:
    Column                       Type      Required  Description
    ---------------------------  ----      --------  -----------
    MailboxIdentity              String    Yes       <fill in description>
    AddSafeSenderDomains         String    Yes       <fill in description>
    RemoveSafeSenderDomains      String    Yes       <fill in description>
    ReplaceAllSafeSenderDomains  String    Yes       <fill in description>
    ClearSafeSenderDomains       String    Yes       <fill in description>
    Notes                        String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3128-Set-ExchangeOnlineSafeSenderDomains_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function ConvertTo-NormalizedSafeSenderEntries {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    $entries = [System.Collections.Generic.List[string]]::new()

    if ($null -eq $Value) {
        return $entries.ToArray()
    }

    $rawItems = @()
    if ($Value -is [string]) {
        $rawItems = ConvertTo-Array -Value ([string]$Value)
    }
    elseif ($Value -is [System.Collections.IEnumerable]) {
        $rawItems = @($Value)
    }
    else {
        $rawItems = @([string]$Value)
    }

    $dedupe = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($raw in $rawItems) {
        $text = Get-TrimmedValue -Value $raw
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }

        $normalized = $text.ToLowerInvariant()
        if ($dedupe.Add($normalized)) {
            $entries.Add($normalized)
        }
    }

    return $entries.ToArray()
}

function Test-StringSetEqual {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Left,

        [Parameter(Mandatory)]
        [string[]]$Right
    )

    $leftSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $Left) {
        [void]$leftSet.Add($item)
    }

    $rightSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $Right) {
        [void]$rightSet.Add($item)
    }

    return $leftSet.SetEquals($rightSet)
}

$requiredHeaders = @(
    'MailboxIdentity',
    'AddSafeSenderDomains',
    'RemoveSafeSenderDomains',
    'ReplaceAllSafeSenderDomains',
    'ClearSafeSenderDomains',
    'Notes'
)

Write-Status -Message 'Starting Exchange Online safe sender-domain update script.'
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
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetSafeSenderDomains' -Status 'NotFound' -Message 'Mailbox not found.'))
            $rowNumber++
            continue
        }

        $addDomains = ConvertTo-NormalizedSafeSenderEntries -Value $row.AddSafeSenderDomains
        $removeDomains = ConvertTo-NormalizedSafeSenderEntries -Value $row.RemoveSafeSenderDomains
        $replaceAllDomains = ConvertTo-NormalizedSafeSenderEntries -Value $row.ReplaceAllSafeSenderDomains

        $clearRaw = Get-TrimmedValue -Value $row.ClearSafeSenderDomains
        $clearSafeSenderDomains = $false
        if (-not [string]::IsNullOrWhiteSpace($clearRaw)) {
            $clearSafeSenderDomains = ConvertTo-Bool -Value $clearRaw
        }

        if ($replaceAllDomains.Count -gt 0 -and ($addDomains.Count -gt 0 -or $removeDomains.Count -gt 0 -or $clearSafeSenderDomains)) {
            throw 'ReplaceAllSafeSenderDomains cannot be combined with AddSafeSenderDomains, RemoveSafeSenderDomains, or ClearSafeSenderDomains.'
        }

        $requestedChange = ($addDomains.Count -gt 0) -or ($removeDomains.Count -gt 0) -or ($replaceAllDomains.Count -gt 0) -or $clearSafeSenderDomains
        if (-not $requestedChange) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetSafeSenderDomains' -Status 'Skipped' -Message 'No safe sender-domain updates were requested.'))
            $rowNumber++
            continue
        }

        $junkConfig = Invoke-WithRetry -OperationName "Load junk email configuration for $mailboxIdentity" -ScriptBlock {
            Get-MailboxJunkEmailConfiguration -Identity $mailbox.Identity -ErrorAction Stop
        }

        $currentDomains = ConvertTo-NormalizedSafeSenderEntries -Value $junkConfig.TrustedSendersAndDomains

        $desiredDomains = @()
        if ($replaceAllDomains.Count -gt 0) {
            $desiredDomains = $replaceAllDomains
        }
        elseif ($clearSafeSenderDomains) {
            $desiredDomains = $addDomains
        }
        else {
            $workingSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($entry in $currentDomains) {
                [void]$workingSet.Add($entry)
            }

            foreach ($entry in $addDomains) {
                [void]$workingSet.Add($entry)
            }

            foreach ($entry in $removeDomains) {
                [void]$workingSet.Remove($entry)
            }

            $desiredDomains = @($workingSet | Sort-Object)
        }

        if (Test-StringSetEqual -Left $currentDomains -Right $desiredDomains) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetSafeSenderDomains' -Status 'Skipped' -Message 'Safe sender domains already match requested state.'))
            $rowNumber++
            continue
        }

        $setParams = @{
            Identity                 = $mailbox.Identity
            TrustedSendersAndDomains = @($desiredDomains)
        }

        $message = "Safe sender domains updated from $($currentDomains.Count) to $($desiredDomains.Count) entries."

        if ($PSCmdlet.ShouldProcess($mailboxIdentity, 'Update mailbox safe sender domains')) {
            Invoke-WithRetry -OperationName "Update safe sender domains for $mailboxIdentity" -ScriptBlock {
                Set-MailboxJunkEmailConfiguration @setParams -ErrorAction Stop
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetSafeSenderDomains' -Status 'Updated' -Message $message))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetSafeSenderDomains' -Status 'WhatIf' -Message "Update skipped due to WhatIf. $message"))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetSafeSenderDomains' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online safe sender-domain update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
