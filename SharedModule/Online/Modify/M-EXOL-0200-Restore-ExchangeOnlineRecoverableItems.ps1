<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-170500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Modifies ExchangeOnlineRecoverableItems in Microsoft 365.

.DESCRIPTION
    Updates ExchangeOnlineRecoverableItems in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M3129-Restore-ExchangeOnlineRecoverableItems.ps1 -InputCsvPath .\3129.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M3129-Restore-ExchangeOnlineRecoverableItems.ps1 -InputCsvPath .\3129.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator
    Limitations:      None known.

    CSV Fields:
    Column                Type      Required  Description
    --------------------  ----      --------  -----------
    MailboxIdentity       String    Yes       <fill in description>
    SourceFolder          String    Yes       <fill in description>
    ResultSize            String    Yes       <fill in description>
    MaxParallelSize       String    Yes       <fill in description>
    PreviewOnly           String    Yes       <fill in description>
    Notes                 String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3129-Restore-ExchangeOnlineRecoverableItems_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-NullableInt {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $parsed = 0
    if (-not [int]::TryParse($text, [ref]$parsed)) {
        throw "Value '$text' is not a valid integer."
    }

    return $parsed
}

function Get-DeletionsSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity
    )

    $folderStats = @(Invoke-WithRetry -OperationName "Load RecoverableItems folder statistics for $Identity" -ScriptBlock {
        Get-ExchangeOnlineMailboxFolderStatistics -Identity $Identity -FolderScope RecoverableItems -ErrorAction Stop
    })

    $deletions = @($folderStats | Where-Object { ([string]$_.Name).Trim() -eq 'Deletions' } | Select-Object -First 1)

    if ($deletions.Count -eq 0) {
        return [PSCustomObject]@{
            RemainingItems = 0
            RemainingSize  = ''
        }
    }

    return [PSCustomObject]@{
        RemainingItems = [int]$deletions[0].ItemsInFolderAndSubfolders
        RemainingSize  = ([string]$deletions[0].FolderAndSubfolderSize).Trim()
    }
}

$requiredHeaders = @(
    'MailboxIdentity',
    'SourceFolder',
    'ResultSize',
    'MaxParallelSize',
    'PreviewOnly',
    'Notes'
)

Write-Status -Message 'Starting Exchange Online recoverable-item restore script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$restoreCommand = Get-Command -Name Restore-RecoverableItems -ErrorAction SilentlyContinue
if (-not $restoreCommand) {
    throw 'Restore-RecoverableItems is not available in this Exchange Online session.'
}

$supportsSourceFolder = $restoreCommand.Parameters.ContainsKey('SourceFolder')
$supportsResultSize = $restoreCommand.Parameters.ContainsKey('ResultSize')
$supportsMaxParallelSize = $restoreCommand.Parameters.ContainsKey('MaxParallelSize')
$supportsNoOutput = $restoreCommand.Parameters.ContainsKey('NoOutput')

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
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'RestoreRecoverableItems' -Status 'NotFound' -Message 'Mailbox not found.'))
            $rowNumber++
            continue
        }

        $sourceFolder = Get-TrimmedValue -Value $row.SourceFolder
        if ([string]::IsNullOrWhiteSpace($sourceFolder)) {
            $sourceFolder = 'RecoverableItems'
        }

        $resultSize = Get-NullableInt -Value $row.ResultSize
        if ($null -eq $resultSize) {
            $resultSize = 1000
        }

        $maxParallelSize = Get-NullableInt -Value $row.MaxParallelSize
        if ($null -eq $maxParallelSize) {
            $maxParallelSize = 1
        }

        if ($resultSize -le 0) {
            throw 'ResultSize must be greater than zero.'
        }

        if ($maxParallelSize -le 0) {
            throw 'MaxParallelSize must be greater than zero.'
        }

        $previewOnly = ConvertTo-Bool -Value (Get-TrimmedValue -Value $row.PreviewOnly)

        $warnings = [System.Collections.Generic.List[string]]::new()
        $setParams = @{
            Identity = @($mailbox.Identity)
        }

        if ($supportsSourceFolder) {
            $setParams['SourceFolder'] = $sourceFolder
        }
        else {
            $warnings.Add('SourceFolder was ignored because Restore-RecoverableItems does not support -SourceFolder in this session.')
        }

        if ($supportsResultSize) {
            $setParams['ResultSize'] = $resultSize
        }
        else {
            $warnings.Add('ResultSize was ignored because Restore-RecoverableItems does not support -ResultSize in this session.')
        }

        if ($supportsMaxParallelSize) {
            $setParams['MaxParallelSize'] = $maxParallelSize
        }
        else {
            $warnings.Add('MaxParallelSize was ignored because Restore-RecoverableItems does not support -MaxParallelSize in this session.')
        }

        if ($supportsNoOutput) {
            $setParams['NoOutput'] = $true
        }

        $before = Get-DeletionsSnapshot -Identity $mailbox.Identity

        if ($previewOnly) {
            $message = "PreviewOnly is TRUE. Current RecoverableItems/Deletions count is $($before.RemainingItems)."
            if ($warnings.Count -gt 0) {
                $message = "$message Warnings: $($warnings -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'RestoreRecoverableItems' -Status 'Preview' -Message $message))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($mailboxIdentity, 'Restore recoverable items')) {
            Invoke-WithRetry -OperationName "Restore recoverable items for $mailboxIdentity" -ScriptBlock {
                Restore-RecoverableItems @setParams -ErrorAction Stop | Out-Null
            }

            $after = Get-DeletionsSnapshot -Identity $mailbox.Identity
            $estimatedRestoredItems = $before.RemainingItems - $after.RemainingItems
            if ($estimatedRestoredItems -lt 0) {
                $estimatedRestoredItems = 0
            }

            $message = "Restore request submitted. Deletions before: $($before.RemainingItems); after: $($after.RemainingItems); estimated restored this run: $estimatedRestoredItems."
            if ($warnings.Count -gt 0) {
                $message = "$message Warnings: $($warnings -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'RestoreRecoverableItems' -Status 'Updated' -Message $message))
        }
        else {
            $message = 'Restore skipped due to WhatIf.'
            if ($warnings.Count -gt 0) {
                $message = "$message Warnings: $($warnings -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'RestoreRecoverableItems' -Status 'WhatIf' -Message $message))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'RestoreRecoverableItems' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online recoverable-item restore script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
