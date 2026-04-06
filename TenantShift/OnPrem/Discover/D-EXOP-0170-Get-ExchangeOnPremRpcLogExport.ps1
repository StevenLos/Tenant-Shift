<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260305-081900

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)

.SYNOPSIS
    Gets ExchangeOnPremRpcLogExport and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnPremRpcLogExport from Active Directory and writes the results to a CSV file.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all objects in scope (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all objects in scope rather than processing from an input CSV file. Uses the DiscoverAll parameter set.

.PARAMETER LogPath
    Path to an optional log file for detailed diagnostic output.

.PARAMETER LookbackDays
    Number of days to look back when evaluating activity or sign-in data. Defaults to 30.

.PARAMETER MaxObjects
    Maximum number of objects to retrieve. 0 (default) means no limit.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-D0226-Get-ExchangeOnPremRpcLogExport.ps1 -InputCsvPath .\0226.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D0226-Get-ExchangeOnPremRpcLogExport.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Exchange Management Shell cmdlets (session)
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$LogPath = 'C:\Program Files\Microsoft\Exchange Server\V15\Logging\RPC Client Access',

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [ValidateRange(1, 3650)]
    [int]$LookbackDays = 15,

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [ValidateRange(0, 10000000)]
    [int]$MaxObjects = 0,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_SM-D0226-Get-ExchangeOnPremRpcLogExport_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-RpcLogHeaders {
    [CmdletBinding()]
    param()

    return @(
        'date-time',
        'session-id',
        'seq-number',
        'client-name',
        'organization-info',
        'client-software',
        'client-software-version',
        'client-mode',
        'client-ip',
        'server-ip',
        'protocol',
        'application-id',
        'operation',
        'rpc-status',
        'processing-time',
        'operation-specific',
        'failures'
    )
}

function Resolve-RpcLogFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RpcLogPath,

        [Parameter(Mandatory)]
        [int]$LookbackDays
    )

    if (-not (Test-Path -LiteralPath $RpcLogPath -PathType Container)) {
        throw "RPC log path '$RpcLogPath' was not found."
    }

    $cutoff = (Get-Date).AddDays(-1 * $LookbackDays)
    $files = @(
        Get-ChildItem -LiteralPath $RpcLogPath -File -ErrorAction Stop |
            Where-Object { $_.LastWriteTime -ge $cutoff } |
            Sort-Object -Property LastWriteTime
    )

    return $files
}

function Parse-RpcEntries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory)]
        [string[]]$Headers
    )

    $lines = @(
        Get-Content -LiteralPath $File.FullName -ErrorAction Stop |
            Where-Object {
                $line = [string]$_
                (-not [string]::IsNullOrWhiteSpace($line)) -and (-not $line.TrimStart().StartsWith('#'))
            }
    )

    if ($lines.Count -eq 0) {
        return @()
    }

    return @($lines | ConvertFrom-Csv -Header $Headers)
}

function New-InventoryResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$RowNumber,

        [Parameter(Mandatory)]
        [string]$PrimaryKey,

        [Parameter(Mandatory)]
        [string]$Action,

        [Parameter(Mandatory)]
        [string]$Status,

        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory)]
        [hashtable]$Data
    )

    $base = New-ResultObject -RowNumber $RowNumber -PrimaryKey $PrimaryKey -Action $Action -Status $Status -Message $Message
    $ordered = [ordered]@{}

    foreach ($prop in $base.PSObject.Properties.Name) {
        $ordered[$prop] = $base.$prop
    }

    foreach ($key in $Data.Keys) {
        $ordered[$key] = $Data[$key]
    }

    return [PSCustomObject]$ordered
}

$requiredHeaders = @(
    'RpcLogPath',
    'LookbackDays'
)

Write-Status -Message 'Starting Exchange on-prem RPC log export inventory script.'
Ensure-ExchangeOnPremConnection

$scopeMode = 'Csv'
$resolvedDefaultLogPath = ''
$resolvedDefaultLookbackDays = ''
$runWasTruncated = $false

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    $resolvedDefaultLogPath = Get-TrimmedValue -Value $LogPath
    $resolvedDefaultLookbackDays = [string]$LookbackDays

    Write-Status -Message "DiscoverAll enabled for RPC log export. LogPath='$resolvedDefaultLogPath'; LookbackDays='$resolvedDefaultLookbackDays'." -Level WARN

    $rows = @([PSCustomObject]@{
            RpcLogPath   = $resolvedDefaultLogPath
            LookbackDays = $resolvedDefaultLookbackDays
        })
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$headers = Get-RpcLogHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $rpcLogPath = Get-TrimmedValue -Value $row.RpcLogPath

    try {
        if ([string]::IsNullOrWhiteSpace($rpcLogPath)) {
            throw 'RpcLogPath is required.'
        }

        $lookbackRaw = Get-TrimmedValue -Value $row.LookbackDays
        $effectiveLookbackDays = if ([string]::IsNullOrWhiteSpace($lookbackRaw)) { 15 } else { [int]$lookbackRaw }
        if ($effectiveLookbackDays -lt 1) {
            throw 'LookbackDays must be a positive integer.'
        }

        $files = @(Resolve-RpcLogFiles -RpcLogPath $rpcLogPath -LookbackDays $effectiveLookbackDays)
        if ($files.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $rpcLogPath -Action 'GetExchangeRpcLogExport' -Status 'Completed' -Message 'No RPC log files matched the requested lookback window.' -Data ([ordered]@{
                        RpcLogPath          = $rpcLogPath
                        LookbackDays        = [string]$effectiveLookbackDays
                        SourceFile          = ''
                        SourceFileLastWriteUtc = ''
                        DateTime            = ''
                        SessionId           = ''
                        SequenceNumber      = ''
                        ClientName          = ''
                        OrganizationInfo    = ''
                        ClientSoftware      = ''
                        ClientSoftwareVersion = ''
                        ClientMode          = ''
                        ClientIp            = ''
                        ServerIp            = ''
                        Protocol            = ''
                        ApplicationId       = ''
                        Operation           = ''
                        RpcStatus           = ''
                        ProcessingTime      = ''
                        OperationSpecific   = ''
                        Failures            = ''
                        FilesScanned        = '0'
                        RowsParsed          = '0'
                        RowsExported        = '0'
                    })))
            $rowNumber++
            continue
        }

        $rowsParsed = 0
        $rowsExported = 0
        $maxReached = $false

        foreach ($file in $files) {
            $entries = @(Parse-RpcEntries -File $file -Headers $headers)
            $rowsParsed += $entries.Count

            foreach ($entry in $entries) {
                if ($MaxObjects -gt 0 -and $rowsExported -ge $MaxObjects) {
                    $maxReached = $true
                    break
                }

                $rowsExported++
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$rpcLogPath|$($file.Name)|$rowsExported" -Action 'GetExchangeRpcLogExport' -Status 'Completed' -Message 'RPC log row exported.' -Data ([ordered]@{
                            RpcLogPath          = $rpcLogPath
                            LookbackDays        = [string]$effectiveLookbackDays
                            SourceFile          = $file.FullName
                            SourceFileLastWriteUtc = $file.LastWriteTimeUtc.ToString('o')
                            DateTime            = Get-TrimmedValue -Value $entry.'date-time'
                            SessionId           = Get-TrimmedValue -Value $entry.'session-id'
                            SequenceNumber      = Get-TrimmedValue -Value $entry.'seq-number'
                            ClientName          = Get-TrimmedValue -Value $entry.'client-name'
                            OrganizationInfo    = Get-TrimmedValue -Value $entry.'organization-info'
                            ClientSoftware      = Get-TrimmedValue -Value $entry.'client-software'
                            ClientSoftwareVersion = Get-TrimmedValue -Value $entry.'client-software-version'
                            ClientMode          = Get-TrimmedValue -Value $entry.'client-mode'
                            ClientIp            = Get-TrimmedValue -Value $entry.'client-ip'
                            ServerIp            = Get-TrimmedValue -Value $entry.'server-ip'
                            Protocol            = Get-TrimmedValue -Value $entry.'protocol'
                            ApplicationId       = Get-TrimmedValue -Value $entry.'application-id'
                            Operation           = Get-TrimmedValue -Value $entry.'operation'
                            RpcStatus           = Get-TrimmedValue -Value $entry.'rpc-status'
                            ProcessingTime      = Get-TrimmedValue -Value $entry.'processing-time'
                            OperationSpecific   = Get-TrimmedValue -Value $entry.'operation-specific'
                            Failures            = Get-TrimmedValue -Value $entry.'failures'
                            FilesScanned        = [string]$files.Count
                            RowsParsed          = [string]$rowsParsed
                            RowsExported        = [string]$rowsExported
                        })))
            }

            if ($maxReached) {
                $runWasTruncated = $true
                break
            }
        }

        if ($rowsExported -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $rpcLogPath -Action 'GetExchangeRpcLogExport' -Status 'Completed' -Message 'RPC log files were found, but no parsable rows were present.' -Data ([ordered]@{
                        RpcLogPath          = $rpcLogPath
                        LookbackDays        = [string]$effectiveLookbackDays
                        SourceFile          = ''
                        SourceFileLastWriteUtc = ''
                        DateTime            = ''
                        SessionId           = ''
                        SequenceNumber      = ''
                        ClientName          = ''
                        OrganizationInfo    = ''
                        ClientSoftware      = ''
                        ClientSoftwareVersion = ''
                        ClientMode          = ''
                        ClientIp            = ''
                        ServerIp            = ''
                        Protocol            = ''
                        ApplicationId       = ''
                        Operation           = ''
                        RpcStatus           = ''
                        ProcessingTime      = ''
                        OperationSpecific   = ''
                        Failures            = ''
                        FilesScanned        = [string]$files.Count
                        RowsParsed          = [string]$rowsParsed
                        RowsExported        = '0'
                    })))
        }
        elseif ($maxReached) {
            Write-Status -Message "Row $rowNumber export for '$rpcLogPath' reached MaxObjects limit of $MaxObjects." -Level WARN
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($rpcLogPath) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $rpcLogPath -Action 'GetExchangeRpcLogExport' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    RpcLogPath          = $rpcLogPath
                    LookbackDays        = Get-TrimmedValue -Value $row.LookbackDays
                    SourceFile          = ''
                    SourceFileLastWriteUtc = ''
                    DateTime            = ''
                    SessionId           = ''
                    SequenceNumber      = ''
                    ClientName          = ''
                    OrganizationInfo    = ''
                    ClientSoftware      = ''
                    ClientSoftwareVersion = ''
                    ClientMode          = ''
                    ClientIp            = ''
                    ServerIp            = ''
                    Protocol            = ''
                    ApplicationId       = ''
                    Operation           = ''
                    RpcStatus           = ''
                    ProcessingTime      = ''
                    OperationSpecific   = ''
                    Failures            = ''
                    FilesScanned        = '0'
                    RowsParsed          = '0'
                    RowsExported        = '0'
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeDefaultLogPath' -NotePropertyValue $resolvedDefaultLogPath -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeDefaultLookbackDays' -NotePropertyValue $resolvedDefaultLookbackDays -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeMaxObjects' -NotePropertyValue ($(if ($scopeMode -eq 'DiscoverAll') { [string]$MaxObjects } else { '' })) -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeWasTruncated' -NotePropertyValue ([string]$runWasTruncated) -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem RPC log export inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
