<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260305-081800

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)
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
    [string]$ClientSoftware = 'OUTLOOK.EXE',

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [ValidateRange(0, 10000000)]
    [int]$MaxObjects = 0,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR0225-Get-ExchangeOnPremOutlookClientVersionsFromRpcLogs_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

    $parsed = @($lines | ConvertFrom-Csv -Header $Headers)
    if ($parsed.Count -eq 0) {
        return @()
    }

    $entries = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $parsed) {
        $entries.Add([PSCustomObject]@{
                SourceFile            = $File.FullName
                SourceFileLastWriteUtc = $File.LastWriteTimeUtc.ToString('o')
                DateTime              = Get-TrimmedValue -Value $entry.'date-time'
                SessionId             = Get-TrimmedValue -Value $entry.'session-id'
                SequenceNumber        = Get-TrimmedValue -Value $entry.'seq-number'
                ClientName            = Get-TrimmedValue -Value $entry.'client-name'
                OrganizationInfo      = Get-TrimmedValue -Value $entry.'organization-info'
                ClientSoftware        = Get-TrimmedValue -Value $entry.'client-software'
                ClientSoftwareVersion = Get-TrimmedValue -Value $entry.'client-software-version'
                ClientMode            = Get-TrimmedValue -Value $entry.'client-mode'
                ClientIp              = Get-TrimmedValue -Value $entry.'client-ip'
                ServerIp              = Get-TrimmedValue -Value $entry.'server-ip'
                Protocol              = Get-TrimmedValue -Value $entry.'protocol'
                ApplicationId         = Get-TrimmedValue -Value $entry.'application-id'
                Operation             = Get-TrimmedValue -Value $entry.'operation'
                RpcStatus             = Get-TrimmedValue -Value $entry.'rpc-status'
                ProcessingTime        = Get-TrimmedValue -Value $entry.'processing-time'
                OperationSpecific     = Get-TrimmedValue -Value $entry.'operation-specific'
                Failures              = Get-TrimmedValue -Value $entry.'failures'
            })
    }

    return $entries.ToArray()
}

$requiredHeaders = @(
    'RpcLogPath',
    'LookbackDays',
    'ClientSoftware'
)

Write-Status -Message 'Starting Exchange on-prem Outlook client version inventory from RPC logs.'
Ensure-ExchangeOnPremConnection

$scopeMode = 'Csv'
$resolvedDefaultLogPath = ''
$resolvedDefaultLookbackDays = ''
$resolvedDefaultClientSoftware = ''
$runWasTruncated = $false

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    $resolvedDefaultLogPath = Get-TrimmedValue -Value $LogPath
    $resolvedDefaultLookbackDays = [string]$LookbackDays
    $resolvedDefaultClientSoftware = Get-TrimmedValue -Value $ClientSoftware

    Write-Status -Message "DiscoverAll enabled for RPC log client-version inventory. LogPath='$resolvedDefaultLogPath'; LookbackDays='$resolvedDefaultLookbackDays'." -Level WARN

    $rows = @([PSCustomObject]@{
            RpcLogPath      = $resolvedDefaultLogPath
            LookbackDays    = $resolvedDefaultLookbackDays
            ClientSoftware  = $resolvedDefaultClientSoftware
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

        $clientSoftwareFilter = Get-TrimmedValue -Value $row.ClientSoftware
        if ([string]::IsNullOrWhiteSpace($clientSoftwareFilter)) {
            $clientSoftwareFilter = 'OUTLOOK.EXE'
        }

        $files = @(Resolve-RpcLogFiles -RpcLogPath $rpcLogPath -LookbackDays $effectiveLookbackDays)
        if ($files.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$rpcLogPath|$clientSoftwareFilter" -Action 'GetExchangeRpcLogClientVersions' -Status 'Completed' -Message 'No RPC log files matched the requested lookback window.' -Data ([ordered]@{
                        RpcLogPath             = $rpcLogPath
                        LookbackDays           = [string]$effectiveLookbackDays
                        ClientSoftwareFilter   = $clientSoftwareFilter
                        ClientSoftware         = ''
                        ClientSoftwareVersion  = ''
                        ConnectionCount        = '0'
                        FilesScanned           = '0'
                        RowsParsed             = '0'
                        RowsMatched            = '0'
                        FirstSeenUtc           = ''
                        LastSeenUtc            = ''
                        UniqueClientNames      = '0'
                        UniqueClientIPs        = '0'
                        UniqueServerIPs        = '0'
                    })))
            $rowNumber++
            continue
        }

        $entries = [System.Collections.Generic.List[object]]::new()
        foreach ($file in $files) {
            foreach ($entry in @(Parse-RpcEntries -File $file -Headers $headers)) {
                $entries.Add($entry)
            }
        }

        if ($entries.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$rpcLogPath|$clientSoftwareFilter" -Action 'GetExchangeRpcLogClientVersions' -Status 'Completed' -Message 'RPC log files were found, but no parsable rows were present.' -Data ([ordered]@{
                        RpcLogPath             = $rpcLogPath
                        LookbackDays           = [string]$effectiveLookbackDays
                        ClientSoftwareFilter   = $clientSoftwareFilter
                        ClientSoftware         = ''
                        ClientSoftwareVersion  = ''
                        ConnectionCount        = '0'
                        FilesScanned           = [string]$files.Count
                        RowsParsed             = '0'
                        RowsMatched            = '0'
                        FirstSeenUtc           = ''
                        LastSeenUtc            = ''
                        UniqueClientNames      = '0'
                        UniqueClientIPs        = '0'
                        UniqueServerIPs        = '0'
                    })))
            $rowNumber++
            continue
        }

        $allEntries = @($entries.ToArray())
        $filteredEntries = if ($clientSoftwareFilter -eq '*') {
            $allEntries
        }
        else {
            @($allEntries | Where-Object { (Get-TrimmedValue -Value $_.ClientSoftware).Equals($clientSoftwareFilter, [System.StringComparison]::OrdinalIgnoreCase) })
        }

        if ($filteredEntries.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$rpcLogPath|$clientSoftwareFilter" -Action 'GetExchangeRpcLogClientVersions' -Status 'Completed' -Message 'No RPC log rows matched the requested client software filter.' -Data ([ordered]@{
                        RpcLogPath             = $rpcLogPath
                        LookbackDays           = [string]$effectiveLookbackDays
                        ClientSoftwareFilter   = $clientSoftwareFilter
                        ClientSoftware         = ''
                        ClientSoftwareVersion  = ''
                        ConnectionCount        = '0'
                        FilesScanned           = [string]$files.Count
                        RowsParsed             = [string]$allEntries.Count
                        RowsMatched            = '0'
                        FirstSeenUtc           = ''
                        LastSeenUtc            = ''
                        UniqueClientNames      = '0'
                        UniqueClientIPs        = '0'
                        UniqueServerIPs        = '0'
                    })))
            $rowNumber++
            continue
        }

        $grouped = @(
            $filteredEntries |
                Group-Object -Property ClientSoftwareVersion |
                Sort-Object -Property Count -Descending
        )

        if ($MaxObjects -gt 0 -and $grouped.Count -gt $MaxObjects) {
            $grouped = @($grouped | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        foreach ($group in $grouped) {
            $groupEntries = @($group.Group)
            $firstSeen = ''
            $lastSeen = ''

            $parsedTimes = @()
            foreach ($entry in $groupEntries) {
                try {
                    $parsedTimes += [datetime]$entry.DateTime
                }
                catch {
                    # Keep as blank when parsing fails.
                }
            }

            if ($parsedTimes.Count -gt 0) {
                $firstSeen = ([datetime]($parsedTimes | Sort-Object | Select-Object -First 1)).ToUniversalTime().ToString('o')
                $lastSeen = ([datetime]($parsedTimes | Sort-Object | Select-Object -Last 1)).ToUniversalTime().ToString('o')
            }

            $clientNameCount = @($groupEntries | ForEach-Object { Get-TrimmedValue -Value $_.ClientName } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique).Count
            $clientIpCount = @($groupEntries | ForEach-Object { Get-TrimmedValue -Value $_.ClientIp } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique).Count
            $serverIpCount = @($groupEntries | ForEach-Object { Get-TrimmedValue -Value $_.ServerIp } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique).Count

            $clientSoftwareName = Get-TrimmedValue -Value ($groupEntries | Select-Object -First 1).ClientSoftware
            $clientSoftwareVersion = Get-TrimmedValue -Value $group.Name
            if ([string]::IsNullOrWhiteSpace($clientSoftwareVersion)) {
                $clientSoftwareVersion = '<unknown>'
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$rpcLogPath|$clientSoftwareName|$clientSoftwareVersion" -Action 'GetExchangeRpcLogClientVersions' -Status 'Completed' -Message 'RPC log client-version aggregate exported.' -Data ([ordered]@{
                        RpcLogPath             = $rpcLogPath
                        LookbackDays           = [string]$effectiveLookbackDays
                        ClientSoftwareFilter   = $clientSoftwareFilter
                        ClientSoftware         = $clientSoftwareName
                        ClientSoftwareVersion  = $clientSoftwareVersion
                        ConnectionCount        = [string]$groupEntries.Count
                        FilesScanned           = [string]$files.Count
                        RowsParsed             = [string]$allEntries.Count
                        RowsMatched            = [string]$filteredEntries.Count
                        FirstSeenUtc           = $firstSeen
                        LastSeenUtc            = $lastSeen
                        UniqueClientNames      = [string]$clientNameCount
                        UniqueClientIPs        = [string]$clientIpCount
                        UniqueServerIPs        = [string]$serverIpCount
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($rpcLogPath) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $rpcLogPath -Action 'GetExchangeRpcLogClientVersions' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    RpcLogPath             = $rpcLogPath
                    LookbackDays           = Get-TrimmedValue -Value $row.LookbackDays
                    ClientSoftwareFilter   = Get-TrimmedValue -Value $row.ClientSoftware
                    ClientSoftware         = ''
                    ClientSoftwareVersion  = ''
                    ConnectionCount        = '0'
                    FilesScanned           = '0'
                    RowsParsed             = '0'
                    RowsMatched            = '0'
                    FirstSeenUtc           = ''
                    LastSeenUtc            = ''
                    UniqueClientNames      = '0'
                    UniqueClientIPs        = '0'
                    UniqueServerIPs        = '0'
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeDefaultLogPath' -NotePropertyValue $resolvedDefaultLogPath -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeDefaultLookbackDays' -NotePropertyValue $resolvedDefaultLookbackDays -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeDefaultClientSoftware' -NotePropertyValue $resolvedDefaultClientSoftware -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeMaxObjects' -NotePropertyValue ($(if ($scopeMode -eq 'DiscoverAll') { [string]$MaxObjects } else { '' })) -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeWasTruncated' -NotePropertyValue ([string]$runWasTruncated) -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem Outlook client version inventory from RPC logs completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
