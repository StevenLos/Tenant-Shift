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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\\..\\Standalone_OutputCsvPath') -ChildPath ("Results_SA-IR0225-Get-ExchangeOnPremOutlookClientVersionsFromRpcLogs_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest


function Write-Status {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'INFO' { 'Cyan' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        'SUCCESS' { 'Green' }
    }

    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Start-RunTranscript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OutputCsvPath,

        [AllowNull()]
        [string]$ScriptPath
    )

    $directory = Split-Path -Path $OutputCsvPath -Parent
    if ([string]::IsNullOrWhiteSpace($directory) -and -not [string]::IsNullOrWhiteSpace($ScriptPath)) {
        $directory = Split-Path -Path $ScriptPath -Parent
    }

    if ([string]::IsNullOrWhiteSpace($directory)) {
        throw "Unable to determine transcript directory from OutputCsvPath '$OutputCsvPath'."
    }

    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $scriptName = 'Script'
    if (-not [string]::IsNullOrWhiteSpace($ScriptPath)) {
        $candidate = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            $scriptName = $candidate
        }
    }

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $transcriptPath = Join-Path -Path $directory -ChildPath ("Transcript_{0}_{1}.log" -f $scriptName, $timestamp)

    try {
        Start-Transcript -LiteralPath $transcriptPath -Force -ErrorAction Stop | Out-Null
    }
    catch {
        throw "Failed to start transcript at '$transcriptPath'. Error: $($_.Exception.Message)"
    }

    Write-Status -Message "Transcript started at '$transcriptPath'."
    return $transcriptPath
}

function Stop-RunTranscript {
    [CmdletBinding()]
    param()

    try {
        Stop-Transcript -ErrorAction Stop | Out-Null
    }
    catch {
        $message = ([string]$_.Exception.Message).ToLowerInvariant()
        if ($message -notmatch 'not currently transcribing') {
            throw
        }
    }
}

function ConvertTo-Bool {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value,

        [bool]$Default = $false
    )

    if ($null -eq $Value) {
        return $Default
    }

    $stringValue = [string]$Value
    if ([string]::IsNullOrWhiteSpace($stringValue)) {
        return $Default
    }

    switch -Regex ($stringValue.Trim().ToLowerInvariant()) {
        '^(1|true|t|yes|y)$' { return $true }
        '^(0|false|f|no|n)$' { return $false }
        default { throw "Invalid boolean value '$stringValue'. Use true/false or yes/no." }
    }
}

function ConvertTo-Array {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value,

        [string]$Delimiter = ';'
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    return @(
        $Value -split [Regex]::Escape($Delimiter) |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Get-TrimmedValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value).Trim()
}

function Get-NullableBool {
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

    return (ConvertTo-Bool -Value $text)
}

function Assert-ModuleCurrent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ModuleNames,

        [switch]$FailOnOutdated,

        [switch]$FailOnGalleryLookupError
    )

    foreach ($moduleName in $ModuleNames) {
        Write-Status -Message "Checking module '$moduleName'."

        $installed = Get-Module -ListAvailable -Name $moduleName |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if (-not $installed) {
            throw "Required module '$moduleName' is not installed."
        }

        Write-Status -Message "Installed version for '$moduleName': $($installed.Version)."

        try {
            $gallery = Find-Module -Name $moduleName -Repository PSGallery -ErrorAction Stop
        }
        catch {
            if ($FailOnGalleryLookupError) {
                throw "Unable to verify the latest version for '$moduleName' from PSGallery. Error: $($_.Exception.Message)"
            }

            Write-Status -Message "PSGallery lookup unavailable for '$moduleName'. Continuing with installed version check only." -Level WARN
            continue
        }

        if ($installed.Version -lt $gallery.Version) {
            $message = "Module '$moduleName' is outdated. Installed: $($installed.Version), current: $($gallery.Version)."
            if ($FailOnOutdated) {
                throw "$message Update with: Update-Module $moduleName"
            }

            Write-Status -Message $message -Level WARN
        }
        else {
            Write-Status -Message "Module '$moduleName' is current ($($installed.Version))." -Level SUCCESS
        }
    }
}

function Import-ValidatedCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$InputCsvPath,

        [Parameter(Mandatory)]
        [string[]]$RequiredHeaders
    )

    if (-not (Test-Path -LiteralPath $InputCsvPath -PathType Leaf)) {
        throw "Input CSV file not found: $InputCsvPath"
    }

    $firstLine = Get-Content -LiteralPath $InputCsvPath -TotalCount 1
    if ([string]::IsNullOrWhiteSpace($firstLine)) {
        throw "CSV file '$InputCsvPath' is missing a header row."
    }

    $rawHeaders = @($firstLine -split ',')
    $headers = [System.Collections.Generic.List[string]]::new()
    foreach ($rawHeader in $rawHeaders) {
        $cleanHeader = ([string]$rawHeader).Trim().Trim('"').TrimStart([char]0xFEFF)
        $headers.Add($cleanHeader)
    }

    if ($headers.Count -eq 0) {
        throw "CSV file '$InputCsvPath' is missing a header row."
    }

    $duplicates = @($headers | Group-Object | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Name)
    if ($duplicates.Count -gt 0) {
        throw "CSV file '$InputCsvPath' contains duplicate headers: $($duplicates -join ', ')"
    }

    $missing = @($RequiredHeaders | Where-Object { $_ -notin $headers })
    if ($missing.Count -gt 0) {
        throw "CSV file '$InputCsvPath' is missing required headers: $($missing -join ', ')"
    }

    $rows = Import-Csv -LiteralPath $InputCsvPath
    if (-not $rows -or @($rows).Count -eq 0) {
        throw "CSV file '$InputCsvPath' has no data rows."
    }

    return @($rows)
}
function Ensure-ActiveDirectoryConnection {
    [CmdletBinding()]
    param()

    $isWindowsHost = $false
    $isWindowsVar = Get-Variable -Name IsWindows -ErrorAction SilentlyContinue
    if ($null -ne $isWindowsVar) {
        $isWindowsHost = [bool]$isWindowsVar.Value
    }
    else {
        $isWindowsHost = ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT)
    }

    if (-not $isWindowsHost) {
        throw 'ActiveDirectory scripts require Windows with RSAT/AD tooling available.'
    }

    Assert-ModuleCurrent -ModuleNames @('ActiveDirectory')

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch {
        throw "Unable to import ActiveDirectory module. Error: $($_.Exception.Message)"
    }

    try {
        Get-ADDomain -ErrorAction Stop | Out-Null
    }
    catch {
        throw "Unable to query Active Directory domain context. Ensure domain connectivity and permissions. Error: $($_.Exception.Message)"
    }

    Write-Status -Message 'Active Directory module loaded and domain context verified.' -Level SUCCESS
}

function Escape-AdFilterValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    return $Value.Replace("'", "''")
}

function ConvertTo-NullableDateTime {
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

    try {
        return [datetime]$text
    }
    catch {
        throw "Invalid datetime value '$text'. Use an ISA-like value (for example 2026-03-02 or 2026-03-02T10:30:00)."
    }
}

function Get-HttpStatusCodeFromException {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Exception]$Exception
    )

    foreach ($propertyName in @('ResponseStatusCode', 'StatusCode', 'HttpStatusCode')) {
        if ($Exception.PSObject.Properties.Name -contains $propertyName) {
            $rawValue = $Exception.$propertyName
            if ($null -eq $rawValue) {
                continue
            }

            try {
                return [int]$rawValue
            }
            catch {
                # Continue searching.
            }
        }
    }

    return $null
}

function Test-IsTransientException {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Exception]$Exception
    )

    $statusCode = Get-HttpStatusCodeFromException -Exception $Exception
    if ($null -ne $statusCode -and ($statusCode -eq 429 -or $statusCode -ge 500)) {
        return $true
    }

    $messageChain = [System.Collections.Generic.List[string]]::new()
    $cursor = $Exception
    while ($null -ne $cursor) {
        if (-not [string]::IsNullOrWhiteSpace($cursor.Message)) {
            $messageChain.Add($cursor.Message)
        }

        $cursor = $cursor.InnerException
    }

    $combinedMessage = ($messageChain -join ' ').ToLowerInvariant()
    return ($combinedMessage -match 'temporar|timeout|timed out|service unavailable|rate limit|try again|gateway|429|500|502|503|504|server is not operational')
}

function Get-RetryDelaySeconds {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Exception]$Exception,

        [Parameter(Mandatory)]
        [int]$AttemptNumber,

        [int]$BaseDelaySeconds = 2,
        [int]$MaxDelaySeconds = 60
    )

    $rawDelay = [Math]::Pow(2, [Math]::Min($AttemptNumber, 6)) * $BaseDelaySeconds
    $jitter = Get-Random -Minimum 0 -Maximum 3
    $delay = [int]([Math]::Min($rawDelay + $jitter, $MaxDelaySeconds))
    return [Math]::Max($delay, 1)
}

function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory)]
        [string]$OperationName,

        [ValidateRange(1, 15)]
        [int]$MaxAttempts = 5
    )

    $attempt = 1
    while ($attempt -le $MaxAttempts) {
        try {
            return & $ScriptBlock
        }
        catch {
            $exception = $_.Exception
            if ($attempt -ge $MaxAttempts -or -not (Test-IsTransientException -Exception $exception)) {
                throw
            }

            $delaySeconds = Get-RetryDelaySeconds -Exception $exception -AttemptNumber $attempt
            Write-Status -Level WARN -Message "Transient error during '$OperationName' (attempt $attempt/$MaxAttempts): $($exception.Message). Retrying in $delaySeconds second(s)."
            Start-Sleep -Seconds $delaySeconds
            $attempt++
        }
    }
}

function New-ResultObject {
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
        [string]$Message
    )

    return [PSCustomObject]@{
        TimestampUtc = (Get-Date).ToUniversalTime().ToString('o')
        RowNumber    = $RowNumber
        PrimaryKey   = $PrimaryKey
        Action       = $Action
        Status       = $Status
        Message      = $Message
    }
}

function Export-ResultsCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Results,

        [Parameter(Mandatory)]
        [string]$OutputCsvPath
    )

    $directory = Split-Path -Path $OutputCsvPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $Results | Export-Csv -LiteralPath $OutputCsvPath -NoTypeInformation -Encoding UTF8
    Write-Status -Message "Results exported to '$OutputCsvPath'." -Level SUCCESS
}


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

