<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260315-000300

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
None declared in this file

.MODULEVERSIONPOLICY
Not declared in this file
#>
#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [ValidateRange(1, 100000)]
    [int]$PasswordCount = 20,

    [string]$OutputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath ("Results_Utility-Generate-Passwords_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))),

    [switch]$NoTranscript
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

    Start-Transcript -LiteralPath $transcriptPath -Force -ErrorAction Stop | Out-Null
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

function Import-ValidatedCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$InputCsvPath
    )

    if (-not (Test-Path -LiteralPath $InputCsvPath -PathType Leaf)) {
        throw "Input CSV file not found: $InputCsvPath"
    }

    $rows = @(Import-Csv -LiteralPath $InputCsvPath)
    if ($rows.Count -eq 0) {
        throw "CSV file '$InputCsvPath' has no data rows."
    }

    $headers = [System.Collections.Generic.List[string]]::new()
    foreach ($header in @($rows[0].PSObject.Properties.Name)) {
        $cleanHeader = ([string]$header).Trim().TrimStart([char]0xFEFF)
        if (-not [string]::IsNullOrWhiteSpace($cleanHeader)) {
            [void]$headers.Add($cleanHeader)
        }
    }

    if ($headers.Count -eq 0) {
        throw "CSV file '$InputCsvPath' has no usable column headers."
    }

    return [pscustomobject]@{
        Rows    = $rows
        Headers = @($headers)
    }
}

function Get-CryptoRandomInt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$MaxExclusive,

        [Parameter(Mandatory)]
        [System.Security.Cryptography.RandomNumberGenerator]$Rng
    )

    if ($MaxExclusive -le 0) {
        throw 'MaxExclusive must be greater than zero.'
    }

    $bytes = New-Object byte[] 4
    $upperBound = [uint32]::MaxValue
    $remainder = $upperBound % [uint32]$MaxExclusive
    $limit = $upperBound - $remainder

    do {
        $Rng.GetBytes($bytes)
        $value = [BitConverter]::ToUInt32($bytes, 0)
    } while ($value -ge $limit)

    return [int]($value % [uint32]$MaxExclusive)
}

function Get-ColumnPools {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,

        [Parameter(Mandatory)]
        [string[]]$Headers
    )

    $pools = [System.Collections.Generic.List[object]]::new()

    foreach ($header in $Headers) {
        $values = [System.Collections.Generic.List[string]]::new()
        $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)

        foreach ($row in $Rows) {
            $property = $row.PSObject.Properties[$header]
            if ($null -eq $property) {
                continue
            }

            $value = Get-TrimmedValue -Value $property.Value
            if ([string]::IsNullOrWhiteSpace($value)) {
                continue
            }

            if ($seen.Add($value)) {
                [void]$values.Add($value)
            }
        }

        if ($values.Count -eq 0) {
            throw "Column '$header' has no non-empty values."
        }

        $pools.Add([pscustomobject]@{
                Header = $header
                Values = @($values)
            })
    }

    return @($pools)
}

function Get-CsvCountControl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,

        [Parameter(Mandatory)]
        [string[]]$Headers
    )

    if ($Headers.Count -eq 0) {
        throw 'CSV header list cannot be empty.'
    }

    $firstHeader = $Headers[0]
    if (-not [string]::Equals($firstHeader, 'PasswordCount', [System.StringComparison]::OrdinalIgnoreCase)) {
        return [pscustomobject]@{
            IsPresent        = $false
            ControlHeader    = $null
            Count            = $null
            ComponentHeaders = @($Headers)
        }
    }

    if ($Headers.Count -lt 2) {
        throw "CSV includes control header '$firstHeader' but has no password component headers."
    }

    $values = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    foreach ($row in $Rows) {
        $property = $row.PSObject.Properties[$firstHeader]
        if ($null -eq $property) {
            continue
        }

        $value = Get-TrimmedValue -Value $property.Value
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        [void]$values.Add($value)
    }

    if ($values.Count -eq 0) {
        throw "Control column '$firstHeader' has no value. Provide at least one integer from 1 to 100000."
    }

    if ($values.Count -gt 1) {
        $displayValues = @($values) -join ', '
        throw "Control column '$firstHeader' has multiple values ($displayValues). Use only one value."
    }

    $rawCount = [string](@($values)[0])
    $parsedCount = 0
    if (-not [int]::TryParse($rawCount, [ref]$parsedCount) -or $parsedCount -lt 1 -or $parsedCount -gt 100000) {
        throw "Control column '$firstHeader' value '$rawCount' is invalid. Use an integer from 1 to 100000."
    }

    return [pscustomobject]@{
        IsPresent        = $true
        ControlHeader    = $firstHeader
        Count            = $parsedCount
        ComponentHeaders = @($Headers | Select-Object -Skip 1)
    }
}

function Get-RandomValueFromPool {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Values,

        [Parameter(Mandatory)]
        [System.Security.Cryptography.RandomNumberGenerator]$Rng
    )

    if ($Values.Count -eq 0) {
        throw 'Value pool cannot be empty.'
    }

    $index = Get-CryptoRandomInt -MaxExclusive $Values.Count -Rng $Rng
    return [string]$Values[$index]
}

$rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
$transcriptStarted = $false

try {
    if (-not $NoTranscript) {
        Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath | Out-Null
        $transcriptStarted = $true
    }

    Write-Status -Message 'Starting utility password generation script.'

    $csvData = Import-ValidatedCsv -InputCsvPath $InputCsvPath
    $countControl = Get-CsvCountControl -Rows $csvData.Rows -Headers $csvData.Headers
    $effectivePasswordCount = $PasswordCount
    if ($countControl.IsPresent) {
        if ($PSBoundParameters.ContainsKey('PasswordCount')) {
            Write-Status -Level WARN -Message "CSV requested $($countControl.Count) password(s) in '$($countControl.ControlHeader)', but explicit -PasswordCount $PasswordCount takes precedence."
        }
        else {
            $effectivePasswordCount = $countControl.Count
            Write-Status -Message "Using PasswordCount=$effectivePasswordCount from CSV control column '$($countControl.ControlHeader)'."
        }
    }

    $columnPools = Get-ColumnPools -Rows $csvData.Rows -Headers $countControl.ComponentHeaders
    if ($columnPools.Count -eq 0) {
        throw 'No password component columns were found after control-column processing.'
    }

    $results = [System.Collections.Generic.List[object]]::new()

    for ($passwordIndex = 1; $passwordIndex -le $effectivePasswordCount; $passwordIndex++) {
        $selectedByHeader = [ordered]@{}
        $components = [System.Collections.Generic.List[string]]::new()

        foreach ($pool in $columnPools) {
            $selectedValue = Get-RandomValueFromPool -Values $pool.Values -Rng $rng
            $selectedByHeader[$pool.Header] = $selectedValue
            [void]$components.Add($selectedValue)
        }

        $password = ($components -join '')

        $resultObject = [ordered]@{
            TimestampUtc = (Get-Date).ToUniversalTime().ToString('o')
            PasswordIndex = $passwordIndex
            Password = $password
        }

        foreach ($header in $selectedByHeader.Keys) {
            $resultObject[$header] = $selectedByHeader[$header]
        }

        $results.Add([pscustomobject]$resultObject)
    }

    $outputDirectory = Split-Path -Path $OutputCsvPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    $results.ToArray() | Export-Csv -LiteralPath $OutputCsvPath -NoTypeInformation -Encoding UTF8

    Write-Status -Level SUCCESS -Message "Generated $effectivePasswordCount password(s) using $($columnPools.Count) column pools."
    Write-Status -Level SUCCESS -Message "Password generation results exported to '$OutputCsvPath'."
}
finally {
    if ($null -ne $rng) {
        $rng.Dispose()
    }

    if ($transcriptStarted) {
        Stop-RunTranscript
    }
}
