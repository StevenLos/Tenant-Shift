<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260302-190500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
None declared in this file

.MODULEVERSIONPOLICY
Best-effort PSGallery check (offline-safe)
#>
Set-StrictMode -Version Latest

# Load shared utility functions (dot-sourced, not Import-Module).
# Write-Status is used as sentinel: if it is already in scope, Shared.Common.ps1 was
# already dot-sourced by a prior module load in this session — skip to avoid redefining.
if (-not (Get-Command Write-Status -ErrorAction SilentlyContinue)) {
    . "$PSScriptRoot\..\Shared\Shared.Common.ps1"
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

function Ensure-ActiveDirectoryConnection {
    [CmdletBinding()]
    # Suppression justification: Internal platform helper. Verb is intentional and
    # predates PSScriptAnalyzer enforcement; renaming requires updating all dependent scripts.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
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

function Ensure-ExchangeOnPremConnection {
    [CmdletBinding()]
    # Suppression justification: Internal platform helper. Verb is intentional and
    # predates PSScriptAnalyzer enforcement; renaming requires updating all dependent scripts.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
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
        throw 'ExchangeOnPrem scripts require Windows and Exchange Management Shell tooling.'
    }

    $version = $PSVersionTable.PSVersion
    $isDesktopEdition = ($PSVersionTable.PSObject.Properties.Name -contains 'PSEdition') -and ([string]$PSVersionTable.PSEdition -eq 'Desktop')
    if (-not $isDesktopEdition -or $version.Major -ne 5 -or $version.Minor -lt 1) {
        throw 'ExchangeOnPrem scripts require Windows PowerShell 5.1 (Desktop edition) in Exchange Management Shell.'
    }

    $requiredCommands = @(
        'Get-Recipient',
        'Get-MailContact',
        'Get-DistributionGroup',
        'Get-DynamicDistributionGroup'
    )

    $missingCommands = @()
    foreach ($commandName in $requiredCommands) {
        if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) {
            $missingCommands += $commandName
        }
    }

    if ($missingCommands.Count -gt 0) {
        throw "Exchange management cmdlets were not found in the current session: $($missingCommands -join ', '). Launch Exchange Management Shell first."
    }

    try {
        Invoke-WithRetry -OperationName 'Validate Exchange on-prem management context' -ScriptBlock {
            Get-Recipient -ResultSize 1 -ErrorAction Stop | Out-Null
        }
    }
    catch {
        throw "Unable to validate Exchange on-prem management context. Ensure EMS session connectivity and permissions. Error: $($_.Exception.Message)"
    }

    Write-Status -Message 'Exchange on-prem management cmdlets detected and session context verified.' -Level SUCCESS
}

function Escape-AdFilterValue {
    [CmdletBinding()]
    # Suppression justification: Internal platform helper. Verb is intentional and
    # predates PSScriptAnalyzer enforcement; renaming requires updating all dependent scripts.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
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

Export-ModuleMember -Function @(
    # Shared utility functions — loaded via dot-source of Shared.Common.ps1 (guard at top of file)
    'Write-Status',
    'Start-RunTranscript',
    'Stop-RunTranscript',
    'ConvertTo-Bool',
    'ConvertTo-Array',
    'Import-ValidatedCsv',
    'New-ResultObject',
    'Export-ResultsCsv',
    'Get-TrimmedValue',
    'Convert-MultiValueToString',
    'Convert-ToOrderedReportObject',
    # OnPrem-only functions
    'Get-NullableBool',
    'Assert-ModuleCurrent',
    'Ensure-ActiveDirectoryConnection',
    'Ensure-ExchangeOnPremConnection',
    'Escape-AdFilterValue',
    'ConvertTo-NullableDateTime',
    'Invoke-WithRetry'
)
