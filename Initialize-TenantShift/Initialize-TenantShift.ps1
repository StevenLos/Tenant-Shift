<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260401-000000

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
None

.MODULEVERSIONPOLICY
Shared prerequisite catalog via PrerequisiteCatalog.psd1

.SYNOPSIS
    Validates the local environment for running SharedModule platform scripts
    and can optionally audit repository module requirements.

.DESCRIPTION
    Resolves one or more prerequisite profiles and evaluates them through the
    shared prerequisite engine. Exits with code 0 if all checks pass or produce
    only warnings. Exits with code 1 if any check fails.

    When repository audit parameters are supplied, the script also scans the
    repository for module requirements, reports present/missing/outdated
    modules, and can optionally fail on selected audit findings.

    Compatibility switches:
      -Online  => Contributor + OnlineOperator
      -OnPrem  => Contributor + OnPremOperator

    With no profile arguments, the script performs a full sweep: validates all
    actionable profiles and audits repository module requirements from the
    repository root.

    Test-override environment variables (used by automated tests only):
      PESTER_VERSION_OVERRIDE            — "none", a version string, or unset
      PSSCRIPTANALYZER_VERSION_OVERRIDE  — "none", a version string, or unset
      MODULE_VERSION_OVERRIDE            — "present" to satisfy module checks
      PNP_CLIENT_ID_OVERRIDE             — "absent" or a client ID value
      POWERSHELL_VERSION_OVERRIDE        — a PowerShell version string
      PS_EDITION_OVERRIDE                — "Desktop" or "Core"
      IS_WINDOWS_OVERRIDE                — "true" or "false"
      EXCHANGE_MANAGEMENT_SHELL_OVERRIDE — "true" or "false"
      EXECUTION_POLICY_OVERRIDE          — execution policy text
      EXECUTION_POLICY_LIST_OVERRIDE     — semicolon-delimited scoped execution policy list, e.g. "MachinePolicy=Undefined;UserPolicy=Undefined;Process=Undefined;CurrentUser=Undefined;LocalMachine=RemoteSigned"

.PARAMETER Profile
    Exact prerequisite profile name(s) to evaluate.

.PARAMETER Online
    Compatibility switch: validates Contributor + OnlineOperator.

.PARAMETER OnPrem
    Compatibility switch: validates Contributor + OnPremOperator.

.PARAMETER AuditRepository
    Enables repository module discovery and audit reporting.

.PARAMETER RepositoryRoot
    Repository root to scan when repository audit is enabled.

.PARAMETER IncludeRelativePathPattern
    Relative-path wildcard patterns to include in repository audit results.

.PARAMETER ExcludeRelativePathPattern
    Relative-path wildcard patterns to exclude from repository audit discovery.

.PARAMETER SkipGalleryCheck
    Skips PSGallery version lookups and reports module currency as Unknown.

.PARAMETER OutputFormat
    Console output mode when -PassThru is not used.

.PARAMETER FailOn
    Optional repository-audit finding types that should produce exit code 1.

.PARAMETER PassThru
    Returns a structured report object instead of exiting the host process.

.EXAMPLE
    .\Initialize-TenantShift.ps1

    Validate Contributor, OnlineOperator, and OnPremOperator and audit
    repository module requirements from the repository root in one report.

.EXAMPLE
    .\Initialize-TenantShift.ps1 -Online

    Validate Contributor + OnlineOperator.

.EXAMPLE
    .\Initialize-TenantShift.ps1 -Profile OnPremOperator

    Validate only the OnPremOperator profile.

.EXAMPLE
    .\Initialize-TenantShift.ps1 -AuditRepository -RepositoryRoot .

    Validate the local environment and audit discovered module requirements
    under the current repository root.
#>
#Requires -Version 5.1

# Write-Host is intentional in this script: it is a console validation utility.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
[CmdletBinding()]
param(
    [ValidateSet('Contributor', 'OnlineOperator', 'OnPremOperator', 'RepoScan')]
    [string[]]$Profile,

    [switch]$Online,
    [switch]$OnPrem,

    [switch]$AuditRepository,

    [string]$RepositoryRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),

    [string[]]$IncludeRelativePathPattern = @('*'),

    [string[]]$ExcludeRelativePathPattern = @(
        'SharedModule\Development\Tests\*'
        'SharedModule\Development\Build\*'
    ),

    [switch]$SkipGalleryCheck,

    [ValidateSet('Table', 'Json')]
    [string]$OutputFormat = 'Table',

    [ValidateSet('Missing', 'OutOfDate', 'ProfileFailure')]
    [string[]]$FailOn = @(),

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:enginePath = Join-Path -Path $PSScriptRoot -ChildPath 'PrerequisiteEngine.psm1'
Import-Module $script:enginePath -Force -ErrorAction Stop | Out-Null

function Write-CheckResult {
    [CmdletBinding()]
    # Suppression justification: this script is a console validation utility.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('PASS', 'WARN', 'FAIL')]
        [string]$Result,

        [Parameter(Mandatory)]
        [string]$Check,

        [string]$Detail = '',
        [string]$Fix    = ''
    )

    $color = switch ($Result) {
        'PASS' { 'Green' }
        'WARN' { 'Yellow' }
        'FAIL' { 'Red' }
    }

    $line = '  [{0}]  {1}' -f $Result, $Check
    $detailLines = @()
    if (-not [string]::IsNullOrWhiteSpace($Detail)) {
        $detailLines = @($Detail -split "\r?\n")
    }

    if ($detailLines.Count -le 1) {
        if ($detailLines.Count -eq 1 -and -not [string]::IsNullOrWhiteSpace($detailLines[0])) {
            $line += " — $($detailLines[0])"
        }

        Write-Host $line -ForegroundColor $color
    }
    else {
        Write-Host $line -ForegroundColor $color
        foreach ($detailLine in $detailLines) {
            if ([string]::IsNullOrWhiteSpace($detailLine)) {
                continue
            }

            Write-Host ('         {0}' -f $detailLine) -ForegroundColor $color
        }
    }

    if (($Result -eq 'FAIL' -or $Result -eq 'WARN') -and -not [string]::IsNullOrWhiteSpace($Fix)) {
        Write-Host ('         Fix: {0}' -f $Fix) -ForegroundColor Yellow
    }
}

function ConvertTo-AsciiBannerLines {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Text
    )

    $glyphMap = @{
        'A' = @('    _    ', '   / \   ', '  / _ \  ', ' / ___ \ ', '/_/   \_\')
        'E' = @(' _____ ', '| ____|', '|  _|  ', '| |___ ', '|_____|')
        'F' = @(' _____ ', '|  ___|', '| |_   ', '|  _|  ', '|_|    ')
        'H' = @(' _   _ ', '| | | |', '| |_| |', '|  _  |', '|_| |_|')
        'I' = @(' ___ ', '|_ _|', ' | | ', ' | | ', '|___|')
        'L' = @(' _     ', '| |    ', '| |    ', '| |___ ', '|_____|')
        'M' = @(' __  __ ', '|  \/  |', '| |\/| |', '| |  | |', '|_|  |_|')
        'N' = @(' _   _ ', '| \ | |', '|  \| |', '| |\  |', '|_| \_|')
        'O' = @('  ___  ', ' / _ \ ', '| | | |', '| |_| |', ' \___/ ')
        'P' = @(' ____  ', '|  _ \ ', '| |_) |', '|  __/ ', '|_|    ')
        'R' = @(' ____  ', '|  _ \ ', '| |_) |', '|  _ < ', '|_| \_\')
        'S' = @(' ____  ', '/ ___| ', '\___ \ ', ' ___) |', '|____/ ')
        'T' = @(' _____ ', '|_   _|', '  | |  ', '  | |  ', '  |_|  ')
        'Z' = @(' _____', '|__  /', '  / / ', ' / /_ ', '/____|')
        '-' = @('       ', '       ', ' _____ ', '|_____|', '       ')
        ' ' = @('  ', '  ', '  ', '  ', '  ')
    }

    $characters = $Text.ToUpperInvariant().ToCharArray()
    $outputLines = [System.Collections.Generic.List[string]]::new()

    for ($row = 0; $row -lt 5; $row++) {
        $segments = [System.Collections.Generic.List[string]]::new()

        foreach ($character in $characters) {
            $key = [string]$character
            if (-not $glyphMap.ContainsKey($key)) {
                throw "Unsupported banner character '$key'."
            }

            [void]$segments.Add($glyphMap[$key][$row])
        }

        [void]$outputLines.Add(($segments -join '  ').TrimEnd())
    }

    return @($outputLines)
}

function Write-StartupBanner {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
    param(
        [Parameter(Mandatory)]
        [string[]]$ResolvedProfiles,

        [Parameter(Mandatory)]
        [bool]$RepositoryAuditEnabled,

        [string]$RepositoryRoot
    )

    $tenantShiftBanner = @(ConvertTo-AsciiBannerLines -Text 'Tenant-Shift')

    $modeLabel = if ($RepositoryAuditEnabled -and $ResolvedProfiles.Count -gt 0) {
        'Full sweep: profiles + repository audit'
    }
    elseif ($RepositoryAuditEnabled) {
        'Repository audit only'
    }
    else {
        'Profile validation'
    }

    Write-Host ''
    foreach ($bannerLine in $tenantShiftBanner) {
        Write-Host $bannerLine -ForegroundColor Cyan
    }

    Write-Host ''
    Write-Host '  Initialize-TenantShift' -ForegroundColor DarkCyan
    Write-Host "  Welcome. I'm warming up the platform checks and stirring the module cupboard." -ForegroundColor Yellow
    Write-Host '  The report is being built now; give it a moment and I will hand over the receipts.' -ForegroundColor Yellow
    Write-Host ("  Mode: {0}" -f $modeLabel) -ForegroundColor DarkCyan
    Write-Host ("  Profiles queued: {0}" -f (($ResolvedProfiles | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ', ')) -ForegroundColor DarkCyan

    if ($RepositoryAuditEnabled -and -not [string]::IsNullOrWhiteSpace($RepositoryRoot)) {
        Write-Host ("  Repository root: {0}" -f $RepositoryRoot) -ForegroundColor DarkCyan
    }

    Write-Host ''
}

function Get-ResolvedProfileNames {
    [CmdletBinding()]
    param(
        [string[]]$RequestedProfiles,

        [bool]$UseOnline,

        [bool]$UseOnPrem,

        [Parameter(Mandatory)]
        [hashtable]$Catalog
    )

    if ($UseOnline -and $UseOnPrem) {
        throw 'Specify only one compatibility switch: -Online or -OnPrem.'
    }

    if ($null -ne $RequestedProfiles -and $RequestedProfiles.Count -gt 0) {
        if ($UseOnline -or $UseOnPrem) {
            throw 'Do not combine -Profile with -Online or -OnPrem.'
        }

        return @($RequestedProfiles | Select-Object -Unique)
    }

    if ($UseOnline) {
        return @('Contributor', 'OnlineOperator')
    }

    if ($UseOnPrem) {
        return @('Contributor', 'OnPremOperator')
    }

    $resolved = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $preferredOrder = @('Contributor', 'OnlineOperator', 'OnPremOperator', 'RepoScan')

    foreach ($profileName in $preferredOrder) {
        if (-not $Catalog.Profiles.ContainsKey($profileName)) {
            continue
        }

        if (@($Catalog.Profiles[$profileName].Checks).Count -eq 0) {
            continue
        }

        if ($seen.Add($profileName)) {
            [void]$resolved.Add($profileName)
        }
    }

    foreach ($profileName in @($Catalog.Profiles.Keys | Sort-Object)) {
        if (@($Catalog.Profiles[$profileName].Checks).Count -eq 0) {
            continue
        }

        if ($seen.Add([string]$profileName)) {
            [void]$resolved.Add([string]$profileName)
        }
    }

    return @($resolved)
}

function ConvertTo-OverrideBoolean {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $normalized = $Value.Trim().ToLowerInvariant()
    switch ($normalized) {
        'true' { return $true }
        '1' { return $true }
        'yes' { return $true }
        'false' { return $false }
        '0' { return $false }
        'no' { return $false }
        default { throw "Invalid boolean override value '$Value'." }
    }
}

function Get-HostOverrideMap {
    [CmdletBinding()]
    param()

    $overrides = @{}

    if (-not [string]::IsNullOrWhiteSpace($env:POWERSHELL_VERSION_OVERRIDE)) {
        $overrides.PowerShellVersion = [version]$env:POWERSHELL_VERSION_OVERRIDE
    }

    if (-not [string]::IsNullOrWhiteSpace($env:PS_EDITION_OVERRIDE)) {
        $overrides.PSEdition = $env:PS_EDITION_OVERRIDE
    }

    $isWindowsOverride = ConvertTo-OverrideBoolean -Value $env:IS_WINDOWS_OVERRIDE
    if ($null -ne $isWindowsOverride) {
        $overrides.IsWindows = $isWindowsOverride
    }

    $exchangeManagementShellOverride = ConvertTo-OverrideBoolean -Value $env:EXCHANGE_MANAGEMENT_SHELL_OVERRIDE
    if ($null -ne $exchangeManagementShellOverride) {
        $overrides.ExchangeManagementShell = $exchangeManagementShellOverride
    }

    return $overrides
}

function Get-ModuleVersionOverrideMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$ProfileDefinitions
    )

    $overrides = @{}

    if ($env:MODULE_VERSION_OVERRIDE -eq 'present') {
        foreach ($profileDefinition in $ProfileDefinitions) {
            foreach ($check in @($profileDefinition.Checks | Where-Object { $_.Category -eq 'Module' })) {
                $overrides[$check.FactName] = [version]'99.0.0'
            }
        }
    }

    if ($env:PESTER_VERSION_OVERRIDE -eq 'none') {
        $overrides.Pester = $null
    }
    elseif (-not [string]::IsNullOrWhiteSpace($env:PESTER_VERSION_OVERRIDE)) {
        $overrides.Pester = [version]$env:PESTER_VERSION_OVERRIDE
    }

    if ($env:PSSCRIPTANALYZER_VERSION_OVERRIDE -eq 'none') {
        $overrides.PSScriptAnalyzer = $null
    }
    elseif (-not [string]::IsNullOrWhiteSpace($env:PSSCRIPTANALYZER_VERSION_OVERRIDE)) {
        $overrides.PSScriptAnalyzer = [version]$env:PSSCRIPTANALYZER_VERSION_OVERRIDE
    }

    return $overrides
}

function Get-EnvironmentVariableOverrideMap {
    [CmdletBinding()]
    param()

    $overrides = @{}

    if ($env:PNP_CLIENT_ID_OVERRIDE -eq 'absent') {
        $overrides.PNP_CLIENT_ID = $null
    }
    elseif (-not [string]::IsNullOrWhiteSpace($env:PNP_CLIENT_ID_OVERRIDE)) {
        $overrides.PNP_CLIENT_ID = $env:PNP_CLIENT_ID_OVERRIDE
    }

    return $overrides
}

function Test-IncludeMatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RelativePath,

        [string[]]$Patterns = @('*')
    )

    $effectivePatterns = @($Patterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($effectivePatterns.Count -eq 0) {
        return $true
    }

    foreach ($pattern in $effectivePatterns) {
        if ($RelativePath -like $pattern) {
            return $true
        }
    }

    return $false
}

function Get-InstallCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [bool]$InPsGallery
    )

    if ($InPsGallery) {
        return "Install-Module -Name $ModuleName -Repository PSGallery -Scope CurrentUser -Force"
    }

    if ($ModuleName -eq 'ActiveDirectory') {
        return 'Install RSAT AD module (Server: Install-WindowsFeature RSAT-AD-PowerShell | Client: Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0)'
    }

    return 'Manual install required (not available from PSGallery).'
}

function Get-UpdateCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [bool]$InPsGallery
    )

    if ($InPsGallery) {
        return "Update-Module -Name $ModuleName -Force"
    }

    return ''
}

function Get-ProfileAuditReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$CheckResults,

        [Parameter(Mandatory)]
        [object[]]$ProfileDefinitions
    )

    $results = [System.Collections.Generic.List[object]]::new()

    foreach ($profileDefinition in $ProfileDefinitions) {
        $checks = @($CheckResults | Where-Object { $_.ProfileName -eq $profileDefinition.Name })
        $results.Add([pscustomobject]@{
                ProfileName  = $profileDefinition.Name
                Description  = $profileDefinition.Description
                Checks       = $checks
                FailedCount  = @($checks | Where-Object { $_.Status -eq 'FAIL' }).Count
                WarningCount = @($checks | Where-Object { $_.Status -eq 'WARN' }).Count
            }) | Out-Null
    }

    return @($results)
}

function Get-HostAuditResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$CheckResults
    )

    $hostChecks = @($CheckResults | Where-Object { $_.ProfileName -eq 'Host' })

    return [pscustomobject]@{
        ProfileName  = 'Host'
        Description  = 'Host execution policy validation for internet-downloaded unsigned scripts.'
        Checks       = $hostChecks
        FailedCount  = @($hostChecks | Where-Object { $_.Status -eq 'FAIL' }).Count
        WarningCount = @($hostChecks | Where-Object { $_.Status -eq 'WARN' }).Count
    }
}

function Get-DiscoveredModuleReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RootPath,

        [string[]]$IncludePatterns = @('*'),

        [string[]]$ExcludePatterns = @(),

        [bool]$SkipGalleryVersionCheck
    )

    $discovered = @(Find-ModuleRequirements -RootPath $RootPath -ExcludeRelativePathPattern $ExcludePatterns)
    $results = [System.Collections.Generic.List[object]]::new()

    foreach ($module in $discovered) {
        $sources = @($module.Sources | Where-Object { Test-IncludeMatch -RelativePath $_ -Patterns $IncludePatterns } | Sort-Object)
        if ($sources.Count -eq 0) {
            continue
        }

        $installed = @(Get-Module -ListAvailable -Name $module.ModuleName | Sort-Object Version -Descending | Select-Object -First 1)
        $isPresent = $installed.Count -gt 0
        $installedVersion = if ($isPresent) { [string]$installed[0].Version } else { '' }

        $galleryVersion = ''
        $galleryLookupStatus = 'Skipped'
        $inPsGallery = $false

        if (-not $SkipGalleryVersionCheck) {
            try {
                $galleryModule = Find-Module -Name $module.ModuleName -Repository PSGallery -ErrorAction Stop -WarningAction SilentlyContinue
                $galleryVersion = [string]$galleryModule.Version
                $galleryLookupStatus = 'Found'
                $inPsGallery = $true
            }
            catch {
                $galleryLookupStatus = 'Unavailable'
            }
        }

        $currencyStatus = 'Unknown'
        if (-not $isPresent) {
            $currencyStatus = 'Missing'
        }
        elseif (-not [string]::IsNullOrWhiteSpace($galleryVersion)) {
            $installedParsed = $null
            $galleryParsed = $null

            [void][version]::TryParse($installedVersion, [ref]$installedParsed)
            [void][version]::TryParse($galleryVersion, [ref]$galleryParsed)

            if ($null -ne $installedParsed -and $null -ne $galleryParsed) {
                if ($installedParsed -lt $galleryParsed) {
                    $currencyStatus = 'OutOfDate'
                }
                else {
                    $currencyStatus = 'UpToDate'
                }
            }
        }

        $sourceDisplay = if ($sources.Count -le 3) {
            $sources -join '; '
        }
        else {
            ($sources[0..2] -join '; ') + "; +$($sources.Count - 3) more"
        }

        $installCommand = ''
        $updateCommand = ''

        if (-not $isPresent) {
            $installCommand = Get-InstallCommand -ModuleName $module.ModuleName -InPsGallery:$inPsGallery
        }

        if ($currencyStatus -eq 'OutOfDate') {
            $updateCommand = Get-UpdateCommand -ModuleName $module.ModuleName -InPsGallery:$inPsGallery
        }

        $results.Add([pscustomobject]@{
                ModuleName          = $module.ModuleName
                Present             = $isPresent
                InstalledVersion    = $installedVersion
                GalleryVersion      = $galleryVersion
                GalleryLookupStatus = $galleryLookupStatus
                CurrencyStatus      = $currencyStatus
                SourceCount         = $sources.Count
                Sources             = $sources
                SourceDisplay       = $sourceDisplay
                InstallCommand      = $installCommand
                UpdateCommand       = $updateCommand
            }) | Out-Null
    }

    return @($results | Sort-Object ModuleName)
}

function Get-TriggeredFailOn {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$ModuleResults,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$ProfileResults,

        [string[]]$FailCriteria = @()
    )

    $triggered = [System.Collections.Generic.List[string]]::new()
    $criteria = @($FailCriteria | Select-Object -Unique)

    if ($criteria -contains 'Missing' -and @($ModuleResults | Where-Object { $_.CurrencyStatus -eq 'Missing' }).Count -gt 0) {
        [void]$triggered.Add('Missing')
    }

    if ($criteria -contains 'OutOfDate' -and @($ModuleResults | Where-Object { $_.CurrencyStatus -eq 'OutOfDate' }).Count -gt 0) {
        [void]$triggered.Add('OutOfDate')
    }

    if ($criteria -contains 'ProfileFailure' -and @($ProfileResults | Where-Object { $_.FailedCount -gt 0 }).Count -gt 0) {
        [void]$triggered.Add('ProfileFailure')
    }

    return @($triggered)
}

function Get-EffectiveExecutionPolicy {
    [CmdletBinding()]
    param()

    if (-not [string]::IsNullOrWhiteSpace($env:EXECUTION_POLICY_OVERRIDE)) {
        return $env:EXECUTION_POLICY_OVERRIDE
    }

    try {
        $policy = Get-ExecutionPolicy -Scope CurrentUser
        if ($policy -eq 'Undefined') {
            $policy = Get-ExecutionPolicy
        }

        return $policy
    }
    catch {
        return 'Unknown'
    }
}

function Get-ExecutionPolicyScopeOrder {
    [CmdletBinding()]
    param()

    return @('MachinePolicy', 'UserPolicy', 'Process', 'CurrentUser', 'LocalMachine')
}

function ConvertTo-ExecutionPolicyScopeMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    $scopePolicies = [ordered]@{}
    foreach ($scope in (Get-ExecutionPolicyScopeOrder)) {
        $scopePolicies[$scope] = 'Undefined'
    }

    foreach ($segment in @($Value -split ';')) {
        if ([string]::IsNullOrWhiteSpace($segment)) {
            continue
        }

        $parts = $segment.Split('=', 2, [System.StringSplitOptions]::None)
        if ($parts.Count -ne 2) {
            throw "Invalid execution policy scope override segment '$segment'. Expected Scope=Policy."
        }

        $scopeName = $parts[0].Trim()
        $policyName = $parts[1].Trim()
        if (-not $scopePolicies.Contains($scopeName)) {
            throw "Unknown execution policy scope '$scopeName' in EXECUTION_POLICY_LIST_OVERRIDE."
        }

        if ([string]::IsNullOrWhiteSpace($policyName)) {
            throw "Execution policy scope '$scopeName' in EXECUTION_POLICY_LIST_OVERRIDE is missing a policy value."
        }

        $scopePolicies[$scopeName] = $policyName
    }

    return $scopePolicies
}

function Get-ExecutionPolicyState {
    [CmdletBinding()]
    param()

    $scopePolicies = [ordered]@{}
    foreach ($scope in (Get-ExecutionPolicyScopeOrder)) {
        $scopePolicies[$scope] = 'Undefined'
    }

    if (-not [string]::IsNullOrWhiteSpace($env:EXECUTION_POLICY_LIST_OVERRIDE)) {
        $scopePolicies = ConvertTo-ExecutionPolicyScopeMap -Value $env:EXECUTION_POLICY_LIST_OVERRIDE
    }
    else {
        try {
            $policyList = @(Get-ExecutionPolicy -List)
            foreach ($entry in $policyList) {
                $scopeName = [string]$entry.Scope
                if ($scopePolicies.Contains($scopeName)) {
                    $scopePolicies[$scopeName] = [string]$entry.ExecutionPolicy
                }
            }
        }
        catch {
            return [pscustomobject]@{
                EffectivePolicy = 'Unknown'
                EffectiveScope  = 'Unknown'
                ScopePolicies   = $scopePolicies
                ScopeSummary    = 'MachinePolicy=Unknown, UserPolicy=Unknown, Process=Unknown, CurrentUser=Unknown, LocalMachine=Unknown'
            }
        }
    }

    $effectivePolicy = $null
    $effectiveScope = $null

    foreach ($scope in (Get-ExecutionPolicyScopeOrder)) {
        $policy = [string]$scopePolicies[$scope]
        if (-not [string]::IsNullOrWhiteSpace($policy) -and $policy -ne 'Undefined') {
            $effectivePolicy = $policy
            $effectiveScope = $scope
            break
        }
    }

    if ($null -eq $effectivePolicy) {
        $effectivePolicy = Get-EffectiveExecutionPolicy
        $effectiveScope = if ($effectivePolicy -eq 'Unknown') { 'Unknown' } else { 'Default' }
    }

    $scopeSummary = ((Get-ExecutionPolicyScopeOrder) | ForEach-Object {
            '{0}={1}' -f $_, [string]$scopePolicies[$_]
        }) -join ', '

    return [pscustomobject]@{
        EffectivePolicy = [string]$effectivePolicy
        EffectiveScope  = [string]$effectiveScope
        ScopePolicies   = $scopePolicies
        ScopeSummary    = $scopeSummary
    }
}

function Test-ExecutionPolicyAllowsDownloadedUnsignedScripts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Policy
    )

    switch ($Policy) {
        'Bypass' { return $true }
        'Unrestricted' { return $true }
        'Undefined' { return $false }
        'RemoteSigned' { return $false }
        'Restricted' { return $false }
        'AllSigned' { return $false }
        'Unknown' { return $false }
        default { return $false }
    }
}

function Get-ExecutionPolicyStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Policy
    )

    switch ($Policy) {
        'Bypass' { return 'PASS' }
        'Unrestricted' { return 'PASS' }
        'Undefined' { return 'PASS' }
        'RemoteSigned' { return 'WARN' }
        'Restricted' { return 'FAIL' }
        'AllSigned' { return 'FAIL' }
        'Unknown' { return 'FAIL' }
        default { return 'FAIL' }
    }
}

function Test-ExecutionPolicyScopeIsNonBlocking {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Policy
    )

    switch ($Policy) {
        'Undefined' { return $true }
        'Bypass' { return $true }
        'Unrestricted' { return $true }
        'RemoteSigned' { return $true }
        default { return $false }
    }
}

function New-ExecutionPolicyResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$ExecutionPolicyState
    )

    $effectiveScope = [string]$ExecutionPolicyState.EffectiveScope
    $results = [System.Collections.Generic.List[object]]::new()
    $effectivePolicy = [string]$ExecutionPolicyState.EffectivePolicy
    $effectiveAllowsDownloadedUnsigned = Test-ExecutionPolicyAllowsDownloadedUnsignedScripts -Policy $effectivePolicy
    $effectiveStatus = Get-ExecutionPolicyStatus -Policy $effectivePolicy
    $effectiveRemediation = ''
    $effectiveDetail = ('{0} from {1}' -f $effectivePolicy, $effectiveScope)

    if ($effectivePolicy -eq 'RemoteSigned') {
        $effectiveRemediation = 'Run Utilities\Utility-Unblock-Files\Utility-Unblock-Files.ps1 against the downloaded scripts before execution, or set the execution policy to Unrestricted or Bypass for the intended scope.'
        $effectiveDetail = ('{0} from {1} — downloaded scripts will need Unblock-File' -f $effectivePolicy, $effectiveScope)
    }
    elseif (-not $effectiveAllowsDownloadedUnsigned) {
        if ($effectiveScope -in @('MachinePolicy', 'UserPolicy')) {
            $effectiveRemediation = 'A Group Policy-defined execution policy is blocking scripts. Adjust the policy in Group Policy or use an approved execution path for this host.'
        }
        else {
            $effectiveRemediation = 'Run Utilities\Utility-Unblock-Files\Utility-Unblock-Files.ps1 against the downloaded scripts, or set the execution policy to Unrestricted or Bypass for the intended scope.'
        }
    }

    $results.Add([pscustomobject]@{
            ProfileName     = 'Host'
            CheckName       = 'Effective execution policy'
            DisplayName     = 'Effective execution policy'
            Category        = 'ExecutionPolicy'
            RequirementType = 'Required'
            Status          = $effectiveStatus
            Severity        = if ($effectiveStatus -eq 'FAIL') { 'Error' } elseif ($effectiveStatus -eq 'WARN') { 'Warning' } else { 'Information' }
            Passed          = ($effectiveStatus -ne 'FAIL')
            Expected        = 'Allows internet-downloaded unsigned scripts directly, or with Unblock-File'
            Actual          = ('{0} (from {1})' -f $effectivePolicy, $effectiveScope)
            Detail          = $effectiveDetail
            Remediation     = $effectiveRemediation
            EffectiveScope  = $effectiveScope
            ScopePolicies   = $ExecutionPolicyState.ScopePolicies
            ScopeSummary    = $ExecutionPolicyState.ScopeSummary
            AffectsOverall  = $true
            SortOrder       = 0
        }) | Out-Null

    $scopeIndex = 0
    foreach ($scopeName in (Get-ExecutionPolicyScopeOrder)) {
        $scopeIndex++
        $scopePolicy = [string]$ExecutionPolicyState.ScopePolicies[$scopeName]
        $scopeStatus = Get-ExecutionPolicyStatus -Policy $scopePolicy
        $scopeDetail = $scopePolicy
        if ($scopePolicy -eq 'RemoteSigned') {
            $scopeDetail = 'RemoteSigned — downloaded scripts will need Unblock-File'
        }

        $results.Add([pscustomobject]@{
                ProfileName     = 'Host'
                CheckName       = ('{0} execution policy' -f $scopeName)
                DisplayName     = ('{0} execution policy' -f $scopeName)
                Category        = 'ExecutionPolicyScope'
                RequirementType = 'Informational'
                Status          = $scopeStatus
                Severity        = 'Information'
                Passed          = ($scopeStatus -ne 'FAIL')
                Expected        = 'Does not block downloaded unsigned scripts, or only requires Unblock-File'
                Actual          = $scopePolicy
                Detail          = $scopeDetail
                Remediation     = ''
                ScopeName       = $scopeName
                ScopePolicies   = $ExecutionPolicyState.ScopePolicies
                ScopeSummary    = $ExecutionPolicyState.ScopeSummary
                AffectsOverall  = $false
                SortOrder       = $scopeIndex
            }) | Out-Null
    }

    return @($results)
}

function Get-CategorySortRank {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Category
    )

    switch ($Category) {
        'Runtime' { return 0 }
        'Module' { return 1 }
        'Configuration' { return 2 }
        'ExecutionPolicy' { return 3 }
        'ExecutionPolicyScope' { return 4 }
        default { return 9 }
    }
}

function Write-SectionTable {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [AllowEmptyCollection()]
        [object[]]$Rows,

        [Parameter(Mandatory)]
        [string[]]$Columns,

        [Parameter(Mandatory)]
        [string]$EmptyMessage
    )

    Write-Host ''
    Write-Host ("[ {0} ]" -f $Title) -ForegroundColor Cyan

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        Write-Host $EmptyMessage -ForegroundColor DarkGray
        return
    }

    $rendered = $Rows |
        Select-Object -Property $Columns |
        Format-Table -AutoSize |
        Out-String

    Write-Host ($rendered.TrimEnd())
}

function Get-PlatformReport {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()]
        [string[]]$RequestedProfiles,

        [Parameter(Mandatory)]
        [string[]]$ResolvedProfiles,

        [Parameter(Mandatory)]
        [object[]]$ProfileDefinitions,

        [AllowEmptyCollection()]
        [object[]]$CheckResults,

        [Parameter(Mandatory)]
        [bool]$RepositoryAuditEnabled,

        [string]$ResolvedRepositoryRoot,

        [string[]]$IncludePatterns = @('*'),

        [string[]]$ExcludePatterns = @(),

        [bool]$SkipGalleryVersionCheck,

        [string[]]$FailCriteria = @()
    )

    $profileResults = @(Get-ProfileAuditReport -CheckResults $CheckResults -ProfileDefinitions $ProfileDefinitions)
    $hostResult = Get-HostAuditResult -CheckResults $CheckResults

    $moduleResults = @()
    if ($RepositoryAuditEnabled) {
        $moduleResults = @(Get-DiscoveredModuleReport -RootPath $ResolvedRepositoryRoot `
                -IncludePatterns $IncludePatterns `
                -ExcludePatterns $ExcludePatterns `
                -SkipGalleryVersionCheck:$SkipGalleryVersionCheck)
    }

    $triggeredFailOn = @(Get-TriggeredFailOn -ModuleResults $moduleResults -ProfileResults $profileResults -FailCriteria $FailCriteria)

    $blockingResults = @($CheckResults | Where-Object {
            -not ($_.PSObject.Properties.Name -contains 'AffectsOverall') -or $_.AffectsOverall
        })
    $blockingFailures = @($blockingResults | Where-Object { $_.Status -eq 'FAIL' })

    $missingCount = @($moduleResults | Where-Object { $_.CurrencyStatus -eq 'Missing' }).Count
    $presentCount = @($moduleResults | Where-Object { $_.Present }).Count
    $upToDateCount = @($moduleResults | Where-Object { $_.CurrencyStatus -eq 'UpToDate' }).Count
    $outOfDateCount = @($moduleResults | Where-Object { $_.CurrencyStatus -eq 'OutOfDate' }).Count
    $unknownCount = @($moduleResults | Where-Object { $_.Present -and $_.CurrencyStatus -eq 'Unknown' }).Count

    $profileFailureCount = 0
    $profileWarningCount = 0
    foreach ($profileResult in $profileResults) {
        $profileFailureCount += $profileResult.FailedCount
        $profileWarningCount += $profileResult.WarningCount
    }

    return [pscustomobject]@{
        RequestedProfiles          = @($RequestedProfiles | Select-Object -Unique)
        ResolvedProfiles           = @($ResolvedProfiles | Select-Object -Unique)
        RepositoryAuditEnabled     = $RepositoryAuditEnabled
        RepositoryRoot             = $ResolvedRepositoryRoot
        IncludeRelativePathPattern = @($IncludePatterns)
        ExcludeRelativePathPattern = @($ExcludePatterns)
        Checks                     = @($CheckResults)
        ProfileResults             = $profileResults
        HostResult                 = $hostResult
        DiscoveredModules          = $moduleResults
        TriggeredFailOn            = $triggeredFailOn
        BlockingFailures           = $blockingFailures
        ShouldFail                 = (($blockingFailures.Count -gt 0) -or ($triggeredFailOn.Count -gt 0))
        Summary                    = [pscustomobject]@{
            CheckPass         = @($blockingResults | Where-Object { $_.Status -eq 'PASS' }).Count
            CheckWarn         = @($blockingResults | Where-Object { $_.Status -eq 'WARN' }).Count
            CheckFail         = $blockingFailures.Count
            DiscoveredModules = $moduleResults.Count
            Present           = $presentCount
            Missing           = $missingCount
            UpToDate          = $upToDateCount
            OutOfDate         = $outOfDateCount
            UnknownCurrency   = $unknownCount
            ProfileFailures   = $profileFailureCount
            ProfileWarnings   = $profileWarningCount
            HostFailures      = $hostResult.FailedCount
            HostWarnings      = $hostResult.WarningCount
        }
    }
}

function Write-PlatformReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Report
    )

    Write-Host ''
    Write-Host '═══════════════════════════════════════════════════' -ForegroundColor Cyan
    Write-Host '  SharedModule Platform — Environment Validation' -ForegroundColor Cyan
    Write-Host '═══════════════════════════════════════════════════' -ForegroundColor Cyan
    Write-Host ''
    Write-Host ('Profiles: {0}' -f (($Report.ResolvedProfiles | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ', ')) -ForegroundColor DarkCyan

    foreach ($profileResult in $Report.ProfileResults) {
        Write-Host ''
        Write-Host ("[ Profile: {0} ]" -f $profileResult.ProfileName) -ForegroundColor Cyan

        if ($profileResult.Checks.Count -eq 0) {
            Write-Host 'No direct prerequisite checks are defined for this profile.' -ForegroundColor DarkGray
            continue
        }

        foreach ($result in $profileResult.Checks) {
            Write-CheckResult -Result $result.Status -Check $result.DisplayName -Detail $result.Detail -Fix $result.Remediation
        }
    }

    if ($Report.HostResult.Checks.Count -gt 0) {
        Write-Host ''
        Write-Host '[ Host ]' -ForegroundColor Cyan
        foreach ($hostCheck in $Report.HostResult.Checks) {
            $hostCheckName = if ($hostCheck.PSObject.Properties.Name -contains 'DisplayName' -and
                -not [string]::IsNullOrWhiteSpace([string]$hostCheck.DisplayName)) {
                [string]$hostCheck.DisplayName
            }
            else {
                [string]$hostCheck.CheckName
            }

            Write-CheckResult -Result $hostCheck.Status -Check $hostCheckName `
                -Detail $hostCheck.Detail `
                -Fix $hostCheck.Remediation
        }
    }

    if ($Report.RepositoryAuditEnabled) {
        Write-Host ''
        Write-Host '[ Repository Audit ]' -ForegroundColor Cyan
        Write-Host ('Root: {0}' -f $Report.RepositoryRoot) -ForegroundColor DarkCyan
        Write-Host ('Include patterns: {0}' -f (@($Report.IncludeRelativePathPattern) -join ', ')) -ForegroundColor DarkCyan
        Write-Host ('Exclude patterns: {0}' -f (@($Report.ExcludeRelativePathPattern) -join ', ')) -ForegroundColor DarkCyan

        Write-SectionTable -Title 'Modules Present' `
            -Rows @($Report.DiscoveredModules | Where-Object { $_.Present }) `
            -Columns @('ModuleName', 'InstalledVersion', 'GalleryVersion', 'CurrencyStatus', 'SourceCount') `
            -EmptyMessage 'No discovered modules are currently present.'

        Write-SectionTable -Title 'Modules Missing' `
            -Rows @($Report.DiscoveredModules | Where-Object { $_.CurrencyStatus -eq 'Missing' }) `
            -Columns @('ModuleName', 'GalleryVersion', 'InstallCommand', 'SourceDisplay') `
            -EmptyMessage 'No discovered modules are missing.'

        Write-SectionTable -Title 'Modules Out Of Date' `
            -Rows @($Report.DiscoveredModules | Where-Object { $_.CurrencyStatus -eq 'OutOfDate' }) `
            -Columns @('ModuleName', 'InstalledVersion', 'GalleryVersion', 'UpdateCommand', 'SourceDisplay') `
            -EmptyMessage 'No discovered modules are currently out of date.'

        Write-SectionTable -Title 'Modules With Unknown Currency' `
            -Rows @($Report.DiscoveredModules | Where-Object { $_.Present -and $_.CurrencyStatus -eq 'Unknown' }) `
            -Columns @('ModuleName', 'InstalledVersion', 'GalleryLookupStatus', 'SourceDisplay') `
            -EmptyMessage 'No discovered modules have unknown currency.'
    }

    Write-Host ''
    Write-Host '═══════════════════════════════════════════════════' -ForegroundColor Cyan
    Write-Host ("  Checks: PASS={0} WARN={1} FAIL={2}" -f $Report.Summary.CheckPass, $Report.Summary.CheckWarn, $Report.Summary.CheckFail) -ForegroundColor Cyan

    if ($Report.RepositoryAuditEnabled) {
        Write-Host ("  Modules: total={0} present={1} missing={2} up-to-date={3} out-of-date={4} unknown={5}" -f `
                $Report.Summary.DiscoveredModules, `
                $Report.Summary.Present, `
                $Report.Summary.Missing, `
                $Report.Summary.UpToDate, `
                $Report.Summary.OutOfDate, `
                $Report.Summary.UnknownCurrency) -ForegroundColor Cyan
    }

    if ($Report.TriggeredFailOn.Count -gt 0) {
        Write-Host ("  FailOn triggered: {0}" -f ($Report.TriggeredFailOn -join ', ')) -ForegroundColor Yellow
    }

    if ($Report.BlockingFailures.Count -gt 0) {
        Write-Host '  OVERALL: FAIL — resolve the items above before running platform scripts.' -ForegroundColor Red
        return
    }

    if ($Report.TriggeredFailOn.Count -gt 0) {
        Write-Host '  OVERALL: FAIL — repository audit triggered the requested fail criteria.' -ForegroundColor Red
        return
    }

    Write-Host '  OVERALL: PASS — environment is ready.' -ForegroundColor Green
}

$catalog = Get-PrerequisiteCatalog
$resolvedProfileNames = Get-ResolvedProfileNames -RequestedProfiles $Profile -UseOnline:$Online -UseOnPrem:$OnPrem -Catalog $catalog
$defaultSelectionParameterNames = @(
    'Profile'
    'Online'
    'OnPrem'
    'AuditRepository'
    'RepositoryRoot'
    'IncludeRelativePathPattern'
    'ExcludeRelativePathPattern'
    'SkipGalleryCheck'
    'FailOn'
)
$isDefaultFullSweep = $true
foreach ($parameterName in $defaultSelectionParameterNames) {
    if ($PSBoundParameters.ContainsKey($parameterName)) {
        $isDefaultFullSweep = $false
        break
    }
}

$repositoryAuditEnabled = $AuditRepository.IsPresent -or $isDefaultFullSweep
$repositoryAuditParameterNames = @(
    'RepositoryRoot'
    'IncludeRelativePathPattern'
    'ExcludeRelativePathPattern'
    'SkipGalleryCheck'
    'FailOn'
)

foreach ($parameterName in $repositoryAuditParameterNames) {
    if ($PSBoundParameters.ContainsKey($parameterName)) {
        $repositoryAuditEnabled = $true
        break
    }
}

$resolvedRepositoryRoot = ''
if ($repositoryAuditEnabled) {
    $resolvedRepositoryRoot = (Resolve-Path -LiteralPath $RepositoryRoot).Path
}

if (-not $PassThru -and $OutputFormat -eq 'Table') {
    Write-StartupBanner -ResolvedProfiles $resolvedProfileNames `
        -RepositoryAuditEnabled:$repositoryAuditEnabled `
        -RepositoryRoot $resolvedRepositoryRoot
}

$profileDefinitions = foreach ($profileName in $resolvedProfileNames) {
    Resolve-PrerequisiteProfile -ProfileName $profileName -Catalog $catalog
}

$hostOverride = Get-HostOverrideMap
$moduleVersionOverride = Get-ModuleVersionOverrideMap -ProfileDefinitions $profileDefinitions
$environmentVariableOverride = Get-EnvironmentVariableOverrideMap

$allResults = [System.Collections.Generic.List[object]]::new()

foreach ($profileDefinition in $profileDefinitions) {
    $state = Get-PrerequisiteState -ProfileDefinition $profileDefinition -Catalog $catalog `
        -HostOverride $hostOverride `
        -ModuleVersionOverride $moduleVersionOverride `
        -EnvironmentVariableOverride $environmentVariableOverride

    foreach ($result in (Test-PrerequisiteProfile -ProfileName $profileDefinition.Name -Catalog $catalog -State $state)) {
        $allResults.Add($result) | Out-Null
    }
}

if (@($profileDefinitions | Where-Object { @($_.Checks).Count -gt 0 }).Count -gt 0) {
    $executionPolicyResults = @(New-ExecutionPolicyResults -ExecutionPolicyState (Get-ExecutionPolicyState))
    foreach ($executionPolicyResult in $executionPolicyResults) {
        $allResults.Add($executionPolicyResult) | Out-Null
    }
}

$orderedResults = @($allResults | Sort-Object `
    @{ Expression = { if ($_.ProfileName -eq 'Host') { 1 } else { 0 } } }, `
    ProfileName, `
    @{ Expression = { Get-CategorySortRank -Category $_.Category } }, `
    @{ Expression = { if ($_.PSObject.Properties.Name -contains 'SortOrder') { [int]$_.SortOrder } else { 999 } } }, `
    CheckName)

$report = Get-PlatformReport `
    -RequestedProfiles @($Profile | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) `
    -ResolvedProfiles $resolvedProfileNames `
    -ProfileDefinitions $profileDefinitions `
    -CheckResults $orderedResults `
    -RepositoryAuditEnabled:$repositoryAuditEnabled `
    -ResolvedRepositoryRoot $resolvedRepositoryRoot `
    -IncludePatterns $IncludeRelativePathPattern `
    -ExcludePatterns $ExcludeRelativePathPattern `
    -SkipGalleryVersionCheck:$SkipGalleryCheck `
    -FailCriteria $FailOn

if ($PassThru) {
    return $report
}

switch ($OutputFormat) {
    'Json' {
        $report | ConvertTo-Json -Depth 8
    }
    default {
        Write-PlatformReport -Report $report
    }
}

if ($report.ShouldFail) {
    exit 1
}

exit 0
