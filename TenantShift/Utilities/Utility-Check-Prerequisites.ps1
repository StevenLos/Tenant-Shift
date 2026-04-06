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
Shared prerequisite engine via TenantShift\Utilities\PrerequisiteEngine.psm1

.SYNOPSIS
    Scans repository PowerShell files for module prerequisites and reports
    local environment readiness against optional prerequisite profiles.

.DESCRIPTION
    Uses the shared prerequisite engine for module discovery and profile-based
    environment validation. By default, the scan excludes development test and
    build paths so the report focuses on runtime-facing scripts.

    Test-override environment variables (used by automated tests only):
      PESTER_VERSION_OVERRIDE            - "none", a version string, or unset
      PSSCRIPTANALYZER_VERSION_OVERRIDE  - "none", a version string, or unset
      MODULE_VERSION_OVERRIDE            - "present" to satisfy profile module checks
      PNP_CLIENT_ID_OVERRIDE             - "absent" or a client ID value
      POWERSHELL_VERSION_OVERRIDE        - a PowerShell version string
      PS_EDITION_OVERRIDE                - "Desktop" or "Core"
      IS_WINDOWS_OVERRIDE                - "true" or "false"
      EXCHANGE_MANAGEMENT_SHELL_OVERRIDE - "true" or "false"

.PARAMETER RepositoryRoot
    Repository root to scan for PowerShell scripts and modules.

.PARAMETER IncludeRelativePathPattern
    Relative-path wildcard patterns to include in the module report.

.PARAMETER ExcludeRelativePathPattern
    Relative-path wildcard patterns to exclude from discovery.

.PARAMETER Profile
    Optional prerequisite profile names to evaluate in addition to repository
    module discovery.

.PARAMETER SkipGalleryCheck
    Skips PSGallery version lookups and reports currency as Unknown.

.PARAMETER OutputFormat
    Console output mode when -PassThru is not used.

.PARAMETER FailOn
    Optional finding types that should produce exit code 1.

.PARAMETER PassThru
    Returns a structured report object and does not exit the host process.
#>
#Requires -Version 5.1

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),

    [string[]]$IncludeRelativePathPattern = @('*'),

    [string[]]$ExcludeRelativePathPattern = @(
        'TenantShift\Development\Tests\*'
        'TenantShift\Development\Build\*'
    ),

    [ValidateSet('Contributor', 'OnlineOperator', 'OnPremOperator', 'RepoScan')]
    [string[]]$Profile,

    [switch]$SkipGalleryCheck,

    [ValidateSet('Table', 'Json')]
    [string]$OutputFormat = 'Table',

    [ValidateSet('Missing', 'OutOfDate', 'ProfileFailure')]
    [string[]]$FailOn = @(),

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:repositoryRoot = (Resolve-Path -LiteralPath $RepositoryRoot).Path
$script:workspaceRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$script:enginePath = Join-Path -Path $script:workspaceRoot -ChildPath 'TenantShift\Utilities\PrerequisiteEngine.psm1'

if (-not (Test-Path -LiteralPath $script:enginePath)) {
    throw "Prerequisite engine not found: $script:enginePath"
}

Import-Module $script:enginePath -Force -ErrorAction Stop | Out-Null

function Write-Status {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $color = switch ($Level) {
        'INFO' { 'Cyan' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        'SUCCESS' { 'Green' }
    }

    Write-Host "[$Level] $Message" -ForegroundColor $color
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

    switch ($Value.Trim().ToLowerInvariant()) {
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

function Get-ProfileAuditResults {
    [CmdletBinding()]
    param(
        [string[]]$ProfileNames
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $resolvedProfileNames = @($ProfileNames | Select-Object -Unique)

    if ($resolvedProfileNames.Count -eq 0) {
        return @($results)
    }

    $catalog = Get-PrerequisiteCatalog
    $profileDefinitions = foreach ($profileName in $resolvedProfileNames) {
        Resolve-PrerequisiteProfile -ProfileName $profileName -Catalog $catalog
    }

    $hostOverride = Get-HostOverrideMap
    $moduleVersionOverride = Get-ModuleVersionOverrideMap -ProfileDefinitions $profileDefinitions
    $environmentVariableOverride = Get-EnvironmentVariableOverrideMap

    foreach ($profileDefinition in $profileDefinitions) {
        $state = Get-PrerequisiteState -ProfileDefinition $profileDefinition -Catalog $catalog `
            -HostOverride $hostOverride `
            -ModuleVersionOverride $moduleVersionOverride `
            -EnvironmentVariableOverride $environmentVariableOverride

        $checks = @(Test-PrerequisiteProfile -ProfileName $profileDefinition.Name -Catalog $catalog -State $state)
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

function Get-DiscoveredModuleResults {
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
                $galleryModule = Find-Module -Name $module.ModuleName -Repository PSGallery -ErrorAction Stop
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

function New-UtilityReport {
    [CmdletBinding()]
    param()

    $moduleResults = @(Get-DiscoveredModuleResults -RootPath $script:repositoryRoot `
            -IncludePatterns $IncludeRelativePathPattern `
            -ExcludePatterns $ExcludeRelativePathPattern `
            -SkipGalleryVersionCheck:$SkipGalleryCheck)

    $profileResults = @(Get-ProfileAuditResults -ProfileNames $Profile)
    $triggeredFailOn = @(Get-TriggeredFailOn -ModuleResults $moduleResults -ProfileResults $profileResults -FailCriteria $FailOn)

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
        RepositoryRoot             = $script:repositoryRoot
        IncludeRelativePathPattern = @($IncludeRelativePathPattern)
        ExcludeRelativePathPattern = @($ExcludeRelativePathPattern)
        RequestedProfiles          = @($Profile | Select-Object -Unique)
        DiscoveredModules          = $moduleResults
        ProfileResults             = $profileResults
        TriggeredFailOn            = $triggeredFailOn
        ShouldFail                 = ($triggeredFailOn.Count -gt 0)
        Summary                    = [pscustomobject]@{
            DiscoveredModules = $moduleResults.Count
            Present           = $presentCount
            Missing           = $missingCount
            UpToDate          = $upToDateCount
            OutOfDate         = $outOfDateCount
            UnknownCurrency   = $unknownCount
            ProfileFailures   = $profileFailureCount
            ProfileWarnings   = $profileWarningCount
        }
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
    Write-Host "=== $Title ===" -ForegroundColor Cyan

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

function Write-ProfileSection {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
    param(
        [Parameter(Mandatory)]
        [object[]]$ProfileResults
    )

    foreach ($profileResult in $ProfileResults) {
        Write-Host ''
        Write-Host ("=== Profile: {0} ===" -f $profileResult.ProfileName) -ForegroundColor Cyan

        if ($profileResult.Checks.Count -eq 0) {
            Write-Host 'No checks defined for this profile.' -ForegroundColor DarkGray
            continue
        }

        $rendered = $profileResult.Checks |
            Select-Object Status, CheckName, Actual, Expected |
            Format-Table -AutoSize |
            Out-String

        Write-Host ($rendered.TrimEnd())
    }
}

function Write-UtilityReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Report
    )

    Write-Status -Message "Repository root: $($Report.RepositoryRoot)"
    Write-Status -Message ("Include patterns: {0}" -f (@($Report.IncludeRelativePathPattern) -join ', '))
    Write-Status -Message ("Exclude patterns: {0}" -f (@($Report.ExcludeRelativePathPattern) -join ', '))

    Write-SectionTable -Title 'Modules Present' `
        -Rows @($Report.DiscoveredModules | Where-Object { $_.Present }) `
        -Columns @('ModuleName', 'InstalledVersion', 'GalleryVersion', 'CurrencyStatus', 'SourceCount') `
        -EmptyMessage 'No discovered modules are currently present.'

    Write-SectionTable -Title 'Modules Missing' `
        -Rows @($Report.DiscoveredModules | Where-Object { $_.CurrencyStatus -eq 'Missing' }) `
        -Columns @('ModuleName', 'GalleryVersion', 'InstallCommand', 'SourceCount') `
        -EmptyMessage 'No discovered modules are missing.'

    Write-SectionTable -Title 'Modules Out Of Date' `
        -Rows @($Report.DiscoveredModules | Where-Object { $_.CurrencyStatus -eq 'OutOfDate' }) `
        -Columns @('ModuleName', 'InstalledVersion', 'GalleryVersion', 'UpdateCommand', 'SourceCount') `
        -EmptyMessage 'No discovered modules are currently out of date.'

    Write-SectionTable -Title 'Modules With Unknown Currency' `
        -Rows @($Report.DiscoveredModules | Where-Object { $_.Present -and $_.CurrencyStatus -eq 'Unknown' }) `
        -Columns @('ModuleName', 'InstalledVersion', 'GalleryLookupStatus', 'SourceCount') `
        -EmptyMessage 'No discovered modules have unknown currency.'

    if ($Report.ProfileResults.Count -gt 0) {
        Write-ProfileSection -ProfileResults $Report.ProfileResults
    }

    Write-Host ''
    Write-Host '=== Summary ===' -ForegroundColor Cyan
    Write-Host ("Discovered modules: {0}" -f $Report.Summary.DiscoveredModules)
    Write-Host ("Present: {0}" -f $Report.Summary.Present)
    Write-Host ("Missing: {0}" -f $Report.Summary.Missing)
    Write-Host ("Up to date: {0}" -f $Report.Summary.UpToDate)
    Write-Host ("Out of date: {0}" -f $Report.Summary.OutOfDate)
    Write-Host ("Unknown currency: {0}" -f $Report.Summary.UnknownCurrency)
    Write-Host ("Profile failures: {0}" -f $Report.Summary.ProfileFailures)
    Write-Host ("Profile warnings: {0}" -f $Report.Summary.ProfileWarnings)

    if ($Report.ShouldFail) {
        Write-Status -Message ("FailOn triggered: {0}" -f ($Report.TriggeredFailOn -join ', ')) -Level WARN
    }
    else {
        Write-Status -Message 'Report completed with no triggered fail conditions.' -Level SUCCESS
    }
}

$report = New-UtilityReport

if ($PassThru) {
    return $report
}

switch ($OutputFormat) {
    'Json' {
        $report | ConvertTo-Json -Depth 8
    }
    default {
        Write-UtilityReport -Report $report
    }
}

if ($report.ShouldFail) {
    exit 1
}

exit 0
