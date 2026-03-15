<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260315-145500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
None declared in this file

.MODULEVERSIONPOLICY
Best-effort PSGallery check (offline-safe)
#>
#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),

    [switch]$SkipGalleryCheck
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

function Test-IsModuleName {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $false
    }

    return ($Name -match '^[A-Za-z0-9][A-Za-z0-9._-]*$')
}

function Get-PowerShellRuntimeInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string[]]$CommandCandidates,

        [Parameter(Mandatory)]
        [scriptblock]$RequirementTest,

        [Parameter(Mandatory)]
        [string]$InstallHint
    )

    $commandPath = $null
    foreach ($candidate in $CommandCandidates) {
        $cmd = Get-Command -Name $candidate -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $cmd) {
            $commandPath = $cmd.Source
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($commandPath)) {
        return [pscustomobject]@{
            Runtime          = $Name
            Present          = $false
            Version          = ''
            MeetsRequirement = $false
            CommandPath      = ''
            InstallHint      = $InstallHint
        }
    }

    $versionText = ''
    $parsedVersion = $null
    try {
        $versionText = (& $commandPath -NoProfile -Command '$PSVersionTable.PSVersion.ToString()' 2>$null | Select-Object -First 1)
        if (-not [string]::IsNullOrWhiteSpace($versionText)) {
            [void][version]::TryParse($versionText.Trim(), [ref]$parsedVersion)
        }
    }
    catch {
        $versionText = ''
    }

    $meets = $false
    if ($null -ne $parsedVersion) {
        $meets = [bool](& $RequirementTest $parsedVersion)
    }

    return [pscustomobject]@{
        Runtime          = $Name
        Present          = $true
        Version          = $versionText
        MeetsRequirement = $meets
        CommandPath      = $commandPath
        InstallHint      = ''
    }
}

function Get-RelativePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$BasePath,

        [Parameter(Mandatory)]
        [string]$FullPath
    )

    $base = (Resolve-Path -LiteralPath $BasePath).Path
    $full = (Resolve-Path -LiteralPath $FullPath).Path

    if ($full.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $full.Substring($base.Length).TrimStart([char[]]@([char]92, [char]47))
    }

    return $full
}

function Add-ModuleRequirement {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$ModuleIndex,

        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter(Mandatory)]
        [string]$SourcePath
    )

    if (-not (Test-IsModuleName -Name $ModuleName)) {
        return
    }

    if (-not $ModuleIndex.ContainsKey($ModuleName)) {
        $ModuleIndex[$ModuleName] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }

    [void]$ModuleIndex[$ModuleName].Add($SourcePath)
}

function Get-HeaderRequiredModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    $lines = Get-Content -LiteralPath $FilePath -TotalCount 180
    $inCommentHeader = $false
    $inRequiredModules = $false
    $modules = [System.Collections.Generic.List[string]]::new()

    foreach ($line in $lines) {
        if (-not $inCommentHeader -and $line -match '^\s*<#') {
            $inCommentHeader = $true
            continue
        }

        if ($inCommentHeader -and $line -match '^\s*#>') {
            break
        }

        if (-not $inCommentHeader) {
            continue
        }

        if ($line -match '^\s*\.REQUIREDMODULES\s*$') {
            $inRequiredModules = $true
            continue
        }

        if ($inRequiredModules -and $line -match '^\s*\.[A-Z]') {
            break
        }

        if (-not $inRequiredModules) {
            continue
        }

        $text = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }

        if ($text -match '^(None declared in this file|None|Not declared in this file)$') {
            continue
        }

        foreach ($part in ($text -split '[,;]')) {
            $candidate = $part.Trim()
            if (Test-IsModuleName -Name $candidate) {
                [void]$modules.Add($candidate)
            }
        }
    }

    return @($modules)
}

function Get-AstRequiredModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.Language.ScriptBlockAst]$Ast
    )

    $modules = [System.Collections.Generic.List[string]]::new()

    $scriptRequirements = $null
    if ($null -ne $Ast.PSObject.Properties['ScriptRequirements']) {
        $scriptRequirements = $Ast.ScriptRequirements
    }

    if ($null -eq $scriptRequirements -or $null -eq $scriptRequirements.PSObject.Properties['RequiredModules']) {
        return @($modules)
    }

    foreach ($required in @($scriptRequirements.RequiredModules)) {
        $candidate = $null

        if ($required -is [string]) {
            $candidate = $required
        }
        elseif ($required -is [hashtable]) {
            $candidate = [string]$required.ModuleName
        }
        elseif ($required -is [System.Collections.IDictionary]) {
            $candidate = [string]$required['ModuleName']
        }
        elseif ($null -ne $required -and $required.PSObject.Properties['Name']) {
            $candidate = [string]$required.Name
        }

        if (Test-IsModuleName -Name $candidate) {
            [void]$modules.Add($candidate)
        }
    }

    return @($modules)
}

function Get-AstImportModuleNames {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.Language.ScriptBlockAst]$Ast
    )

    $modules = [System.Collections.Generic.List[string]]::new()
    $importCalls = $Ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Import-Module'
        }, $true)

    foreach ($call in $importCalls) {
        $commandElements = @($call.CommandElements)
        if ($commandElements.Count -lt 2) {
            continue
        }

        $candidate = $null

        for ($i = 1; $i -lt $commandElements.Count; $i++) {
            $element = $commandElements[$i]
            if ($element -is [System.Management.Automation.Language.CommandParameterAst] -and $element.ParameterName -eq 'Name') {
                if ($i + 1 -lt $commandElements.Count -and $commandElements[$i + 1] -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    $candidate = $commandElements[$i + 1].Value
                    break
                }
            }
        }

        if ([string]::IsNullOrWhiteSpace($candidate)) {
            for ($i = 1; $i -lt $commandElements.Count; $i++) {
                $element = $commandElements[$i]
                if ($element -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    $candidate = $element.Value
                    break
                }
            }
        }

        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        if ($candidate.Contains('\\') -or $candidate.Contains('/') -or $candidate.EndsWith('.psm1', [System.StringComparison]::OrdinalIgnoreCase) -or $candidate.StartsWith('.')) {
            continue
        }

        if (Test-IsModuleName -Name $candidate) {
            [void]$modules.Add($candidate)
        }
    }

    return @($modules)
}

function Get-RequiredModuleIndex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RepositoryRoot
    )

    $index = @{}
    $scriptFiles = @(Get-ChildItem -LiteralPath $RepositoryRoot -Recurse -Include *.ps1,*.psm1 -File)

    foreach ($file in $scriptFiles) {
        $relativePath = Get-RelativePath -BasePath $RepositoryRoot -FullPath $file.FullName

        foreach ($moduleName in (Get-HeaderRequiredModules -FilePath $file.FullName)) {
            Add-ModuleRequirement -ModuleIndex $index -ModuleName $moduleName -SourcePath $relativePath
        }

        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)

        foreach ($moduleName in (Get-AstRequiredModules -Ast $ast)) {
            Add-ModuleRequirement -ModuleIndex $index -ModuleName $moduleName -SourcePath $relativePath
        }

        foreach ($moduleName in (Get-AstImportModuleNames -Ast $ast)) {
            Add-ModuleRequirement -ModuleIndex $index -ModuleName $moduleName -SourcePath $relativePath
        }
    }

    return $index
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

function Write-SectionTable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [AllowEmptyCollection()]
        [object[]]$Rows,

        [Parameter(Mandatory)]
        [string[]]$Columns,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$EmptyMessage
    )

    Write-Host ''
    Write-Host "=== $Title ===" -ForegroundColor Cyan

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        Write-Host $EmptyMessage -ForegroundColor DarkGray
        return
    }

    $rendered = $Rows |
        Sort-Object ModuleName |
        Select-Object -Property $Columns |
        Format-Table -AutoSize |
        Out-String

    Write-Host ($rendered.TrimEnd())
}

$resolvedRoot = (Resolve-Path -LiteralPath $RepositoryRoot).Path
Write-Status -Message "Repository root: $resolvedRoot"

$runtimeChecks = @(
    Get-PowerShellRuntimeInfo -Name 'PowerShell 5.1' -CommandCandidates @('powershell.exe', 'powershell') -RequirementTest { param([version]$v) $v.Major -eq 5 -and $v.Minor -eq 1 } -InstallHint 'Windows PowerShell 5.1 should be present on supported Windows builds.'
    Get-PowerShellRuntimeInfo -Name 'PowerShell 7+' -CommandCandidates @('pwsh.exe', 'pwsh') -RequirementTest { param([version]$v) $v.Major -ge 7 } -InstallHint 'Install PowerShell 7+: winget install Microsoft.PowerShell'
)

$runtimeTable = $runtimeChecks |
    Select-Object Runtime, Present, Version, MeetsRequirement, CommandPath, InstallHint |
    Format-Table -AutoSize |
    Out-String

Write-Host ''
Write-Host '=== PowerShell Runtime Check ===' -ForegroundColor Cyan
Write-Host ($runtimeTable.TrimEnd())

if (@($runtimeChecks.Where({ -not $_.MeetsRequirement })).Count -gt 0) {
    Write-Status -Message 'One or more runtime requirements are not met.' -Level WARN
}
else {
    Write-Status -Message 'PowerShell runtime requirements are met.' -Level SUCCESS
}

Write-Status -Message 'Discovering required modules from script metadata and imports.'
$moduleIndex = Get-RequiredModuleIndex -RepositoryRoot $resolvedRoot
$requiredModuleNames = @($moduleIndex.Keys | Sort-Object)

if ($requiredModuleNames.Count -eq 0) {
    Write-Status -Message 'No required modules discovered in this repository.' -Level WARN
    return
}

Write-Status -Message ("Discovered {0} unique required module(s)." -f $requiredModuleNames.Count)

$moduleResults = [System.Collections.Generic.List[object]]::new()
foreach ($moduleName in $requiredModuleNames) {
    $installed = @(Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1)
    $isPresent = $installed.Count -gt 0
    $installedVersion = if ($isPresent) { [string]$installed[0].Version } else { '' }

    $galleryVersion = ''
    $galleryLookupStatus = 'Skipped'
    $inPsGallery = $false

    if (-not $SkipGalleryCheck) {
        try {
            $galleryModule = Find-Module -Name $moduleName -Repository PSGallery -ErrorAction Stop
            $galleryVersion = [string]$galleryModule.Version
            $inPsGallery = $true
            $galleryLookupStatus = 'Found'
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

    $requiredBy = @($moduleIndex[$moduleName] | ForEach-Object { $_ } | Sort-Object)
    $requiredByDisplay = if ($requiredBy.Count -le 3) {
        ($requiredBy -join '; ')
    }
    else {
        (($requiredBy[0..2] -join '; ') + "; +$($requiredBy.Count - 3) more")
    }

    $installCommand = ''
    $updateCommand = ''

    if (-not $isPresent) {
        $installCommand = Get-InstallCommand -ModuleName $moduleName -InPsGallery:$inPsGallery
    }

    if ($currencyStatus -eq 'OutOfDate') {
        $updateCommand = Get-UpdateCommand -ModuleName $moduleName -InPsGallery:$inPsGallery
    }

    $moduleResults.Add([pscustomobject]@{
            ModuleName          = $moduleName
            Present             = $isPresent
            InstalledVersion    = $installedVersion
            GalleryVersion      = $galleryVersion
            GalleryLookupStatus = $galleryLookupStatus
            CurrencyStatus      = $currencyStatus
            RequiredByCount     = $requiredBy.Count
            RequiredBy          = $requiredByDisplay
            InstallCommand      = $installCommand
            UpdateCommand       = $updateCommand
        }) | Out-Null
}

$present = @($moduleResults.Where({ $_.Present }))
$missing = @($moduleResults.Where({ -not $_.Present }))
$upToDate = @($moduleResults.Where({ $_.CurrencyStatus -eq 'UpToDate' }))
$outOfDate = @($moduleResults.Where({ $_.CurrencyStatus -eq 'OutOfDate' }))
$unknownCurrency = @($moduleResults.Where({ $_.Present -and $_.CurrencyStatus -eq 'Unknown' }))

Write-SectionTable -Title 'Modules Present' -Rows $present -Columns @('ModuleName', 'InstalledVersion', 'GalleryVersion', 'CurrencyStatus', 'RequiredByCount') -EmptyMessage 'No modules are currently present.'
Write-SectionTable -Title 'Modules Missing' -Rows $missing -Columns @('ModuleName', 'GalleryVersion', 'InstallCommand', 'RequiredByCount') -EmptyMessage 'No modules are missing.'
Write-SectionTable -Title 'Modules Up To Date' -Rows $upToDate -Columns @('ModuleName', 'InstalledVersion', 'GalleryVersion', 'RequiredByCount') -EmptyMessage 'No modules are confirmed up to date.'
Write-SectionTable -Title 'Modules Out Of Date' -Rows $outOfDate -Columns @('ModuleName', 'InstalledVersion', 'GalleryVersion', 'UpdateCommand', 'RequiredByCount') -EmptyMessage 'No modules are currently out of date.'

if ($unknownCurrency.Count -gt 0) {
    Write-SectionTable -Title 'Modules With Unknown Currency' -Rows $unknownCurrency -Columns @('ModuleName', 'InstalledVersion', 'GalleryLookupStatus', 'RequiredByCount') -EmptyMessage ''
}

Write-Host ''
Write-Host '=== Summary ===' -ForegroundColor Cyan
Write-Host ("Required modules discovered: {0}" -f $requiredModuleNames.Count)
Write-Host ("Present: {0}" -f $present.Count)
Write-Host ("Missing: {0}" -f $missing.Count)
Write-Host ("Up to date: {0}" -f $upToDate.Count)
Write-Host ("Out of date: {0}" -f $outOfDate.Count)
Write-Host ("Unknown currency: {0}" -f $unknownCurrency.Count)

if ($missing.Count -eq 0 -and $outOfDate.Count -eq 0 -and @($runtimeChecks.Where({ -not $_.MeetsRequirement })).Count -eq 0) {
    Write-Status -Message 'Environment checks completed with no blocking findings.' -Level SUCCESS
}
else {
    Write-Status -Message 'Environment checks completed with findings to address.' -Level WARN
}
