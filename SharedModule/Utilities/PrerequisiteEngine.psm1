<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260401-000000

.SYNOPSIS
    Shared prerequisite catalog and discovery engine for SharedModule.

.DESCRIPTION
    Provides profile-based prerequisite evaluation and repository module
    discovery helpers. This module is the shared foundation for environment
    validation and prerequisite reporting tools.

.POWERSHELLREQUIRED
    5.1+ (PS 7 supported; no PS7-exclusive syntax used)

.REQUIREDMODULES
    None

.MODULEVERSIONPOLICY
    Version managed via PrerequisiteCatalog.psd1 (CatalogVersion = 1.0.0)
#>
#Requires -Version 5.1

Set-StrictMode -Version Latest

function ConvertTo-NormalizedVersion {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [version]) {
        return $Value
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $parsed = $null
    if ([version]::TryParse($text, [ref]$parsed)) {
        return $parsed
    }

    return $null
}

function ConvertTo-DisplayValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return '<not detected>'
    }

    if ($Value -is [version]) {
        return $Value.ToString()
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @()
        foreach ($item in $Value) {
            $items += (ConvertTo-DisplayValue -Value $item)
        }

        if ($items.Count -eq 0) {
            return '<empty>'
        }

        return ($items -join '; ')
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return '<empty>'
    }

    return $text
}

function ConvertTo-CheckDisplayValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory)]
        [hashtable]$Check
    )

    if ($Check.ContainsKey('ActualValueMap') -and $null -ne $Check.ActualValueMap -and $null -ne $Value) {
        $mapKey = [string]$Value
        if ($Check.ActualValueMap.ContainsKey($mapKey)) {
            return [string]$Check.ActualValueMap[$mapKey]
        }
    }

    return ConvertTo-DisplayValue -Value $Value
}

function Test-ModuleToken {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $false
    }

    if ($Name.Contains('\') -or $Name.Contains('/') -or $Name.StartsWith('.')) {
        return $false
    }

    if ($Name.EndsWith('.psm1', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }

    return ($Name -match '^[A-Za-z0-9][A-Za-z0-9._-]*$')
}

function Resolve-RepositoryRelativePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$BasePath,

        [Parameter(Mandatory)]
        [string]$FullPath
    )

    $resolvedBase = (Resolve-Path -LiteralPath $BasePath).Path
    $resolvedFull = (Resolve-Path -LiteralPath $FullPath).Path

    if ($resolvedFull.StartsWith($resolvedBase, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $resolvedFull.Substring($resolvedBase.Length).TrimStart([char[]]@([char]92, [char]47))
    }

    return $resolvedFull
}

function Add-DiscoveredModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$ModuleIndex,

        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter(Mandatory)]
        [string]$SourcePath
    )

    if (-not (Test-ModuleToken -Name $ModuleName)) {
        return
    }

    if (-not $ModuleIndex.ContainsKey($ModuleName)) {
        $ModuleIndex[$ModuleName] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }

    [void]$ModuleIndex[$ModuleName].Add($SourcePath)
}

function Get-HeaderRequiredModuleNames {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    $lines = Get-Content -LiteralPath $FilePath -TotalCount 180
    $inHeader = $false
    $inRequiredModules = $false
    $modules = [System.Collections.Generic.List[string]]::new()

    foreach ($line in $lines) {
        if (-not $inHeader -and $line -match '^\s*<#') {
            $inHeader = $true
            continue
        }

        if ($inHeader -and $line -match '^\s*#>') {
            break
        }

        if (-not $inHeader) {
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
            if (Test-ModuleToken -Name $candidate) {
                [void]$modules.Add($candidate)
            }
        }
    }

    return @($modules)
}

function Get-AstRequiredModuleNames {
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

        if (Test-ModuleToken -Name $candidate) {
            [void]$modules.Add($candidate)
        }
    }

    return @($modules)
}

function Get-ConstantModuleNamesFromAstValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [System.Management.Automation.Language.Ast]$AstValue
    )

    $modules = [System.Collections.Generic.List[string]]::new()

    if ($null -eq $AstValue) {
        return @($modules)
    }

    $value = $null
    try {
        $value = $AstValue.SafeGetValue()
    }
    catch {
        return @($modules)
    }

    if ($null -eq $value) {
        return @($modules)
    }

    if ($value -is [string]) {
        if (Test-ModuleToken -Name $value) {
            [void]$modules.Add($value)
        }

        return @($modules)
    }

    if ($value -is [System.Collections.IEnumerable]) {
        foreach ($item in $value) {
            if ($item -is [string] -and (Test-ModuleToken -Name $item)) {
                [void]$modules.Add($item)
            }
        }
    }

    return @($modules)
}

function Get-AstAssertModuleCurrentNames {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.Language.ScriptBlockAst]$Ast
    )

    $modules = [System.Collections.Generic.List[string]]::new()
    $calls = $Ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Assert-ModuleCurrent'
        }, $true)

    foreach ($call in $calls) {
        $elements = @($call.CommandElements)
        $argumentAst = $null

        for ($i = 1; $i -lt $elements.Count; $i++) {
            if ($elements[$i] -is [System.Management.Automation.Language.CommandParameterAst] -and
                $elements[$i].ParameterName -eq 'ModuleNames') {
                if ($i + 1 -lt $elements.Count) {
                    $argumentAst = $elements[$i + 1]
                }
                break
            }
        }

        if ($null -eq $argumentAst -and $elements.Count -ge 2 -and $elements[1] -isnot [System.Management.Automation.Language.CommandParameterAst]) {
            $argumentAst = $elements[1]
        }

        foreach ($moduleName in (Get-ConstantModuleNamesFromAstValue -AstValue $argumentAst)) {
            [void]$modules.Add($moduleName)
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
    $calls = $Ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Import-Module'
        }, $true)

    foreach ($call in $calls) {
        $elements = @($call.CommandElements)
        if ($elements.Count -lt 2) {
            continue
        }

        $moduleArgumentAst = $null

        for ($i = 1; $i -lt $elements.Count; $i++) {
            if ($elements[$i] -is [System.Management.Automation.Language.CommandParameterAst] -and
                $elements[$i].ParameterName -eq 'Name') {
                if ($i + 1 -lt $elements.Count) {
                    $moduleArgumentAst = $elements[$i + 1]
                }
                break
            }
        }

        if ($null -eq $moduleArgumentAst -and $elements[1] -isnot [System.Management.Automation.Language.CommandParameterAst]) {
            $moduleArgumentAst = $elements[1]
        }

        foreach ($moduleName in (Get-ConstantModuleNamesFromAstValue -AstValue $moduleArgumentAst)) {
            [void]$modules.Add($moduleName)
        }
    }

    return @($modules)
}

function Get-ExpectedValueDescription {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Check
    )

    if ($Check.ContainsKey('ExpectedDescription') -and -not [string]::IsNullOrWhiteSpace([string]$Check.ExpectedDescription)) {
        return [string]$Check.ExpectedDescription
    }

    switch ($Check.RequirementType) {
        'VersionRange' {
            $parts = [System.Collections.Generic.List[string]]::new()

            if ($Check.ContainsKey('MinimumVersion')) {
                [void]$parts.Add((">= {0}" -f $Check.MinimumVersion))
            }

            if ($Check.ContainsKey('MaximumVersionExclusive')) {
                [void]$parts.Add(("< {0}" -f $Check.MaximumVersionExclusive))
            }

            return ($parts -join ', ')
        }
        'AllowedValues' {
            return (@($Check.AllowedValues) -join ', ')
        }
        'NonEmpty' {
            return 'configured'
        }
        'Present' {
            return 'present'
        }
        default {
            return [string]$Check.RequirementType
        }
    }
}

function Get-ActualCheckValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Check,

        [Parameter(Mandatory)]
        [pscustomobject]$State
    )

    switch ($Check.Category) {
        'Runtime' {
            return $State.HostFacts[$Check.FactName]
        }
        'Module' {
            return $State.ModuleVersions[$Check.FactName]
        }
        'Configuration' {
            return $State.EnvironmentVariables[$Check.FactName]
        }
        default {
            throw "Unsupported prerequisite category '$($Check.Category)'."
        }
    }
}

function Test-CheckRequirement {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Check,

        [AllowNull()]
        [object]$ActualValue
    )

    switch ($Check.RequirementType) {
        'VersionRange' {
            $actualVersion = ConvertTo-NormalizedVersion -Value $ActualValue
            if ($null -eq $actualVersion) {
                return $false
            }

            if ($Check.ContainsKey('MinimumVersion')) {
                $minimumVersion = ConvertTo-NormalizedVersion -Value $Check.MinimumVersion
                if ($null -ne $minimumVersion -and $actualVersion -lt $minimumVersion) {
                    return $false
                }
            }

            if ($Check.ContainsKey('MaximumVersionExclusive')) {
                $maximumVersion = ConvertTo-NormalizedVersion -Value $Check.MaximumVersionExclusive
                if ($null -ne $maximumVersion -and $actualVersion -ge $maximumVersion) {
                    return $false
                }
            }

            return $true
        }
        'AllowedValues' {
            return (@($Check.AllowedValues) -contains $ActualValue)
        }
        'NonEmpty' {
            return (-not [string]::IsNullOrWhiteSpace([string]$ActualValue))
        }
        'Present' {
            if ($null -eq $ActualValue) {
                return $false
            }

            if ($ActualValue -is [string]) {
                return (-not [string]::IsNullOrWhiteSpace($ActualValue))
            }

            if ($ActualValue -is [bool]) {
                return $ActualValue
            }

            return $true
        }
        default {
            throw "Unsupported prerequisite requirement type '$($Check.RequirementType)'."
        }
    }
}

function Get-CurrentHostFacts {
    [CmdletBinding()]
    param()

    $isWindowsHost = $false
    if (Get-Variable -Name IsWindows -ErrorAction SilentlyContinue) {
        $isWindowsHost = [bool]$IsWindows
    }
    else {
        $isWindowsHost = ($env:OS -eq 'Windows_NT')
    }

    $psEdition = 'Desktop'
    if ($PSVersionTable.PSObject.Properties.Name -contains 'PSEdition' -and
        -not [string]::IsNullOrWhiteSpace([string]$PSVersionTable.PSEdition)) {
        $psEdition = [string]$PSVersionTable.PSEdition
    }
    elseif ($PSVersionTable.PSVersion.Major -ge 6) {
        $psEdition = 'Core'
    }

    $exchangeManagementShellAvailable = $true
    foreach ($commandName in @('Get-Recipient', 'Get-MailContact', 'Get-DistributionGroup', 'Get-DynamicDistributionGroup')) {
        if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) {
            $exchangeManagementShellAvailable = $false
            break
        }
    }

    return @{
        PowerShellVersion      = $PSVersionTable.PSVersion
        PSEdition              = $psEdition
        IsWindows              = $isWindowsHost
        ExchangeManagementShell = $exchangeManagementShellAvailable
    }
}

function Get-PrerequisiteCatalog {
    [CmdletBinding()]
    param(
        [string]$CatalogPath = (Join-Path -Path $PSScriptRoot -ChildPath 'PrerequisiteCatalog.psd1')
    )

    if (-not (Test-Path -LiteralPath $CatalogPath)) {
        throw "Prerequisite catalog not found: $CatalogPath"
    }

    $catalog = Import-PowerShellDataFile -LiteralPath $CatalogPath
    if (-not $catalog.ContainsKey('Profiles')) {
        throw "Prerequisite catalog '$CatalogPath' does not contain a Profiles section."
    }

    return $catalog
}

function Resolve-PrerequisiteProfile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProfileName,

        [hashtable]$Catalog = (Get-PrerequisiteCatalog)
    )

    if (-not $Catalog.ContainsKey('Profiles')) {
        throw 'Catalog does not contain a Profiles section.'
    }

    $profiles = $Catalog.Profiles
    if (-not $profiles.ContainsKey($ProfileName)) {
        throw "Unknown prerequisite profile '$ProfileName'."
    }

    $profile = $profiles[$ProfileName]
    return [pscustomobject]@{
        Name        = $ProfileName
        Description = [string]$profile.Description
        Checks      = @($profile.Checks)
    }
}

function Get-PrerequisiteState {
    [CmdletBinding()]
    param(
        [string]$ProfileName,

        [object]$ProfileDefinition,

        [hashtable]$Catalog = (Get-PrerequisiteCatalog),

        [hashtable]$HostOverride,

        [hashtable]$ModuleVersionOverride,

        [hashtable]$EnvironmentVariableOverride
    )

    if ($null -eq $ProfileDefinition) {
        if ([string]::IsNullOrWhiteSpace($ProfileName)) {
            throw 'Either ProfileName or ProfileDefinition must be provided.'
        }

        $ProfileDefinition = Resolve-PrerequisiteProfile -ProfileName $ProfileName -Catalog $Catalog
    }

    $hostFacts = Get-CurrentHostFacts
    if ($null -ne $HostOverride) {
        foreach ($key in $HostOverride.Keys) {
            $hostFacts[$key] = $HostOverride[$key]
        }
    }

    $moduleVersions = @{}
    $environmentVariables = @{}

    foreach ($check in @($ProfileDefinition.Checks)) {
        switch ($check.Category) {
            'Module' {
                if ($moduleVersions.ContainsKey($check.FactName)) {
                    continue
                }

                if ($null -ne $ModuleVersionOverride -and $ModuleVersionOverride.ContainsKey($check.FactName)) {
                    $moduleVersions[$check.FactName] = ConvertTo-NormalizedVersion -Value $ModuleVersionOverride[$check.FactName]
                    continue
                }

                $module = Get-Module -ListAvailable -Name $check.FactName |
                    Sort-Object Version -Descending |
                    Select-Object -First 1

                if ($null -eq $module) {
                    $moduleVersions[$check.FactName] = $null
                }
                else {
                    $moduleVersions[$check.FactName] = ConvertTo-NormalizedVersion -Value $module.Version
                }
            }
            'Configuration' {
                if ($environmentVariables.ContainsKey($check.FactName)) {
                    continue
                }

                if ($null -ne $EnvironmentVariableOverride -and $EnvironmentVariableOverride.ContainsKey($check.FactName)) {
                    $environmentVariables[$check.FactName] = $EnvironmentVariableOverride[$check.FactName]
                    continue
                }

                $environmentVariables[$check.FactName] = [Environment]::GetEnvironmentVariable($check.FactName)
            }
        }
    }

    return [pscustomobject]@{
        HostFacts            = $hostFacts
        ModuleVersions       = $moduleVersions
        EnvironmentVariables = $environmentVariables
    }
}

function Test-PrerequisiteProfile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProfileName,

        [hashtable]$Catalog = (Get-PrerequisiteCatalog),

        [pscustomobject]$State
    )

    $profile = Resolve-PrerequisiteProfile -ProfileName $ProfileName -Catalog $Catalog
    if ($null -eq $State) {
        $State = Get-PrerequisiteState -ProfileDefinition $profile -Catalog $Catalog
    }

    $results = [System.Collections.Generic.List[object]]::new()

    foreach ($check in @($profile.Checks)) {
        $actualValue = Get-ActualCheckValue -Check $check -State $State
        $passed = Test-CheckRequirement -Check $check -ActualValue $actualValue
        $severity = if ([string]::IsNullOrWhiteSpace([string]$check.Severity)) { 'Error' } else { [string]$check.Severity }
        $displayName = if ($check.ContainsKey('DisplayName') -and -not [string]::IsNullOrWhiteSpace([string]$check.DisplayName)) {
            [string]$check.DisplayName
        }
        else {
            [string]$check.Name
        }
        $expected = Get-ExpectedValueDescription -Check $check
        $actualDisplay = ConvertTo-CheckDisplayValue -Value $actualValue -Check $check

        $status = 'PASS'
        if (-not $passed) {
            if ($severity -eq 'Warning') {
                $status = 'WARN'
            }
            else {
                $status = 'FAIL'
            }
        }

        $results.Add([pscustomobject]@{
                ProfileName     = $profile.Name
                CheckName       = [string]$check.Name
                DisplayName     = $displayName
                Category        = [string]$check.Category
                RequirementType = [string]$check.RequirementType
                Status          = $status
                Severity        = $severity
                Passed          = $passed
                Expected        = $expected
                Actual          = $actualDisplay
                Detail          = ("Expected {0}; detected {1}." -f $expected, $actualDisplay)
                Remediation     = [string]$check.Remediation
            }) | Out-Null
    }

    return @($results)
}

function Find-ModuleRequirements {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RootPath,

        [string[]]$ExcludeRelativePathPattern = @()
    )

    $resolvedRoot = (Resolve-Path -LiteralPath $RootPath).Path
    $moduleIndex = @{}
    $scriptFiles = @(Get-ChildItem -LiteralPath $resolvedRoot -Recurse -Include '*.ps1', '*.psm1' -File)

    foreach ($file in $scriptFiles) {
        $relativePath = Resolve-RepositoryRelativePath -BasePath $resolvedRoot -FullPath $file.FullName

        $isExcluded = $false
        foreach ($pattern in @($ExcludeRelativePathPattern)) {
            if (-not [string]::IsNullOrWhiteSpace($pattern) -and $relativePath -like $pattern) {
                $isExcluded = $true
                break
            }
        }

        if ($isExcluded) {
            continue
        }

        foreach ($moduleName in (Get-HeaderRequiredModuleNames -FilePath $file.FullName)) {
            Add-DiscoveredModule -ModuleIndex $moduleIndex -ModuleName $moduleName -SourcePath $relativePath
        }

        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)

        foreach ($moduleName in (Get-AstRequiredModuleNames -Ast $ast)) {
            Add-DiscoveredModule -ModuleIndex $moduleIndex -ModuleName $moduleName -SourcePath $relativePath
        }

        foreach ($moduleName in (Get-AstAssertModuleCurrentNames -Ast $ast)) {
            Add-DiscoveredModule -ModuleIndex $moduleIndex -ModuleName $moduleName -SourcePath $relativePath
        }

        foreach ($moduleName in (Get-AstImportModuleNames -Ast $ast)) {
            Add-DiscoveredModule -ModuleIndex $moduleIndex -ModuleName $moduleName -SourcePath $relativePath
        }
    }

    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($moduleName in @($moduleIndex.Keys | Sort-Object)) {
        $sources = @($moduleIndex[$moduleName] | ForEach-Object { $_ } | Sort-Object)
        $results.Add([pscustomobject]@{
                ModuleName  = $moduleName
                SourceCount = $sources.Count
                Sources     = $sources
            }) | Out-Null
    }

    return @($results)
}

Export-ModuleMember -Function @(
    'Find-ModuleRequirements'
    'Get-PrerequisiteCatalog'
    'Get-PrerequisiteState'
    'Resolve-PrerequisiteProfile'
    'Test-PrerequisiteProfile'
)
