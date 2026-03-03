<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-000002

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
None declared in this file

.MODULEVERSIONPOLICY
Not declared in this file
#>
#Requires -Version 7.0

[CmdletBinding()]
param(
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot

$issues = [System.Collections.Generic.List[object]]::new()

function Add-Issue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Type,

        [Parameter(Mandatory)]
        [string]$File,

        [Parameter(Mandatory)]
        [string]$Details
    )

    $issues.Add([pscustomobject]@{
        Type    = $Type
        File    = $File
        Details = $Details
    })
}

function Get-RelativePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return [System.IO.Path]::GetRelativePath($repoRoot, $Path)
}

function Get-RequiredHeadersFromScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Content
    )

    $requiredMatch = [regex]::Match(
        $Content,
        '\$requiredHeaders\s*=\s*@\((?<arr>[\s\S]*?)\)\s*(?:\r?\n\r?\n|Write-Status|Ensure-|if\s*\(|\$rows\s*=)',
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    if (-not $requiredMatch.Success) {
        return @()
    }

    $arr = $requiredMatch.Groups['arr'].Value
    return @(
        [regex]::Matches($arr, '''([^'']+)''|"([^"]+)"') |
            ForEach-Object {
                if ($_.Groups[1].Success) {
                    $_.Groups[1].Value
                }
                else {
                    $_.Groups[2].Value
                }
            }
    )
}

$allScripts = Get-ChildItem -Path (Join-Path $repoRoot 'OnPrem'), (Join-Path $repoRoot 'Online'), (Join-Path $repoRoot 'Common'), (Join-Path $repoRoot 'Build') -Recurse -Include '*.ps1', '*.psm1'

$requiredTags = @(
    '.SCRIPTVERSION',
    '.POWERSHELLREQUIRED',
    '.REQUIREDMODULES',
    '.MODULEVERSIONPOLICY'
)

foreach ($script in $allScripts) {
    $content = Get-Content -LiteralPath $script.FullName -Raw

    foreach ($tag in $requiredTags) {
        if ($content -notmatch [regex]::Escape($tag)) {
            Add-Issue -Type 'MissingMetadataTag' -File (Get-RelativePath -Path $script.FullName) -Details "Missing $tag"
        }
    }
}

$operationScripts = Get-ChildItem -Path (Join-Path $repoRoot 'OnPrem'), (Join-Path $repoRoot 'Online') -Recurse -Filter '*.ps1' |
    Where-Object { $_.Name -match '^(P|M|IR)\d{4}-' }

foreach ($script in $operationScripts) {
    $content = Get-Content -LiteralPath $script.FullName -Raw
    $relativePath = Get-RelativePath -Path $script.FullName

    if ($content -notmatch '#Requires\s+-Version') {
        Add-Issue -Type 'MissingRequiresVersion' -File $relativePath -Details 'Missing #Requires -Version declaration.'
    }

    if ($content -notmatch '\[CmdletBinding') {
        Add-Issue -Type 'MissingCmdletBinding' -File $relativePath -Details 'Missing [CmdletBinding()] declaration.'
    }

    $requiredHeaders = Get-RequiredHeadersFromScript -Content $content
    if (@($requiredHeaders).Count -eq 0) {
        continue
    }

    $inputCsvPath = Join-Path -Path $script.Directory.FullName -ChildPath ($script.BaseName + '.input.csv')
    if (-not (Test-Path -LiteralPath $inputCsvPath)) {
        continue
    }

    $rows = @(Import-Csv -LiteralPath $inputCsvPath)
    if (@($rows).Count -eq 0) {
        Add-Issue -Type 'EmptyTemplate' -File (Get-RelativePath -Path $inputCsvPath) -Details 'CSV template has no data rows.'
        continue
    }

    $csvHeaders = @($rows[0].PSObject.Properties.Name)
    $missingHeaders = @($requiredHeaders | Where-Object { $_ -notin $csvHeaders })

    if (@($missingHeaders).Count -gt 0) {
        Add-Issue -Type 'MissingRequiredHeaders' -File (Get-RelativePath -Path $inputCsvPath) -Details ("Missing headers: {0}" -f ($missingHeaders -join ', '))
    }
}

$csvTemplates = Get-ChildItem -Path (Join-Path $repoRoot 'OnPrem'), (Join-Path $repoRoot 'Online') -Recurse -Filter '*.input.csv'
foreach ($template in $csvTemplates) {
    try {
        $rows = @(Import-Csv -LiteralPath $template.FullName)
        if (@($rows).Count -eq 0) {
            Add-Issue -Type 'EmptyTemplate' -File (Get-RelativePath -Path $template.FullName) -Details 'CSV template has no data rows.'
        }
    }
    catch {
        Add-Issue -Type 'TemplateParseError' -File (Get-RelativePath -Path $template.FullName) -Details $_.Exception.Message
    }
}

if ($issues.Count -gt 0) {
    $issues |
        Sort-Object Type, File |
        Format-Table -AutoSize |
        Out-String |
        Write-Host

    throw "Repository contract validation failed with $($issues.Count) issue(s)."
}

Write-Host 'Repository contract validation passed.' -ForegroundColor Green

if ($PassThru) {
    return $issues.ToArray()
}
