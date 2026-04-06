<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260406-000000

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
(none — post-processing script, no AD or module calls)

.MODULEVERSIONPOLICY
N/A

.SYNOPSIS
    Merges D-ADGP-0010 through D-ADGP-0040 outputs into a single flat export for Excel/Power BI.

.DESCRIPTION
    Post-processing script. Merges the output CSVs from:
    - D-ADGP-0010-Get-GroupPolicyObjects.ps1  (GPO metadata)
    - D-ADGP-0020-Get-GroupPolicyLinks.ps1    (GPO link details)
    - D-ADGP-0030-Get-GroupPolicyScopeTree.ps1 (tree key, scope depth, precedence)
    - D-ADGP-0040-Get-GroupPolicySummary.ps1  (risk tier, empty-GPO flag, link counts)

    Produces a single flat CSV with one row per GPO-scope link, combining all metadata fields.
    Designed for Excel pivot table analysis and Power BI ingestion.
    No Active Directory calls are made — this script is a pure CSV transformer.

.PARAMETER GpoObjectsCsvPath
    Path to the output CSV from D-ADGP-0010-Get-GroupPolicyObjects.ps1.

.PARAMETER LinksCsvPath
    Path to the output CSV from D-ADGP-0020-Get-GroupPolicyLinks.ps1.

.PARAMETER ScopeTreeCsvPath
    Optional path to the output CSV from D-ADGP-0030-Get-GroupPolicyScopeTree.ps1.
    When provided, TreeKey, ScopeDepth, and GpoPrecedence are included in the output.

.PARAMETER SummaryCsvPath
    Optional path to the output CSV from D-ADGP-0040-Get-GroupPolicySummary.ps1.
    When provided, RiskTier and IsEmptyGpo are included in the output.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-ADGP-0050-Get-GroupPolicyUnifiedView.ps1 `
        -GpoObjectsCsvPath .\Results_D-ADGP-0010*.csv `
        -LinksCsvPath .\Results_D-ADGP-0020*.csv `
        -ScopeTreeCsvPath .\Results_D-ADGP-0030*.csv `
        -SummaryCsvPath .\Results_D-ADGP-0040*.csv

    Build the unified view from all four prior script outputs.

.NOTES
    Version:          1.0
    Required modules: None (pure CSV post-processor)
    Required roles:   None (reads local CSV files only)
    Limitations:      Requires completed D-ADGP-0010 and D-ADGP-0020 outputs at minimum.
                      D-ADGP-0030 and D-ADGP-0040 outputs are optional; columns are empty when not provided.
                      Run scripts in order: 0010 -> 0020 -> 0030 + 0040 -> 0050.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$GpoObjectsCsvPath,

    [Parameter(Mandatory)]
    [string]$LinksCsvPath,

    [string]$ScopeTreeCsvPath,

    [string]$SummaryCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-ADGP-0050-Get-GroupPolicyUnifiedView_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

Write-Status -Message 'Starting Group Policy unified view builder.'

# Validate required inputs.
if (-not (Test-Path -LiteralPath $GpoObjectsCsvPath)) { throw "GPO objects CSV not found: $GpoObjectsCsvPath" }
if (-not (Test-Path -LiteralPath $LinksCsvPath))      { throw "Links CSV not found: $LinksCsvPath" }

# --- Load D-ADGP-0010: GPO metadata keyed by GpoGuid ---
Write-Status -Message "Reading GPO objects from: $GpoObjectsCsvPath"
$gpoMetaMap = @{}
foreach ($row in (Import-Csv -LiteralPath $GpoObjectsCsvPath | Where-Object { [string]$_.Status -eq 'Completed' })) {
    $guid = ([string]$row.GpoGuid).Trim().ToLowerInvariant()
    if (-not [string]::IsNullOrWhiteSpace($guid) -and -not $gpoMetaMap.ContainsKey($guid)) {
        $gpoMetaMap[$guid] = $row
    }
}
Write-Status -Message "Loaded $($gpoMetaMap.Count) GPO metadata rows."

# --- Load D-ADGP-0020: link rows ---
Write-Status -Message "Reading GPO links from: $LinksCsvPath"
$linkRows = @(Import-Csv -LiteralPath $LinksCsvPath | Where-Object { [string]$_.Status -eq 'Completed' })
Write-Status -Message "Loaded $($linkRows.Count) link rows."

# --- Load D-ADGP-0030: scope tree, keyed by ScopeDn|GpoGuid ---
$treeMap = @{}
$resolvedScopeTreeCsv = if ($PSBoundParameters.ContainsKey('ScopeTreeCsvPath')) { $ScopeTreeCsvPath.Trim() } else { '' }
if (-not [string]::IsNullOrWhiteSpace($resolvedScopeTreeCsv) -and (Test-Path -LiteralPath $resolvedScopeTreeCsv)) {
    Write-Status -Message "Reading scope tree from: $resolvedScopeTreeCsv"
    foreach ($row in (Import-Csv -LiteralPath $resolvedScopeTreeCsv)) {
        $treeKey = ("{0}|{1}" -f ([string]$row.ScopeDn).Trim(), ([string]$row.GpoGuid).Trim().ToLowerInvariant())
        $treeMap[$treeKey] = $row
    }
    Write-Status -Message "Loaded $($treeMap.Count) scope tree rows."
}

# --- Load D-ADGP-0040: per-GPO summary, keyed by GpoGuid ---
$summaryMap = @{}
$resolvedSummaryCsv = if ($PSBoundParameters.ContainsKey('SummaryCsvPath')) { $SummaryCsvPath.Trim() } else { '' }
if (-not [string]::IsNullOrWhiteSpace($resolvedSummaryCsv) -and (Test-Path -LiteralPath $resolvedSummaryCsv)) {
    Write-Status -Message "Reading summary from: $resolvedSummaryCsv"
    foreach ($row in (Import-Csv -LiteralPath $resolvedSummaryCsv | Where-Object { [string]$_.SummaryType -eq 'PerGpo' })) {
        $guid = ([string]$row.GpoGuid).Trim().ToLowerInvariant()
        if (-not [string]::IsNullOrWhiteSpace($guid)) {
            $summaryMap[$guid] = $row
        }
    }
    Write-Status -Message "Loaded $($summaryMap.Count) per-GPO summary rows."
}

# --- Build unified rows ---
$results = [System.Collections.Generic.List[object]]::new()

foreach ($link in $linkRows) {
    $guid       = ([string]$link.GpoGuid).Trim().ToLowerInvariant()
    $scopeDn    = ([string]$link.ScopeDn).Trim()
    $treeKey    = "${scopeDn}|${guid}"

    # GPO metadata (from 0010).
    $meta = if ($gpoMetaMap.ContainsKey($guid)) { $gpoMetaMap[$guid] } else { $null }

    # Scope tree (from 0030).
    $tree = if ($treeMap.ContainsKey($treeKey)) { $treeMap[$treeKey] } else { $null }

    # Per-GPO summary (from 0040).
    $summary = if ($summaryMap.ContainsKey($guid)) { $summaryMap[$guid] } else { $null }

    $results.Add([PSCustomObject][ordered]@{
        # Link identity
        ScopeDn               = $scopeDn
        ScopeType             = [string]$link.ScopeType
        ScopeName             = [string]$link.ScopeName
        GpoGuid               = $guid
        GpoName               = [string]$link.GpoName
        LinkOrder             = [string]$link.LinkOrder
        LinkEnabled           = [string]$link.LinkEnabled
        Enforced              = [string]$link.Enforced

        # GPO metadata (0010)
        GpoStatus             = if ($meta) { [string]$meta.GpoStatus } else { '' }
        ComputerConfigEnabled = if ($meta) { [string]$meta.ComputerConfigEnabled } else { '' }
        UserConfigEnabled     = if ($meta) { [string]$meta.UserConfigEnabled } else { '' }
        Description           = if ($meta) { [string]$meta.Description } else { '' }
        Owner                 = if ($meta) { [string]$meta.Owner } else { '' }
        WmiFilterName         = if ($meta) { [string]$meta.WmiFilterName } else { '' }
        WmiFilterQuery        = if ($meta) { [string]$meta.WmiFilterQuery } else { '' }
        CreationTime          = if ($meta) { [string]$meta.CreationTime } else { '' }
        ModificationTime      = if ($meta) { [string]$meta.ModificationTime } else { '' }
        GpoLinksCount         = if ($meta) { [string]$meta.LinksCount } else { '' }
        GpoEnforcedLinksCount = if ($meta) { [string]$meta.EnforcedLinksCount } else { '' }

        # Scope tree (0030)
        TreeKey               = if ($tree) { [string]$tree.TreeKey } else { '' }
        ScopeDepth            = if ($tree) { [string]$tree.ScopeDepth } else { '' }
        GpoPrecedence         = if ($tree) { [string]$tree.GpoPrecedence } else { '' }

        # Summary (0040)
        RiskTier              = if ($summary) { [string]$summary.RiskTier } else { '' }
        IsEmptyGpo            = if ($summary) { [string]$summary.IsEmptyGpo } else { '' }
    })
}

if ($results.Count -eq 0) {
    Write-Status -Message 'No link rows could be merged. Output will be empty.' -Level WARN
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Group Policy unified view builder completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
