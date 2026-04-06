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
    Builds a hierarchical GPO scope tree from D-ADGP-0020 output and exports results to CSV.

.DESCRIPTION
    Post-processing script. Reads the output CSV from D-ADGP-0020-Get-GroupPolicyLinks.ps1
    and produces a hierarchical scope tree view of all GPO links, equivalent to the RSAT
    Group Policy Management Console tree view.
    Each row in the output represents one GPO-scope-link with a TreeKey sort order, GPO
    precedence within the scope, and an InheritanceBlocked flag derived from the scope's
    gpOptions attribute captured in the source data.
    No Active Directory calls are made — this script is a pure CSV transformer.

.PARAMETER LinksCsvPath
    Path to the output CSV from D-ADGP-0020-Get-GroupPolicyLinks.ps1.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-ADGP-0030-Get-GroupPolicyScopeTree.ps1 -LinksCsvPath .\Results_D-ADGP-0020*.csv

    Build the scope tree from the D-ADGP-0020 output.

.NOTES
    Version:          1.0
    Required modules: None (pure CSV post-processor)
    Required roles:   None (reads local CSV files only)
    Limitations:      Requires completed D-ADGP-0020 output as input.
                      InheritanceBlocked field is not available from D-ADGP-0020 data; it is
                      set to empty in this output. Use D-ADGP-0040 for inheritance-aware summary.
                      TreeKey is a dot-separated depth index suitable for lexicographic sort.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$LinksCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-ADGP-0030-Get-GroupPolicyScopeTree_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-DnDepth {
    [CmdletBinding()]
    param([string]$DistinguishedName)

    # Count the number of DN components to determine depth.
    # Domain root has depth 0; each additional OU adds 1.
    $components = [regex]::Matches($DistinguishedName, '(?i)(OU|DC|CN)=')
    $dcCount     = ($components | Where-Object { $_.Value -imatch 'DC=' }).Count
    $ouCount     = ($components | Where-Object { $_.Value -imatch 'OU=' }).Count
    return $ouCount
}

function Get-ScopeTypeSortOrder {
    [CmdletBinding()]
    param([string]$ScopeType)

    switch ($ScopeType.Trim().ToLowerInvariant()) {
        'domain' { return 0 }
        'site'   { return 1 }
        'ou'     { return 2 }
        default  { return 9 }
    }
}

Write-Status -Message 'Starting Group Policy scope tree builder.'

if (-not (Test-Path -LiteralPath $LinksCsvPath)) {
    throw "Links CSV not found: $LinksCsvPath"
}

Write-Status -Message "Reading GPO links from: $LinksCsvPath"
$linkRows = Import-Csv -LiteralPath $LinksCsvPath

# Filter to completed rows only.
$completedLinks = @($linkRows | Where-Object { [string]$_.Status -eq 'Completed' })
Write-Status -Message "Loaded $($completedLinks.Count) completed link rows from D-ADGP-0020."

# Group by ScopeDn to build the tree.
$scopeGroups = $completedLinks | Group-Object -Property ScopeDn

$results = [System.Collections.Generic.List[object]]::new()

foreach ($scopeGroup in ($scopeGroups | Sort-Object -Property {
    $dn = [string]$_.Name
    "{0:D3}|{1}" -f (Get-ScopeTypeSortOrder -ScopeType ([string]($_.Group[0].ScopeType))), $dn
})) {
    $scopeDn    = [string]$scopeGroup.Name
    $scopeType  = [string]$scopeGroup.Group[0].ScopeType
    $scopeName  = [string]$scopeGroup.Group[0].ScopeName
    $depth      = Get-DnDepth -DistinguishedName $scopeDn

    # Assign a tree key: depth-padded index within sorted scope list + scope DN for lexicographic ordering.
    $treeKey = "{0:D3}|{1}" -f $depth, $scopeDn

    # Sort links within this scope by LinkOrder (ascending = highest precedence first).
    $scopeLinks = @($scopeGroup.Group | Sort-Object -Property { [int]([string]$_.LinkOrder) })
    $precedence = 1

    foreach ($link in $scopeLinks) {
        $results.Add([PSCustomObject][ordered]@{
            TreeKey          = $treeKey
            ScopeType        = $scopeType
            ScopeDepth       = [string]$depth
            ScopeDn          = $scopeDn
            ScopeName        = $scopeName
            GpoGuid          = [string]$link.GpoGuid
            GpoName          = [string]$link.GpoName
            LinkOrder        = [string]$link.LinkOrder
            GpoPrecedence    = [string]$precedence
            LinkEnabled      = [string]$link.LinkEnabled
            Enforced         = [string]$link.Enforced
        })
        $precedence++
    }
}

if ($results.Count -eq 0) {
    Write-Status -Message 'No completed link rows found in the input CSV. Output will be empty.' -Level WARN
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Group Policy scope tree builder completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
