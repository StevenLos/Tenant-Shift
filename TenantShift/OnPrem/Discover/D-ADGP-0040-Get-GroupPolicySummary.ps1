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
    Produces per-scope and per-GPO summary statistics from D-ADGP-0020 output.

.DESCRIPTION
    Post-processing script. Reads the output CSV from D-ADGP-0020-Get-GroupPolicyLinks.ps1
    and produces two summary views written to a single output CSV:
    - Per-scope summary: total linked GPOs, enforced count, disabled link count.
    - Per-GPO summary: total link count, enforced link count, empty-GPO flag (from D-ADGP-0010
      output when provided), and a risk tier.

    Risk tier logic:
    - High:   GPO is enforced on any scope, or linked directly to the domain root.
    - Medium: GPO is linked to one or more OUs (non-enforced, non-domain-root).
    - Low:    GPO is linked to Sites only, or has no active links (unlinked).

    No Active Directory calls are made — this script is a pure CSV transformer.

.PARAMETER LinksCsvPath
    Path to the output CSV from D-ADGP-0020-Get-GroupPolicyLinks.ps1.

.PARAMETER GpoObjectsCsvPath
    Optional path to the output CSV from D-ADGP-0010-Get-GroupPolicyObjects.ps1.
    When provided, the GpoStatus field is used to flag empty GPOs (AllSettingsDisabled).

.PARAMETER SummaryMode
    Controls which summary rows are written to the output.
    PerScope: one row per scope (OU/domain/site).
    PerGpo:   one row per GPO.
    Both:     all rows (default).

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-ADGP-0040-Get-GroupPolicySummary.ps1 -LinksCsvPath .\Results_D-ADGP-0020*.csv

    Generate summary statistics from D-ADGP-0020 output.

.EXAMPLE
    .\D-ADGP-0040-Get-GroupPolicySummary.ps1 -LinksCsvPath .\Results_D-ADGP-0020*.csv -GpoObjectsCsvPath .\Results_D-ADGP-0010*.csv -SummaryMode PerGpo

    Generate per-GPO summary including empty-GPO flags from D-ADGP-0010.

.NOTES
    Version:          1.0
    Required modules: None (pure CSV post-processor)
    Required roles:   None (reads local CSV files only)
    Limitations:      Requires completed D-ADGP-0020 output as input.
                      Empty-GPO detection requires D-ADGP-0010 output via -GpoObjectsCsvPath.
                      Risk tier is a heuristic — review High-tier GPOs manually.

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

    [string]$GpoObjectsCsvPath,

    [ValidateSet('PerScope', 'PerGpo', 'Both')]
    [string]$SummaryMode = 'Both',

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-ADGP-0040-Get-GroupPolicySummary_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Resolve-RiskTier {
    [CmdletBinding()]
    param(
        [string]$GpoGuid,
        [System.Collections.Generic.List[object]]$GpoLinks
    )

    $gpoLinkRows = @($GpoLinks | Where-Object { [string]$_.GpoGuid -eq $GpoGuid })

    if ($gpoLinkRows.Count -eq 0) { return 'Low' }

    # High: enforced on any scope.
    if ($gpoLinkRows | Where-Object { [string]$_.Enforced -eq 'True' }) {
        return 'High'
    }

    # High: linked to domain root (ScopeType = Domain).
    if ($gpoLinkRows | Where-Object { [string]$_.ScopeType -eq 'Domain' }) {
        return 'High'
    }

    # Medium: linked to any OU.
    if ($gpoLinkRows | Where-Object { [string]$_.ScopeType -eq 'OU' }) {
        return 'Medium'
    }

    # Low: site-linked only, or unlinked.
    return 'Low'
}

Write-Status -Message 'Starting Group Policy summary script.'

if (-not (Test-Path -LiteralPath $LinksCsvPath)) {
    throw "Links CSV not found: $LinksCsvPath"
}

Write-Status -Message "Reading GPO links from: $LinksCsvPath"
$linkRows = Import-Csv -LiteralPath $LinksCsvPath
$completedLinks = [System.Collections.Generic.List[object]]::new()
foreach ($row in ($linkRows | Where-Object { [string]$_.Status -eq 'Completed' })) {
    $completedLinks.Add($row)
}
Write-Status -Message "Loaded $($completedLinks.Count) completed link rows."

# Load GPO objects if path provided.
$gpoStatusMap = @{}
$resolvedGpoCsv = if ($PSBoundParameters.ContainsKey('GpoObjectsCsvPath')) { $GpoObjectsCsvPath.Trim() } else { '' }
if (-not [string]::IsNullOrWhiteSpace($resolvedGpoCsv) -and (Test-Path -LiteralPath $resolvedGpoCsv)) {
    Write-Status -Message "Reading GPO objects from: $resolvedGpoCsv"
    $gpoRows = Import-Csv -LiteralPath $resolvedGpoCsv
    foreach ($gpoRow in ($gpoRows | Where-Object { [string]$_.Status -eq 'Completed' })) {
        $guid = ([string]$gpoRow.GpoGuid).Trim().ToLowerInvariant()
        if (-not [string]::IsNullOrWhiteSpace($guid)) {
            $gpoStatusMap[$guid] = [string]$gpoRow.GpoStatus
        }
    }
    Write-Status -Message "Loaded $($gpoStatusMap.Count) GPO status entries."
}

$results = [System.Collections.Generic.List[object]]::new()

# --- Per-Scope summary ---
if ($SummaryMode -eq 'PerScope' -or $SummaryMode -eq 'Both') {
    $scopeGroups = $completedLinks | Group-Object -Property ScopeDn
    foreach ($scopeGroup in ($scopeGroups | Sort-Object -Property Name)) {
        $scopeLinks       = @($scopeGroup.Group)
        $totalLinked      = $scopeLinks.Count
        $enforcedCount    = ($scopeLinks | Where-Object { [string]$_.Enforced -eq 'True' }).Count
        $disabledCount    = ($scopeLinks | Where-Object { [string]$_.LinkEnabled -eq 'False' }).Count
        $scopeType        = [string]$scopeLinks[0].ScopeType
        $scopeName        = [string]$scopeLinks[0].ScopeName

        $results.Add([PSCustomObject][ordered]@{
            SummaryType        = 'PerScope'
            ScopeDn            = [string]$scopeGroup.Name
            ScopeType          = $scopeType
            ScopeName          = $scopeName
            GpoGuid            = ''
            GpoName            = ''
            TotalLinkedGpos    = [string]$totalLinked
            EnforcedLinkCount  = [string]$enforcedCount
            DisabledLinkCount  = [string]$disabledCount
            IsEmptyGpo         = ''
            RiskTier           = ''
        })
    }
}

# --- Per-GPO summary ---
if ($SummaryMode -eq 'PerGpo' -or $SummaryMode -eq 'Both') {
    $gpoGroups = $completedLinks | Group-Object -Property GpoGuid
    foreach ($gpoGroup in ($gpoGroups | Sort-Object -Property { [string]$_.Group[0].GpoName })) {
        $guid             = [string]$gpoGroup.Name
        $gpoLinks         = @($gpoGroup.Group)
        $gpoName          = [string]$gpoLinks[0].GpoName
        $totalLinks       = $gpoLinks.Count
        $enforcedLinks    = ($gpoLinks | Where-Object { [string]$_.Enforced -eq 'True' }).Count
        $gpoStatus        = if ($gpoStatusMap.ContainsKey($guid)) { $gpoStatusMap[$guid] } else { '' }
        $isEmptyGpo       = if ($gpoStatus -eq 'AllSettingsDisabled') { 'True' } elseif ([string]::IsNullOrWhiteSpace($gpoStatus)) { '' } else { 'False' }
        $riskTier         = Resolve-RiskTier -GpoGuid $guid -GpoLinks $completedLinks

        $results.Add([PSCustomObject][ordered]@{
            SummaryType        = 'PerGpo'
            ScopeDn            = ''
            ScopeType          = ''
            ScopeName          = ''
            GpoGuid            = $guid
            GpoName            = $gpoName
            TotalLinkedGpos    = [string]$totalLinks
            EnforcedLinkCount  = [string]$enforcedLinks
            DisabledLinkCount  = ''
            IsEmptyGpo         = $isEmptyGpo
            RiskTier           = $riskTier
        })
    }
}

if ($results.Count -eq 0) {
    Write-Status -Message 'No completed link rows found in the input CSV. Output will be empty.' -Level WARN
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Group Policy summary script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
