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
GroupPolicy

.MODULEVERSIONPOLICY
Best-effort PSGallery check (offline-safe)

.SYNOPSIS
    Gets Group Policy Object links and exports results to CSV.

.DESCRIPTION
    Gets all GPO links in the domain and writes the results to a CSV file.
    For each link, exports the scope type (Domain / OU / Site), the scope distinguished name,
    link order, enforcement state, link enabled state, and resolves the GPO GUID to a display
    name using the output of D-ADGP-0010 when available.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath
    to scope discovery to specific OUs or domain DNs) or by enumerating all linked scopes
    in the domain (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
    Requires the GroupPolicy RSAT module (Windows PowerShell 5.1 only).

.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include a ScopeDn (OU or domain distinguished name).
    See the companion .input.csv template for the full column list.

.PARAMETER DiscoverAll
    Enumerate all GPO-linked scopes in the domain rather than processing from an input CSV file.

.PARAMETER GpoObjectsCsvPath
    Optional path to the output CSV from D-ADGP-0010-Get-GroupPolicyObjects.ps1.
    When provided, GPO display names are resolved from this file instead of querying AD again.

.PARAMETER Domain
    Active Directory domain to query. If omitted, uses the current user's domain.

.PARAMETER Server
    Domain controller to target. If omitted, uses the default DC for the domain.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-ADGP-0020-Get-GroupPolicyLinks.ps1 -InputCsvPath .\D-ADGP-0020-Get-GroupPolicyLinks.input.csv

    Export GPO links for the scopes listed in the input CSV file.

.EXAMPLE
    .\D-ADGP-0020-Get-GroupPolicyLinks.ps1 -DiscoverAll -GpoObjectsCsvPath .\Results_D-ADGP-0010*.csv

    Discover all GPO links in the domain and resolve GPO names from the D-ADGP-0010 output.

.NOTES
    Version:          1.0
    Required modules: GroupPolicy (RSAT Group Policy Management Tools, Windows PowerShell 5.1 only)
    Required roles:   Domain Administrator or delegated read access (read-only sufficient for discovery)
    Limitations:      Requires RSAT Group Policy Management Tools installed on the machine running this script.
                      Run D-ADGP-0010 first to enable GPO name resolution via -GpoObjectsCsvPath.
                      Output of this script is consumed by D-ADGP-0030, D-ADGP-0040, and D-ADGP-0050.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(ParameterSetName = 'FromCsv')]
    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$GpoObjectsCsvPath,

    [Parameter(ParameterSetName = 'FromCsv')]
    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$Domain,

    [Parameter(ParameterSetName = 'FromCsv')]
    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$Server,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-ADGP-0020-Get-GroupPolicyLinks_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Assert-GroupPolicyRsatAvailable {
    # Fail-fast guard: verify the GroupPolicy RSAT module is present before any GP cmdlet calls.
    if (-not (Get-Command -Name 'Get-GPO' -ErrorAction SilentlyContinue)) {
        throw 'The GroupPolicy RSAT module is not available on this machine. Install RSAT Group Policy Management Tools and ensure the GroupPolicy module is importable under Windows PowerShell 5.1 before running this script.'
    }
    Import-Module -Name GroupPolicy -ErrorAction Stop
}

function Get-GpoNameMap {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$GpoObjectsCsvPath,

        [AllowEmptyString()]
        [string]$Domain,

        [AllowEmptyString()]
        [string]$Server
    )

    # Build a GUID->DisplayName map, preferring the D-ADGP-0010 CSV when available.
    $map = @{}

    if (-not [string]::IsNullOrWhiteSpace($GpoObjectsCsvPath) -and (Test-Path -LiteralPath $GpoObjectsCsvPath)) {
        Write-Status -Message "Loading GPO name map from D-ADGP-0010 output: $GpoObjectsCsvPath"
        $csvRows = Import-Csv -LiteralPath $GpoObjectsCsvPath
        foreach ($csvRow in $csvRows) {
            if ($csvRow.PSObject.Properties['GpoGuid'] -and $csvRow.PSObject.Properties['GpoName']) {
                $guid = ([string]$csvRow.GpoGuid).Trim().ToLowerInvariant()
                $name = ([string]$csvRow.GpoName).Trim()
                if (-not [string]::IsNullOrWhiteSpace($guid) -and -not $map.ContainsKey($guid)) {
                    $map[$guid] = $name
                }
            }
        }
        Write-Status -Message "Loaded $($map.Count) GPO name entries from D-ADGP-0010 output."
        return $map
    }

    Write-Status -Message 'No D-ADGP-0010 CSV provided — loading GPO names from Active Directory.'
    $getParams = @{ All = $true; ErrorAction = 'Stop' }
    if (-not [string]::IsNullOrWhiteSpace($Domain)) { $getParams['Domain'] = $Domain }
    if (-not [string]::IsNullOrWhiteSpace($Server)) { $getParams['Server'] = $Server }

    $allGpos = Get-GPO @getParams
    foreach ($gpo in $allGpos) {
        $guid = $gpo.Id.Guid.ToString().ToLowerInvariant()
        $map[$guid] = Get-TrimmedValue -Value $gpo.DisplayName
    }
    Write-Status -Message "Loaded $($map.Count) GPO names from Active Directory."
    return $map
}

function Get-LinkedScopeObjects {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$SearchDn,

        [AllowEmptyString()]
        [string]$Server
    )

    # Use System.DirectoryServices to find all objects with gpLink attribute under the given DN.
    # When SearchDn is empty, searches from the domain root.
    $ldapPath = if (-not [string]::IsNullOrWhiteSpace($Server)) {
        'LDAP://' + $Server + '/RootDSE'
    } else {
        'LDAP://RootDSE'
    }

    $rootDse   = [adsi]$ldapPath
    $defaultNC = $rootDse.defaultNamingContext.ToString()

    $baseDn = if (-not [string]::IsNullOrWhiteSpace($SearchDn)) { $SearchDn } else { $defaultNC }

    $searchRoot = if (-not [string]::IsNullOrWhiteSpace($Server)) {
        [adsi]("LDAP://$Server/" + $baseDn)
    } else {
        [adsi]("LDAP://" + $baseDn)
    }

    $searcher = New-Object System.DirectoryServices.DirectorySearcher($searchRoot)
    $searcher.Filter      = '(gpLink=*)'
    $searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
    $searcher.PageSize    = 1000
    $searcher.PropertiesToLoad.Add('distinguishedName') | Out-Null
    $searcher.PropertiesToLoad.Add('objectClass') | Out-Null
    $searcher.PropertiesToLoad.Add('gpLink') | Out-Null
    $searcher.PropertiesToLoad.Add('gpOptions') | Out-Null
    $searcher.PropertiesToLoad.Add('name') | Out-Null

    return $searcher.FindAll()
}

function Resolve-ScopeType {
    [CmdletBinding()]
    param([string[]]$ObjectClasses)

    if ($ObjectClasses -contains 'domainDNS') { return 'Domain' }
    if ($ObjectClasses -contains 'organizationalUnit') { return 'OU' }
    if ($ObjectClasses -contains 'site') { return 'Site' }
    return 'Other'
}

function ConvertFrom-GpLinkString {
    [CmdletBinding()]
    param([string]$GpLinkString)

    # Parse the gpLink attribute into structured link entries.
    # Returns list of [ordered]@{ GpoGuid; LinkOrder; LinkEnabled; Enforced }
    $entries = [System.Collections.Generic.List[object]]::new()
    $linkMatches = [regex]::Matches($GpLinkString, '\[([^\]]+)\]')
    $order = 1

    # gpLink entries are listed highest-priority first (link order 1 = highest priority).
    foreach ($match in $linkMatches) {
        $parts = $match.Groups[1].Value -split ';'
        if ($parts.Count -lt 2) { continue }

        $ldapLink = $parts[0].Trim()
        $flagStr  = $parts[1].Trim()
        $flag     = 0
        [int]::TryParse($flagStr, [ref]$flag) | Out-Null

        $guidMatch = [regex]::Match($ldapLink, '\{([0-9a-fA-F\-]{36})\}')
        if (-not $guidMatch.Success) { continue }

        $gpoGuid    = $guidMatch.Groups[1].Value.ToLowerInvariant()
        # Flag bit 0 set means link is disabled; bit 1 set means enforced.
        $linkEnabled = -not [bool]($flag -band 1)
        $enforced    = [bool]($flag -band 2)

        $entries.Add([ordered]@{
            GpoGuid     = $gpoGuid
            LinkOrder   = $order
            LinkEnabled = $linkEnabled
            Enforced    = $enforced
        })
        $order++
    }

    return $entries
}

function New-InventoryResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$RowNumber,

        [Parameter(Mandatory)]
        [string]$PrimaryKey,

        [Parameter(Mandatory)]
        [string]$Action,

        [Parameter(Mandatory)]
        [string]$Status,

        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory)]
        [hashtable]$Data
    )

    $base    = New-ResultObject -RowNumber $RowNumber -PrimaryKey $PrimaryKey -Action $Action -Status $Status -Message $Message
    $ordered = [ordered]@{}

    foreach ($prop in $base.PSObject.Properties.Name) {
        $ordered[$prop] = $base.$prop
    }

    foreach ($key in $Data.Keys) {
        $ordered[$key] = $Data[$key]
    }

    return [PSCustomObject]$ordered
}

function New-EmptyLinkData {
    [CmdletBinding()]
    param(
        [string]$ScopeDnRequested = ''
    )

    return [ordered]@{
        ScopeDnRequested = $ScopeDnRequested
        ScopeDn          = ''
        ScopeType        = ''
        ScopeName        = ''
        GpoGuid          = ''
        GpoName          = ''
        LinkOrder        = ''
        LinkEnabled      = ''
        Enforced         = ''
    }
}

Write-Status -Message 'Starting Group Policy link inventory script.'

# Fail-fast: ensure GroupPolicy RSAT is available before any GP cmdlet calls.
Assert-GroupPolicyRsatAvailable

$resolvedDomain  = Get-TrimmedValue -Value $Domain
$resolvedServer  = Get-TrimmedValue -Value $Server
$resolvedGpoCsv  = Get-TrimmedValue -Value $GpoObjectsCsvPath
$scopeMode       = 'Csv'

# Build GPO GUID -> display name map.
$gpoNameMap = Invoke-WithRetry -OperationName 'Build GPO name map' -ScriptBlock {
    Get-GpoNameMap -GpoObjectsCsvPath $resolvedGpoCsv -Domain $resolvedDomain -Server $resolvedServer
}

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled for Group Policy links.' -Level WARN
    $rows = @([PSCustomObject]@{ ScopeDn = '' })
} else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders @('ScopeDn')
}

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $scopeDnReq = Get-TrimmedValue -Value $row.ScopeDn
    $primaryKey = if (-not [string]::IsNullOrWhiteSpace($scopeDnReq)) { $scopeDnReq } else { 'DomainRoot' }

    try {
        $linkedObjects = Invoke-WithRetry -OperationName "Get linked scopes for $primaryKey" -ScriptBlock {
            Get-LinkedScopeObjects -SearchDn $scopeDnReq -Server $resolvedServer
        }

        if ($null -eq $linkedObjects -or $linkedObjects.Count -eq 0) {
            $emptyData = New-EmptyLinkData -ScopeDnRequested $scopeDnReq
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetGroupPolicyLink' -Status 'NotFound' -Message 'No GPO-linked scopes were found under the specified DN.' -Data $emptyData))
            $rowNumber++
            continue
        }

        foreach ($linkedObj in $linkedObjects) {
            $scopeDn       = [string]$linkedObj.Properties['distinguishedName'][0]
            $scopeName     = if ($linkedObj.Properties['name'].Count -gt 0) { [string]$linkedObj.Properties['name'][0] } else { '' }
            $gpLinkString  = if ($linkedObj.Properties['gpLink'].Count -gt 0) { [string]$linkedObj.Properties['gpLink'][0] } else { '' }
            $objClasses    = @($linkedObj.Properties['objectClass'] | ForEach-Object { [string]$_ })
            $scopeType     = Resolve-ScopeType -ObjectClasses $objClasses

            if ([string]::IsNullOrWhiteSpace($gpLinkString)) { continue }

            $linkEntries = ConvertFrom-GpLinkString -GpLinkString $gpLinkString

            foreach ($linkEntry in $linkEntries) {
                $gpoGuid    = $linkEntry.GpoGuid
                $gpoName    = if ($gpoNameMap.ContainsKey($gpoGuid)) { $gpoNameMap[$gpoGuid] } else { '' }
                $linkPk     = "${scopeDn}|${gpoGuid}"

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $linkPk -Action 'GetGroupPolicyLink' -Status 'Completed' -Message 'GPO link exported.' -Data ([ordered]@{
                    ScopeDnRequested = $scopeDnReq
                    ScopeDn          = Get-TrimmedValue -Value $scopeDn
                    ScopeType        = $scopeType
                    ScopeName        = Get-TrimmedValue -Value $scopeName
                    GpoGuid          = $gpoGuid
                    GpoName          = $gpoName
                    LinkOrder        = [string]$linkEntry.LinkOrder
                    LinkEnabled      = [string]$linkEntry.LinkEnabled
                    Enforced         = [string]$linkEntry.Enforced
                })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetGroupPolicyLink' -Status 'Failed' -Message $_.Exception.Message -Data (New-EmptyLinkData -ScopeDnRequested $scopeDnReq)))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode'   -NotePropertyValue $scopeMode -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeDomain' -NotePropertyValue $resolvedDomain -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeServer' -NotePropertyValue $resolvedServer -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Group Policy link inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
