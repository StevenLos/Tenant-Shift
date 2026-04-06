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
    Gets Group Policy Objects and exports results to CSV.

.DESCRIPTION
    Gets Group Policy Objects (GPOs) from the domain and writes the results to a CSV file.
    Exports GPO metadata including name, GUID, status, WMI filter details, link counts,
    enforced link counts, and creation/modification timestamps.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all GPOs in the domain (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
    Requires the GroupPolicy RSAT module (Windows PowerShell 5.1 only).

.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include GpoName or GpoGuid (or both).
    See the companion .input.csv template for the full column list.

.PARAMETER DiscoverAll
    Enumerate all GPOs in the domain rather than processing from an input CSV file.

.PARAMETER Domain
    Active Directory domain to query. If omitted, uses the current user's domain.

.PARAMETER Server
    Domain controller to target. If omitted, uses the default DC for the domain.

.PARAMETER MaxObjects
    Maximum number of GPOs to retrieve in DiscoverAll mode. 0 (default) means no limit.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-ADGP-0010-Get-GroupPolicyObjects.ps1 -InputCsvPath .\D-ADGP-0010-Get-GroupPolicyObjects.input.csv

    Inventory the GPOs listed in the input CSV file.

.EXAMPLE
    .\D-ADGP-0010-Get-GroupPolicyObjects.ps1 -DiscoverAll

    Discover and inventory all GPOs in the domain, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: GroupPolicy (RSAT Group Policy Management Tools, Windows PowerShell 5.1 only)
    Required roles:   Domain Administrator or Group Policy Creator Owners (read-only access sufficient for discovery)
    Limitations:      Requires RSAT Group Policy Management Tools installed on the machine running this script.
                      Link counts are resolved by querying the gpLink AD attribute via System.DirectoryServices.
                      Output of this script is consumed by D-ADGP-0020, D-ADGP-0030, D-ADGP-0040, D-ADGP-0050.

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
    [string]$Domain,

    [Parameter(ParameterSetName = 'FromCsv')]
    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$Server,

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [ValidateRange(0, 10000000)]
    [int]$MaxObjects = 0,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-ADGP-0010-Get-GroupPolicyObjects_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-GpoLinkCountsMap {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Server
    )

    # Query all AD objects that carry a gpLink attribute via System.DirectoryServices (no AD module required).
    # gpLink format per link entry: [LDAP://CN={GUID},CN=Policies,CN=System,DC=...;flags]
    # Flags: 0=enabled, 1=disabled, 2=enforced+enabled, 3=enforced+disabled
    $linkCountMap    = @{}
    $enforcedCountMap = @{}

    $ldapPath = if (-not [string]::IsNullOrWhiteSpace($Server)) {
        "LDAP://$Server/RootDSE"
    } else {
        'LDAP://RootDSE'
    }

    $rootDse    = [adsi]$ldapPath
    $defaultNC  = $rootDse.defaultNamingContext.ToString()
    $searchRoot = if (-not [string]::IsNullOrWhiteSpace($Server)) {
        [adsi]"LDAP://$Server/$defaultNC"
    } else {
        [adsi]"LDAP://$defaultNC"
    }

    $searcher = New-Object System.DirectoryServices.DirectorySearcher($searchRoot)
    $searcher.Filter      = '(gpLink=*)'
    $searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
    $searcher.PageSize    = 1000
    $searcher.PropertiesToLoad.Add('gpLink') | Out-Null

    $adResults = $searcher.FindAll()
    foreach ($adResult in $adResults) {
        $gpLinkProp = $adResult.Properties['gpLink']
        if ($null -eq $gpLinkProp -or $gpLinkProp.Count -eq 0) { continue }

        $gpLinkString = [string]$gpLinkProp[0]
        $linkEntries  = [regex]::Matches($gpLinkString, '\[([^\]]+)\]')

        foreach ($entry in $linkEntries) {
            $parts = $entry.Groups[1].Value -split ';'
            if ($parts.Count -lt 2) { continue }

            $ldapLink = $parts[0].Trim()
            $flagStr  = $parts[1].Trim()
            $flag     = 0
            [int]::TryParse($flagStr, [ref]$flag) | Out-Null

            $guidMatch = [regex]::Match($ldapLink, '\{([0-9a-fA-F\-]{36})\}')
            if (-not $guidMatch.Success) { continue }

            $gpoGuidKey = $guidMatch.Groups[1].Value.ToLowerInvariant()

            if (-not $linkCountMap.ContainsKey($gpoGuidKey)) {
                $linkCountMap[$gpoGuidKey]     = 0
                $enforcedCountMap[$gpoGuidKey] = 0
            }
            $linkCountMap[$gpoGuidKey]++

            # Bit 1 set (value 2) means enforced.
            if ($flag -band 2) {
                $enforcedCountMap[$gpoGuidKey]++
            }
        }
    }

    return @{
        LinksCount         = $linkCountMap
        EnforcedLinksCount = $enforcedCountMap
    }
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

function New-EmptyGpoData {
    [CmdletBinding()]
    param(
        [string]$GpoNameRequested = '',
        [string]$GpoGuidRequested = ''
    )

    return [ordered]@{
        GpoNameRequested      = $GpoNameRequested
        GpoGuidRequested      = $GpoGuidRequested
        GpoName               = ''
        GpoGuid               = ''
        DomainName            = ''
        GpoStatus             = ''
        ComputerConfigEnabled = ''
        UserConfigEnabled     = ''
        Description           = ''
        Owner                 = ''
        WmiFilterName         = ''
        WmiFilterQuery        = ''
        CreationTime          = ''
        ModificationTime      = ''
        LinksCount            = ''
        EnforcedLinksCount    = ''
    }
}

Write-Status -Message 'Starting Group Policy Object inventory script.'

# Fail-fast: ensure GroupPolicy RSAT is available before any GP cmdlet calls.
Assert-GroupPolicyRsatAvailable

$resolvedDomain  = Get-TrimmedValue -Value $Domain
$resolvedServer  = Get-TrimmedValue -Value $Server
$scopeMode       = 'Csv'
$runWasTruncated = $false

Write-Status -Message 'Resolving GPO link counts from Active Directory.'
$linkCountsMap = Invoke-WithRetry -OperationName 'Build GPO link count map' -ScriptBlock {
    Get-GpoLinkCountsMap -Server $resolvedServer
}

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled for Group Policy Objects.' -Level WARN
    $rows = @([PSCustomObject]@{ GpoName = ''; GpoGuid = '' })
} else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders @('GpoName', 'GpoGuid')
}

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $gpoNameReq = Get-TrimmedValue -Value $row.GpoName
    $gpoGuidReq = Get-TrimmedValue -Value $row.GpoGuid
    $primaryKey = if (-not [string]::IsNullOrWhiteSpace($gpoGuidReq)) {
        $gpoGuidReq
    } elseif (-not [string]::IsNullOrWhiteSpace($gpoNameReq)) {
        $gpoNameReq
    } else {
        "Row$rowNumber"
    }

    try {
        if ($scopeMode -ne 'DiscoverAll' -and [string]::IsNullOrWhiteSpace($gpoNameReq) -and [string]::IsNullOrWhiteSpace($gpoGuidReq)) {
            throw 'Each CSV row must supply GpoName or GpoGuid (or both).'
        }

        $gpoList = Invoke-WithRetry -OperationName "Get GPOs for $primaryKey" -ScriptBlock {
            if ($scopeMode -eq 'DiscoverAll') {
                $getParams = @{ All = $true; ErrorAction = 'Stop' }
                if (-not [string]::IsNullOrWhiteSpace($resolvedDomain)) { $getParams['Domain'] = $resolvedDomain }
                if (-not [string]::IsNullOrWhiteSpace($resolvedServer)) { $getParams['Server'] = $resolvedServer }
                return @(Get-GPO @getParams)
            } elseif (-not [string]::IsNullOrWhiteSpace($gpoGuidReq)) {
                $getParams = @{ Guid = [guid]$gpoGuidReq; ErrorAction = 'SilentlyContinue' }
                if (-not [string]::IsNullOrWhiteSpace($resolvedDomain)) { $getParams['Domain'] = $resolvedDomain }
                if (-not [string]::IsNullOrWhiteSpace($resolvedServer)) { $getParams['Server'] = $resolvedServer }
                $found = Get-GPO @getParams
                if ($found) { return @($found) }
                return @()
            } else {
                $getParams = @{ Name = $gpoNameReq; ErrorAction = 'SilentlyContinue' }
                if (-not [string]::IsNullOrWhiteSpace($resolvedDomain)) { $getParams['Domain'] = $resolvedDomain }
                if (-not [string]::IsNullOrWhiteSpace($resolvedServer)) { $getParams['Server'] = $resolvedServer }
                $found = Get-GPO @getParams
                if ($found) { return @($found) }
                return @()
            }
        }

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $gpoList.Count -gt $MaxObjects) {
            $gpoList         = @($gpoList | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($gpoList.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetGroupPolicyObject' -Status 'NotFound' -Message 'No matching GPO was found.' -Data (New-EmptyGpoData -GpoNameRequested $gpoNameReq -GpoGuidRequested $gpoGuidReq)))
            $rowNumber++
            continue
        }

        foreach ($gpo in @($gpoList | Sort-Object -Property DisplayName)) {
            $gpoPrimaryKey  = $gpo.Id.Guid.ToString().ToLowerInvariant()
            $wmiFilterName  = if ($gpo.WmiFilter) { Get-TrimmedValue -Value $gpo.WmiFilter.Name } else { '' }
            $wmiFilterQuery = if ($gpo.WmiFilter) { Get-TrimmedValue -Value $gpo.WmiFilter.Query } else { '' }
            $guidKey        = $gpo.Id.Guid.ToString().ToLowerInvariant()
            $linksCount         = if ($linkCountsMap.LinksCount.ContainsKey($guidKey)) { [string]$linkCountsMap.LinksCount[$guidKey] } else { '0' }
            $enforcedLinksCount = if ($linkCountsMap.EnforcedLinksCount.ContainsKey($guidKey)) { [string]$linkCountsMap.EnforcedLinksCount[$guidKey] } else { '0' }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $gpoPrimaryKey -Action 'GetGroupPolicyObject' -Status 'Completed' -Message 'GPO exported.' -Data ([ordered]@{
                GpoNameRequested      = $gpoNameReq
                GpoGuidRequested      = $gpoGuidReq
                GpoName               = Get-TrimmedValue -Value $gpo.DisplayName
                GpoGuid               = $gpoPrimaryKey
                DomainName            = Get-TrimmedValue -Value $gpo.DomainName
                GpoStatus             = [string]$gpo.GpoStatus
                ComputerConfigEnabled = [string]$gpo.Computer.Enabled
                UserConfigEnabled     = [string]$gpo.User.Enabled
                Description           = Get-TrimmedValue -Value $gpo.Description
                Owner                 = Get-TrimmedValue -Value $gpo.Owner
                WmiFilterName         = $wmiFilterName
                WmiFilterQuery        = $wmiFilterQuery
                CreationTime          = Get-TrimmedValue -Value $gpo.CreationTime
                ModificationTime      = Get-TrimmedValue -Value $gpo.ModificationTime
                LinksCount            = $linksCount
                EnforcedLinksCount    = $enforcedLinksCount
            })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetGroupPolicyObject' -Status 'Failed' -Message $_.Exception.Message -Data (New-EmptyGpoData -GpoNameRequested $gpoNameReq -GpoGuidRequested $gpoGuidReq)))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode'         -NotePropertyValue $scopeMode -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeDomain'       -NotePropertyValue $resolvedDomain -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeServer'       -NotePropertyValue $resolvedServer -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeMaxObjects'   -NotePropertyValue ($(if ($scopeMode -eq 'DiscoverAll') { [string]$MaxObjects } else { '' })) -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeWasTruncated' -NotePropertyValue ([string]$runWasTruncated) -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Group Policy Object inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
