<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260406-000000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
PnP.PowerShell

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Exports the list experience mode (Classic vs Modern) for all lists and libraries in SharePoint sites.

.DESCRIPTION
    For each site, retrieves all lists and document libraries and reports the configured
    list experience mode: Auto (tenant default), NewExperience (Modern), or ClassicExperience.
    One row per list. Hidden system lists are excluded by default.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all site collections in the tenant (-DiscoverAll parameter set).
    All results — including sites that could not be queried — are written to the output CSV.
    Useful as a pre-migration assessment before applying M-SPOL-0110.

.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include SiteUrl.
    See the companion .input.csv template for the full column list.

.PARAMETER DiscoverAll
    Enumerate all site collections in the tenant rather than processing from an input CSV file.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).
    Required for DiscoverAll mode to enumerate all site collections.
    Also used for the initial PnP tenant connection.

.PARAMETER IncludeHidden
    Include hidden system lists in the output. Defaults to false (hidden lists excluded).

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-SPOL-0050-Get-SharePointListExperience.ps1 -InputCsvPath .\D-SPOL-0050-Get-SharePointListExperience.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com

    Export list experience for the sites listed in the input CSV.

.EXAMPLE
    .\D-SPOL-0050-Get-SharePointListExperience.ps1 -DiscoverAll -SharePointAdminUrl https://los-admin.sharepoint.com

    Export list experience for all site collections in the tenant.

.NOTES
    Version:          1.0
    Required modules: PnP.PowerShell
    Required roles:   SharePoint Administrator
    Limitations:      Hidden lists are excluded by default (-IncludeHidden to override).
                      Reconnects to each site as SiteUrl changes — grouping input by site improves performance.

    CSV Fields:
    Column      Type      Required  Description
    ----------  --------  --------  -----------
    SiteUrl     String    Yes       Absolute URL of the site collection
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(Mandatory)]
    [ValidatePattern('^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$')]
    [string]$SharePointAdminUrl,

    [switch]$IncludeHidden,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-SPOL-0050-Get-SharePointListExperience_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function New-InventoryResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$RowNumber,
        [Parameter(Mandatory)][string]$PrimaryKey,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$Message,
        [Parameter(Mandatory)][hashtable]$Data
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

function Connect-PnPSite {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Url)
    Connect-PnPOnline -Url $Url -Interactive -ErrorAction Stop
}

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'SiteUrl',
    'SiteTitle',
    'ListId',
    'ListTitle',
    'ListTemplate',
    'ListExperience',
    'ItemCount',
    'IsHidden',
    'LastModified'
)

$requiredHeaders = @('SiteUrl')

Write-Status -Message 'Starting SharePoint list experience export script.'
Assert-ModuleCurrent -ModuleNames @('PnP.PowerShell')

$adminUrlTrimmed = $SharePointAdminUrl.TrimEnd('/')
$scopeMode       = 'Csv'

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. Fetching all site collections.' -Level WARN

    Connect-PnPSite -Url $adminUrlTrimmed

    $allSites = Invoke-WithRetry -OperationName 'Get all tenant sites' -ScriptBlock {
        Get-PnPTenantSite -ErrorAction Stop | Select-Object Url, Title
    }

    Write-Status -Message "Fetched $($allSites.Count) site collections."
    $rows = @($allSites | ForEach-Object { [PSCustomObject]@{ SiteUrl = $_.Url } })
} else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results        = [System.Collections.Generic.List[object]]::new()
$rowNumber      = 1
$currentSiteUrl = ''

foreach ($row in $rows) {
    $siteUrl    = Get-TrimmedValue -Value $row.SiteUrl
    $primaryKey = $siteUrl

    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        Write-Status -Message "Row $rowNumber skipped: SiteUrl is empty." -Level WARN
        $rowNumber++
        continue
    }

    try {
        if ($siteUrl -ne $currentSiteUrl) {
            Write-Status -Message "Connecting to site: $siteUrl"
            Connect-PnPSite -Url $siteUrl
            $currentSiteUrl = $siteUrl
        }

        $web = Invoke-WithRetry -OperationName "Get web properties for $siteUrl" -ScriptBlock {
            Get-PnPWeb -Includes Title -ErrorAction Stop
        }
        $siteTitle = if ($web.Title) { $web.Title.Trim() } else { '' }

        $lists = Invoke-WithRetry -OperationName "Get lists for $siteUrl" -ScriptBlock {
            Get-PnPList -Includes Id, Title, BaseTemplate, ListExperienceOptions, ItemCount, Hidden, LastItemModifiedDate -ErrorAction Stop
        }

        $listsToProcess = if ($IncludeHidden) { $lists } else { $lists | Where-Object { -not $_.Hidden } }

        if (-not $listsToProcess -or @($listsToProcess).Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetSharePointListExperience' -Status 'Completed' -Message 'No lists found on this site (or all lists are hidden).' -Data ([ordered]@{
                SiteUrl        = $siteUrl
                SiteTitle      = $siteTitle
                ListId         = ''
                ListTitle      = ''
                ListTemplate   = ''
                ListExperience = ''
                ItemCount      = ''
                IsHidden       = ''
                LastModified   = ''
            })))
        } else {
            foreach ($list in $listsToProcess) {
                $experienceLabel = switch ([string]$list.ListExperienceOptions) {
                    '0' { 'Auto' }
                    '1' { 'NewExperience' }
                    '2' { 'ClassicExperience' }
                    default { [string]$list.ListExperienceOptions }
                }

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetSharePointListExperience' -Status 'Completed' -Message 'List experience exported.' -Data ([ordered]@{
                    SiteUrl        = $siteUrl
                    SiteTitle      = $siteTitle
                    ListId         = [string]$list.Id
                    ListTitle      = $list.Title
                    ListTemplate   = [string]$list.BaseTemplate
                    ListExperience = $experienceLabel
                    ItemCount      = [string]$list.ItemCount
                    IsHidden       = [string]$list.Hidden
                    LastModified   = if ($list.LastItemModifiedDate) { $list.LastItemModifiedDate.ToString('o') } else { '' }
                })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetSharePointListExperience' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            SiteUrl = $siteUrl; SiteTitle = ''; ListId = ''; ListTitle = ''; ListTemplate = ''
            ListExperience = ''; ItemCount = ''; IsHidden = ''; LastModified = ''
        })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint list experience export script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
