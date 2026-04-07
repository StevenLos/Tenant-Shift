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
Microsoft.Graph.Authentication
Microsoft.Graph.Reports

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Exports SharePoint Online site activity metrics from Microsoft Graph.

.DESCRIPTION
    Exports site-level activity data from the Microsoft Graph SharePoint site usage
    detail report (/reports/getSharePointSiteUsageDetail). For each site collection,
    outputs the last activity date, active user count, file count, storage used (bytes),
    and site owner. Sites with no activity in the reporting period still appear in the
    report with empty activity fields.
    Only -DiscoverAll is supported — the Graph report endpoint returns all sites and
    cannot be filtered to a subset by site URL. The reporting period defaults to 30 days
    (D30) and can be changed via -ReportPeriod.
    All results are written to the output CSV.

.PARAMETER DiscoverAll
    Enumerate all sites from the Graph usage report. This is the only supported mode.

.PARAMETER ReportPeriod
    Reporting period for the usage detail report. Accepted values: D7, D30, D90, D180.
    Defaults to D30.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-SPOL-0040-Get-SharePointActiveSites.ps1 -DiscoverAll

    Export 30-day activity data for all SharePoint site collections.

.EXAMPLE
    .\D-SPOL-0040-Get-SharePointActiveSites.ps1 -DiscoverAll -ReportPeriod D90

    Export 90-day activity data for all SharePoint site collections.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Reports
    Required roles:   Reports Reader, Global Reader, or SharePoint Administrator
    Limitations:      Uses the Microsoft Graph v1.0 reports endpoint.
                      Data may be delayed up to 48 hours from real-time.
                      Personal (OneDrive) sites are included in the report.
                      The report period defines the lookback window; sites with no
                      activity in that window appear with empty activity fields.

    No input CSV — DiscoverAll is the only supported mode for this script.
#>
#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [switch]$DiscoverAll,

    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$ReportPeriod = 'D30',

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-SPOL-0040-Get-SharePointActiveSites_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'ReportPeriod',
    'SiteUrl',
    'SiteId',
    'SiteTitle',
    'OwnerDisplayName',
    'OwnerPrincipalName',
    'LastActivityDate',
    'ActiveUserCount',
    'FileCount',
    'ActiveFileCount',
    'StorageUsedBytes',
    'StorageAllocatedBytes',
    'IsDeleted',
    'RootWebTemplate'
)

Write-Status -Message 'Starting SharePoint active sites export script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Reports')
Ensure-GraphConnection -RequiredScopes @('Reports.Read.All')

Write-Status -Message "Fetching SharePoint site usage detail report (period: $ReportPeriod)." -Level WARN

# The Graph report returns CSV content directly. Fetch as raw string then parse.
$reportUri = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='$ReportPeriod')"

$reportContent = Invoke-WithRetry -OperationName "Get SharePoint site usage detail ($ReportPeriod)" -ScriptBlock {
    Invoke-MgGraphRequest -Method GET -Uri $reportUri -OutputType HttpResponseMessage -ErrorAction Stop |
        ForEach-Object { $_.Content.ReadAsStringAsync().Result }
}

# Parse CSV content from the report response.
$reportRows = $reportContent | ConvertFrom-Csv
Write-Status -Message "Fetched $($reportRows.Count) site usage records."

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($reportRow in $reportRows) {
    $siteUrl    = if ($reportRow.'Site URL') { $reportRow.'Site URL'.Trim() } else { '' }
    $primaryKey = if ($siteUrl) { $siteUrl } else { "Row$rowNumber" }

    try {
        $siteId             = if ($reportRow.'Site Id') { $reportRow.'Site Id'.Trim() } else { '' }
        $siteTitle          = if ($reportRow.'Site Title') { $reportRow.'Site Title'.Trim() } else { '' }
        $ownerDisplay       = if ($reportRow.'Owner Display Name') { $reportRow.'Owner Display Name'.Trim() } else { '' }
        $ownerPrincipal     = if ($reportRow.'Owner Principal Name') { $reportRow.'Owner Principal Name'.Trim() } else { '' }
        $lastActivity       = if ($reportRow.'Last Activity Date') { $reportRow.'Last Activity Date'.Trim() } else { '' }
        $activeUserCount    = if ($reportRow.'Active User Count') { $reportRow.'Active User Count'.Trim() } else { '' }
        $fileCount          = if ($reportRow.'File Count') { $reportRow.'File Count'.Trim() } else { '' }
        $activeFileCount    = if ($reportRow.'Active File Count') { $reportRow.'Active File Count'.Trim() } else { '' }
        $storageUsed        = if ($reportRow.'Storage Used (Byte)') { $reportRow.'Storage Used (Byte)'.Trim() } else { '' }
        $storageAllocated   = if ($reportRow.'Storage Allocated (Byte)') { $reportRow.'Storage Allocated (Byte)'.Trim() } else { '' }
        $isDeleted          = if ($reportRow.'Is Deleted') { $reportRow.'Is Deleted'.Trim() } else { '' }
        $rootWebTemplate    = if ($reportRow.'Root Web Template') { $reportRow.'Root Web Template'.Trim() } else { '' }

        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetSharePointActiveSites' -Status 'Completed' -Message 'Site usage data exported.' -Data ([ordered]@{
            SiteUrl              = $siteUrl
            SiteId               = $siteId
            SiteTitle            = $siteTitle
            OwnerDisplayName     = $ownerDisplay
            OwnerPrincipalName   = $ownerPrincipal
            LastActivityDate     = $lastActivity
            ActiveUserCount      = $activeUserCount
            FileCount            = $fileCount
            ActiveFileCount      = $activeFileCount
            StorageUsedBytes     = $storageUsed
            StorageAllocatedBytes = $storageAllocated
            IsDeleted            = $isDeleted
            RootWebTemplate      = $rootWebTemplate
        })))
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetSharePointActiveSites' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            SiteUrl = $siteUrl; SiteId = ''; SiteTitle = ''; OwnerDisplayName = ''; OwnerPrincipalName = ''
            LastActivityDate = ''; ActiveUserCount = ''; FileCount = ''; ActiveFileCount = ''
            StorageUsedBytes = ''; StorageAllocatedBytes = ''; IsDeleted = ''; RootWebTemplate = ''
        })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode'    -NotePropertyValue 'DiscoverAll'  -Force
    Add-Member -InputObject $result -NotePropertyName 'ReportPeriod' -NotePropertyValue $ReportPeriod  -Force
}

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint active sites export script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
