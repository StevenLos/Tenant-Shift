<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-192500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Online.SharePoint.PowerShell
Microsoft.Graph.Authentication
Microsoft.Graph.Groups
Microsoft.Graph.Teams

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets SharePointSites and exports results to CSV.

.DESCRIPTION
    Gets SharePointSites from Microsoft 365 and writes the results to a CSV file.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all objects in scope (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all objects in scope rather than processing from an input CSV file. Uses the DiscoverAll parameter set.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-D3240-Get-SharePointSites.ps1 -InputCsvPath .\3240.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D3240-Get-SharePointSites.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Online.SharePoint.PowerShell, Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Teams
    Required roles:   SharePoint Administrator
    Limitations:      None known.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-SPOL-0010-Get-SharePointSites_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

    $base = New-ResultObject -RowNumber $RowNumber -PrimaryKey $PrimaryKey -Action $Action -Status $Status -Message $Message
    $ordered = [ordered]@{}

    foreach ($prop in $base.PSObject.Properties.Name) {
        $ordered[$prop] = $base.$prop
    }

    foreach ($key in $Data.Keys) {
        $ordered[$key] = $Data[$key]
    }

    return [PSCustomObject]$ordered
}

function Get-SitePropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Site,

        [Parameter(Mandatory)]
        [string[]]$PropertyNames
    )

    foreach ($name in $PropertyNames) {
        if ($Site.PSObject.Properties.Name -contains $name) {
            return [string]$Site.$name
        }
    }

    return ''
}

function Test-HasGroupConnection {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$GroupId
    )

    $value = ([string]$GroupId).Trim()
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $false
    }

    return ($value -ne '00000000-0000-0000-0000-000000000000')
}

function Ensure-TeamLookupReady {
    [CmdletBinding()]
    param()

    if ($script:graphDependenciesReady) {
        return
    }

    Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Teams')
    Ensure-GraphConnection -RequiredScopes @('Group.Read.All', 'Team.ReadBasic.All', 'Directory.Read.All')
    $script:graphDependenciesReady = $true
}

function Get-TeamConnectionFlag {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$GroupId
    )

    $normalizedGroupId = ([string]$GroupId).Trim()
    if (-not (Test-HasGroupConnection -GroupId $normalizedGroupId)) {
        return 'False'
    }

    if ($script:teamConnectionByGroupId.ContainsKey($normalizedGroupId)) {
        return $script:teamConnectionByGroupId[$normalizedGroupId]
    }

    try {
        Ensure-TeamLookupReady

        $team = Invoke-WithRetry -OperationName "Lookup Team for SharePoint site group $normalizedGroupId" -ScriptBlock {
            Get-MgGroupTeam -GroupId $normalizedGroupId -ErrorAction SilentlyContinue
        }

        $value = if ($team) { 'True' } else { 'False' }
        $script:teamConnectionByGroupId[$normalizedGroupId] = $value
        return $value
    }
    catch {
        Write-Status -Message "Unable to resolve Team attachment for group '$normalizedGroupId'. Error: $($_.Exception.Message)" -Level WARN
        $script:teamConnectionByGroupId[$normalizedGroupId] = ''
        return ''
    }
}

$requiredHeaders = @(
    'SiteUrl'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'SiteUrl',
    'Title',
    'Owner',
    'Template',
    'GroupId',
    'IsMicrosoft365GroupConnected',
    'IsMicrosoftTeamConnected',
    'HubSiteId',
    'IsHubSite',
    'SiteStatus',
    'SharingCapability',
    'StorageQuotaMB',
    'StorageUsageCurrentMB'
)

$script:graphDependenciesReady = $false
$script:teamConnectionByGroupId = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)

Write-Status -Message 'Starting SharePoint site inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Online.SharePoint.PowerShell')

if ([string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
    throw 'SharePointAdminUrl is required.'
}

$adminUrlTrimmed = $SharePointAdminUrl.Trim()
if ($adminUrlTrimmed -notmatch '^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$') {
    throw "SharePointAdminUrl '$adminUrlTrimmed' is invalid. Use: https://<tenant>-admin.sharepoint.com"
}

Ensure-SharePointConnection -AdminUrl $adminUrlTrimmed

$scopeMode = 'Csv'
if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. CSV input is bypassed.' -Level WARN

    $discoverRow = [ordered]@{}
    foreach ($header in $requiredHeaders) {
        $discoverRow[$header] = '*'
    }

    $rows = @([PSCustomObject]$discoverRow)
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $siteUrl = ([string]$row.SiteUrl).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl)) {
            throw 'SiteUrl is required. Use * to inventory all sites.'
        }

        $sites = @()
        if ($siteUrl -eq '*') {
            $sites = @(Invoke-WithRetry -OperationName 'Load all SharePoint sites' -ScriptBlock {
                Get-SPOSite -Limit All -Detailed -ErrorAction Stop
            })
        }
        else {
            $site = $null
            try {
                $site = Invoke-WithRetry -OperationName "Lookup SharePoint site $siteUrl" -ScriptBlock {
                    Get-SPOSite -Identity $siteUrl -Detailed -ErrorAction Stop
                }
            }
            catch {
                $message = ([string]$_.Exception.Message).ToLowerInvariant()
                if ($message -match 'cannot find|was not found|does not exist|not found') {
                    $site = $null
                }
                else {
                    throw
                }
            }

            if ($site) {
                $sites = @($site)
            }
        }

        if ($sites.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'GetSharePointSite' -Status 'NotFound' -Message 'No matching SharePoint sites were found.' -Data ([ordered]@{
                        SiteUrl                 = $siteUrl
                        Title                   = ''
                        Owner                   = ''
                        Template                = ''
                        StorageQuotaMB          = ''
                        StorageUsageCurrentMB   = ''
                        SiteStatus              = ''
                        SharingCapability       = ''
                        HubSiteId               = ''
                        IsHubSite               = ''
                        GroupId                 = ''
                        IsMicrosoft365GroupConnected = ''
                        IsMicrosoftTeamConnected     = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($site in @($sites | Sort-Object -Property Url)) {
            $resolvedSiteUrl = ([string]$site.Url).Trim()
            if ([string]::IsNullOrWhiteSpace($resolvedSiteUrl)) {
                $resolvedSiteUrl = ([string]$siteUrl).Trim()
            }

            $groupId = Get-SitePropertyValue -Site $site -PropertyNames @('GroupId', 'RelatedGroupId')
            $isGroupConnected = if (Test-HasGroupConnection -GroupId $groupId) { 'True' } else { 'False' }
            $isTeamConnected = Get-TeamConnectionFlag -GroupId $groupId

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $resolvedSiteUrl -Action 'GetSharePointSite' -Status 'Completed' -Message 'SharePoint site exported.' -Data ([ordered]@{
                        SiteUrl                 = $resolvedSiteUrl
                        Title                   = Get-SitePropertyValue -Site $site -PropertyNames @('Title')
                        Owner                   = Get-SitePropertyValue -Site $site -PropertyNames @('Owner')
                        Template                = Get-SitePropertyValue -Site $site -PropertyNames @('Template')
                        StorageQuotaMB          = Get-SitePropertyValue -Site $site -PropertyNames @('StorageQuota')
                        StorageUsageCurrentMB   = Get-SitePropertyValue -Site $site -PropertyNames @('StorageUsageCurrent')
                        SiteStatus              = Get-SitePropertyValue -Site $site -PropertyNames @('Status')
                        SharingCapability       = Get-SitePropertyValue -Site $site -PropertyNames @('SharingCapability')
                        HubSiteId               = Get-SitePropertyValue -Site $site -PropertyNames @('HubSiteId', 'HubSiteID')
                        IsHubSite               = Get-SitePropertyValue -Site $site -PropertyNames @('IsHubSite')
                        GroupId                 = $groupId
                        IsMicrosoft365GroupConnected = $isGroupConnected
                        IsMicrosoftTeamConnected     = $isTeamConnected
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($siteUrl) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'GetSharePointSite' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    SiteUrl                 = $siteUrl
                    Title                   = ''
                    Owner                   = ''
                    Template                = ''
                    StorageQuotaMB          = ''
                    StorageUsageCurrentMB   = ''
                    SiteStatus              = ''
                    SharingCapability       = ''
                    HubSiteId               = ''
                    IsHubSite               = ''
                    GroupId                 = ''
                    IsMicrosoft365GroupConnected = ''
                    IsMicrosoftTeamConnected     = ''
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
Write-Status -Message 'SharePoint site inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}









