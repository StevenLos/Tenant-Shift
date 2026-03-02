<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-004416

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Online.SharePoint.PowerShell

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3240-Get-SharePointSites_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$requiredHeaders = @(
    'SiteUrl'
)

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

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
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
                        Status                  = ''
                        SharingCapability       = ''
                        HubSiteId               = ''
                        IsHubSite               = ''
                        GroupId                 = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($site in @($sites | Sort-Object -Property Url)) {
            $resolvedSiteUrl = ([string]$site.Url).Trim()
            if ([string]::IsNullOrWhiteSpace($resolvedSiteUrl)) {
                $resolvedSiteUrl = ([string]$siteUrl).Trim()
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $resolvedSiteUrl -Action 'GetSharePointSite' -Status 'Completed' -Message 'SharePoint site exported.' -Data ([ordered]@{
                        SiteUrl                 = $resolvedSiteUrl
                        Title                   = Get-SitePropertyValue -Site $site -PropertyNames @('Title')
                        Owner                   = Get-SitePropertyValue -Site $site -PropertyNames @('Owner')
                        Template                = Get-SitePropertyValue -Site $site -PropertyNames @('Template')
                        StorageQuotaMB          = Get-SitePropertyValue -Site $site -PropertyNames @('StorageQuota')
                        StorageUsageCurrentMB   = Get-SitePropertyValue -Site $site -PropertyNames @('StorageUsageCurrent')
                        Status                  = Get-SitePropertyValue -Site $site -PropertyNames @('Status')
                        SharingCapability       = Get-SitePropertyValue -Site $site -PropertyNames @('SharingCapability')
                        HubSiteId               = Get-SitePropertyValue -Site $site -PropertyNames @('HubSiteId', 'HubSiteID')
                        IsHubSite               = Get-SitePropertyValue -Site $site -PropertyNames @('IsHubSite')
                        GroupId                 = Get-SitePropertyValue -Site $site -PropertyNames @('GroupId', 'RelatedGroupId')
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
                    Status                  = ''
                    SharingCapability       = ''
                    HubSiteId               = ''
                    IsHubSite               = ''
                    GroupId                 = ''
                })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint site inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







