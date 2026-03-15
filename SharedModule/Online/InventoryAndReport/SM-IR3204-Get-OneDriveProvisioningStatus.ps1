<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-014648

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Online.SharePoint.PowerShell

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR3204-Get-OneDriveProvisioningStatus_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function ConvertTo-OneDriveUrlKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    return (($UserPrincipalName.Trim().ToLowerInvariant()) -replace '[^a-z0-9]', '_')
}

$requiredHeaders = @(
    'UserPrincipalName'
)

Write-Status -Message 'Starting OneDrive provisioning status inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Online.SharePoint.PowerShell')

if ([string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
    throw 'SharePointAdminUrl is required.'
}

$adminUrlTrimmed = $SharePointAdminUrl.Trim()
if ($adminUrlTrimmed -notmatch '^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$') {
    throw "SharePointAdminUrl '$adminUrlTrimmed' is invalid. Use format: https://<tenant>-admin.sharepoint.com"
}

Ensure-SharePointConnection -AdminUrl $adminUrlTrimmed

$adminUri = [uri]$adminUrlTrimmed
$oneDriveHost = ($adminUri.Host -replace '-admin\.', '-my.')

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

Write-Status -Message 'Loading personal sites to map OneDrive URLs by owner and URL key.'
$allSites = @(Invoke-WithRetry -OperationName 'Load personal sites' -ScriptBlock {
    Get-SPOSite -IncludePersonalSite $true -Limit All -Detailed -ErrorAction Stop
})

$personalSites = @($allSites | Where-Object {
    $url = ([string]$_.Url).Trim().ToLowerInvariant()
    $url.Contains('/personal/')
})

$sitesByOwner = @{}
$sitesByUrlKey = @{}

foreach ($site in $personalSites) {
    $siteUrl = ([string]$site.Url).Trim()
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        continue
    }

    $ownerKey = ([string]$site.Owner).Trim().ToLowerInvariant()
    if (-not [string]::IsNullOrWhiteSpace($ownerKey)) {
        if (-not $sitesByOwner.ContainsKey($ownerKey)) {
            $sitesByOwner[$ownerKey] = [System.Collections.Generic.List[object]]::new()
        }
        $sitesByOwner[$ownerKey].Add($site)
    }

    try {
        $uri = [uri]$siteUrl
        $parts = @($uri.AbsolutePath.Trim('/') -split '/')
        if ($parts.Count -ge 2 -and $parts[0].ToLowerInvariant() -eq 'personal') {
            $urlKey = $parts[1].ToLowerInvariant()
            if (-not [string]::IsNullOrWhiteSpace($urlKey)) {
                if (-not $sitesByUrlKey.ContainsKey($urlKey)) {
                    $sitesByUrlKey[$urlKey] = [System.Collections.Generic.List[object]]::new()
                }
                $sitesByUrlKey[$urlKey].Add($site)
            }
        }
    }
    catch {
        # Ignore malformed URLs in site results.
    }
}

$rowNumber = 1
foreach ($row in $rows) {
    $userPrincipalName = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
            throw 'UserPrincipalName is required. Use * to export all discovered OneDrive personal sites.'
        }

        if ($userPrincipalName -eq '*') {
            if ($personalSites.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey '*' -Action 'GetOneDriveProvisioningStatus' -Status 'NotFound' -Message 'No OneDrive personal sites were found.' -Data ([ordered]@{
                            UserPrincipalName           = '*'
                            OneDriveUrl                 = ''
                            ExpectedOneDriveUrl         = ''
                            SiteOwner                   = ''
                            SiteStatus                  = ''
                            IsProvisioned               = 'False'
                            StorageUsageCurrentMB       = ''
                            StorageQuotaMB              = ''
                        })))
                $rowNumber++
                continue
            }

            foreach ($site in @($personalSites | Sort-Object -Property Owner, Url)) {
                $siteUrl = ([string]$site.Url).Trim()
                $siteOwner = ([string]$site.Owner).Trim()
                $primaryKey = if (-not [string]::IsNullOrWhiteSpace($siteOwner)) { $siteOwner } else { $siteUrl }

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetOneDriveProvisioningStatus' -Status 'Completed' -Message 'OneDrive personal site exported.' -Data ([ordered]@{
                            UserPrincipalName           = $siteOwner
                            OneDriveUrl                 = $siteUrl
                            ExpectedOneDriveUrl         = ''
                            SiteOwner                   = $siteOwner
                            SiteStatus                  = ([string]$site.Status).Trim()
                            IsProvisioned               = 'True'
                            StorageUsageCurrentMB       = [string]$site.StorageUsageCurrent
                            StorageQuotaMB              = [string]$site.StorageQuota
                        })))
            }

            $rowNumber++
            continue
        }

        $ownerKey = $userPrincipalName.ToLowerInvariant()
        $urlKey = ConvertTo-OneDriveUrlKey -UserPrincipalName $userPrincipalName
        $expectedOneDriveUrl = "https://$oneDriveHost/personal/$urlKey"

        $matches = [System.Collections.Generic.List[object]]::new()

        if ($sitesByOwner.ContainsKey($ownerKey)) {
            foreach ($site in $sitesByOwner[$ownerKey]) {
                $matches.Add($site)
            }
        }

        if ($sitesByUrlKey.ContainsKey($urlKey)) {
            foreach ($site in $sitesByUrlKey[$urlKey]) {
                $matches.Add($site)
            }
        }

        $uniqueSitesByUrl = @{}
        foreach ($site in $matches) {
            $siteUrl = ([string]$site.Url).Trim().ToLowerInvariant()
            if (-not [string]::IsNullOrWhiteSpace($siteUrl) -and -not $uniqueSitesByUrl.ContainsKey($siteUrl)) {
                $uniqueSitesByUrl[$siteUrl] = $site
            }
        }

        $resolvedSites = @($uniqueSitesByUrl.Values)

        if ($resolvedSites.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'GetOneDriveProvisioningStatus' -Status 'NotFound' -Message 'No matching OneDrive personal site found for user.' -Data ([ordered]@{
                        UserPrincipalName           = $userPrincipalName
                        OneDriveUrl                 = ''
                        ExpectedOneDriveUrl         = $expectedOneDriveUrl
                        SiteOwner                   = ''
                        SiteStatus                  = ''
                        IsProvisioned               = 'False'
                        StorageUsageCurrentMB       = ''
                        StorageQuotaMB              = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($site in @($resolvedSites | Sort-Object -Property Url)) {
            $siteUrl = ([string]$site.Url).Trim()
            $siteOwner = ([string]$site.Owner).Trim()
            $message = if ($resolvedSites.Count -gt 1) {
                'Multiple OneDrive site matches found; row exported for each match.'
            }
            else {
                'OneDrive personal site exported.'
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'GetOneDriveProvisioningStatus' -Status 'Completed' -Message $message -Data ([ordered]@{
                        UserPrincipalName           = $userPrincipalName
                        OneDriveUrl                 = $siteUrl
                        ExpectedOneDriveUrl         = $expectedOneDriveUrl
                        SiteOwner                   = $siteOwner
                        SiteStatus                  = ([string]$site.Status).Trim()
                        IsProvisioned               = 'True'
                        StorageUsageCurrentMB       = [string]$site.StorageUsageCurrent
                        StorageQuotaMB              = [string]$site.StorageQuota
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($userPrincipalName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'GetOneDriveProvisioningStatus' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    UserPrincipalName           = $userPrincipalName
                    OneDriveUrl                 = ''
                    ExpectedOneDriveUrl         = ''
                    SiteOwner                   = ''
                    SiteStatus                  = ''
                    IsProvisioned               = ''
                    StorageUsageCurrentMB       = ''
                    StorageQuotaMB              = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'OneDrive provisioning status inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}









