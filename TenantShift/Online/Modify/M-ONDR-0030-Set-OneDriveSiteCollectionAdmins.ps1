<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-015515

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Online.SharePoint.PowerShell

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Modifies OneDriveSiteCollectionAdmins in Microsoft 365.

.DESCRIPTION
    Updates OneDriveSiteCollectionAdmins in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).

.PARAMETER AllowLastAdminRemoval
    When specified, allows removing the last site collection administrator. Use with caution.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M3206-Set-OneDriveSiteCollectionAdmins.ps1 -InputCsvPath .\3206.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M3206-Set-OneDriveSiteCollectionAdmins.ps1 -InputCsvPath .\3206.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Online.SharePoint.PowerShell
    Required roles:   SharePoint Administrator
    Limitations:      None known.

    CSV Fields:
    Column                      Type      Required  Description
    --------------------------  ----      --------  -----------
    UserPrincipalName           String    Yes       <fill in description>
    AddSiteCollectionAdmins     String    Yes       <fill in description>
    RemoveSiteCollectionAdmins  String    Yes       <fill in description>
    EnsureOneDriveOwnerIsAdmin  String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [switch]$AllowLastAdminRemoval,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3206-Set-OneDriveSiteCollectionAdmins_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function ConvertTo-OneDriveUrlKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    return (($UserPrincipalName.Trim().ToLowerInvariant()) -replace '[^a-z0-9]', '_')
}

function Get-NormalizedLogin {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$LoginName
    )

    $value = ([string]$LoginName).Trim()
    if ([string]::IsNullOrWhiteSpace($value)) {
        return ''
    }

    if ($value.Contains('|')) {
        $parts = $value.Split('|')
        $value = $parts[$parts.Length - 1]
    }

    return $value.Trim().ToLowerInvariant()
}

function Resolve-OneDriveSite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [hashtable]$SitesByOwner,

        [Parameter(Mandatory)]
        [hashtable]$SitesByUrlKey,

        [Parameter(Mandatory)]
        [string]$OneDriveHost
    )

    $ownerKey = $UserPrincipalName.Trim().ToLowerInvariant()
    $urlKey = ConvertTo-OneDriveUrlKey -UserPrincipalName $UserPrincipalName
    $expectedUrl = "https://$OneDriveHost/personal/$urlKey"

    $matches = [System.Collections.Generic.List[object]]::new()

    if ($SitesByOwner.ContainsKey($ownerKey)) {
        foreach ($site in $SitesByOwner[$ownerKey]) {
            $matches.Add($site)
        }
    }

    if ($SitesByUrlKey.ContainsKey($urlKey)) {
        foreach ($site in $SitesByUrlKey[$urlKey]) {
            $matches.Add($site)
        }
    }

    $uniqueSitesByUrl = @{}
    foreach ($site in $matches) {
        $siteUrlKey = ([string]$site.Url).Trim().ToLowerInvariant()
        if (-not [string]::IsNullOrWhiteSpace($siteUrlKey) -and -not $uniqueSitesByUrl.ContainsKey($siteUrlKey)) {
            $uniqueSitesByUrl[$siteUrlKey] = $site
        }
    }

    $resolvedSites = @($uniqueSitesByUrl.Values)
    if ($resolvedSites.Count -eq 0) {
        return [PSCustomObject]@{
            Status          = 'NotFound'
            Message         = 'No matching OneDrive personal site found for user.'
            Site            = $null
            ExpectedSiteUrl = $expectedUrl
        }
    }

    if ($resolvedSites.Count -eq 1) {
        return [PSCustomObject]@{
            Status          = 'Resolved'
            Message         = 'OneDrive personal site resolved.'
            Site            = $resolvedSites[0]
            ExpectedSiteUrl = $expectedUrl
        }
    }

    $expectedMatch = @($resolvedSites | Where-Object { ([string]$_.Url).Trim().Equals($expectedUrl, [System.StringComparison]::OrdinalIgnoreCase) })
    if ($expectedMatch.Count -eq 1) {
        return [PSCustomObject]@{
            Status          = 'Resolved'
            Message         = 'Multiple matches found; expected OneDrive URL selected.'
            Site            = $expectedMatch[0]
            ExpectedSiteUrl = $expectedUrl
        }
    }

    return [PSCustomObject]@{
        Status          = 'Ambiguous'
        Message         = 'Multiple OneDrive personal site matches found for user. Resolve ambiguity before applying updates.'
        Site            = $null
        ExpectedSiteUrl = $expectedUrl
    }
}

function Get-CurrentSiteAdminMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    $map = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $users = @(Invoke-WithRetry -OperationName "Load site users $SiteUrl" -ScriptBlock {
        Get-SPOUser -Site $SiteUrl -Limit All -ErrorAction Stop
    })

    foreach ($user in $users) {
        $isAdmin = $false
        if ($user.PSObject.Properties.Name -contains 'IsSiteAdmin') {
            $isAdmin = [bool]$user.IsSiteAdmin
        }
        elseif ($user.PSObject.Properties.Name -contains 'IsSiteCollectionAdmin') {
            $isAdmin = [bool]$user.IsSiteCollectionAdmin
        }

        if (-not $isAdmin) {
            continue
        }

        $loginName = ([string]$user.LoginName).Trim()
        $normalized = Get-NormalizedLogin -LoginName $loginName
        if ([string]::IsNullOrWhiteSpace($normalized)) {
            continue
        }

        if (-not $map.ContainsKey($normalized)) {
            $map[$normalized] = $loginName
        }
    }

    return $map
}

$requiredHeaders = @(
    'UserPrincipalName',
    'AddSiteCollectionAdmins',
    'RemoveSiteCollectionAdmins',
    'EnsureOneDriveOwnerIsAdmin'
)

Write-Status -Message 'Starting OneDrive site collection admin update script.'
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

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
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
            throw 'UserPrincipalName is required.'
        }

        if ($userPrincipalName -eq '*') {
            throw 'UserPrincipalName cannot be * for modify operations.'
        }

        $resolution = Resolve-OneDriveSite -UserPrincipalName $userPrincipalName -SitesByOwner $sitesByOwner -SitesByUrlKey $sitesByUrlKey -OneDriveHost $oneDriveHost
        if ($resolution.Status -eq 'NotFound') {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'SetOneDriveSiteAdmins' -Status 'NotFound' -Message "$($resolution.Message) Expected URL: $($resolution.ExpectedSiteUrl)"))
            $rowNumber++
            continue
        }

        if ($resolution.Status -eq 'Ambiguous') {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'SetOneDriveSiteAdmins' -Status 'Failed' -Message "$($resolution.Message) Expected URL: $($resolution.ExpectedSiteUrl)"))
            $rowNumber++
            continue
        }

        $site = $resolution.Site
        $siteUrl = ([string]$site.Url).Trim()
        $siteOwnerNormalized = Get-NormalizedLogin -LoginName ([string]$site.Owner)

        $ensureOwner = ConvertTo-Bool -Value $row.EnsureOneDriveOwnerIsAdmin -Default $true
        $addAdminsRaw = ConvertTo-Array -Value ([string]$row.AddSiteCollectionAdmins)
        $removeAdminsRaw = ConvertTo-Array -Value ([string]$row.RemoveSiteCollectionAdmins)

        $addAdmins = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $addAdminsRaw) {
            $normalized = Get-NormalizedLogin -LoginName $entry
            if (-not [string]::IsNullOrWhiteSpace($normalized)) {
                $null = $addAdmins.Add($normalized)
            }
        }

        $removeAdmins = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $removeAdminsRaw) {
            $normalized = Get-NormalizedLogin -LoginName $entry
            if (-not [string]::IsNullOrWhiteSpace($normalized)) {
                $null = $removeAdmins.Add($normalized)
            }
        }

        if ($ensureOwner -and -not [string]::IsNullOrWhiteSpace($siteOwnerNormalized)) {
            $null = $addAdmins.Add($siteOwnerNormalized)
            if ($removeAdmins.Contains($siteOwnerNormalized)) {
                $null = $removeAdmins.Remove($siteOwnerNormalized)
            }
        }

        $currentAdminMap = Get-CurrentSiteAdminMap -SiteUrl $siteUrl
        $finalAdmins = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($adminKey in $currentAdminMap.Keys) {
            $null = $finalAdmins.Add($adminKey)
        }

        foreach ($adminToAdd in $addAdmins) {
            $null = $finalAdmins.Add($adminToAdd)
        }

        foreach ($adminToRemove in $removeAdmins) {
            $null = $finalAdmins.Remove($adminToRemove)
        }

        if ((-not $AllowLastAdminRemoval) -and $finalAdmins.Count -eq 0) {
            throw 'Requested change would remove the last OneDrive site collection administrator. Use -AllowLastAdminRemoval to override.'
        }

        $adminsToAdd = [System.Collections.Generic.List[string]]::new()
        foreach ($candidate in $addAdmins) {
            if (-not $currentAdminMap.ContainsKey($candidate)) {
                $adminsToAdd.Add($candidate)
            }
        }

        $adminsToRemove = [System.Collections.Generic.List[string]]::new()
        foreach ($candidate in $removeAdmins) {
            if ($currentAdminMap.ContainsKey($candidate)) {
                $adminsToRemove.Add($candidate)
            }
        }

        if ($adminsToAdd.Count -eq 0 -and $adminsToRemove.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'SetOneDriveSiteAdmins' -Status 'Skipped' -Message 'OneDrive site admin membership already matches requested state.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($siteUrl, "Update OneDrive site collection administrators for $userPrincipalName")) {
            foreach ($adminToAdd in $adminsToAdd) {
                Invoke-WithRetry -OperationName "Grant OneDrive site admin $adminToAdd on $siteUrl" -ScriptBlock {
                    Set-SPOUser -Site $siteUrl -LoginName $adminToAdd -IsSiteCollectionAdmin $true -ErrorAction Stop
                }
            }

            foreach ($adminToRemove in $adminsToRemove) {
                $loginToRemove = $currentAdminMap[$adminToRemove]
                Invoke-WithRetry -OperationName "Remove OneDrive site admin $adminToRemove on $siteUrl" -ScriptBlock {
                    Set-SPOUser -Site $siteUrl -LoginName $loginToRemove -IsSiteCollectionAdmin $false -ErrorAction Stop
                }
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'SetOneDriveSiteAdmins' -Status 'Completed' -Message "OneDrive site admins updated on $siteUrl. Admins added: $($adminsToAdd.Count). Admins removed: $($adminsToRemove.Count)."))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'SetOneDriveSiteAdmins' -Status 'WhatIf' -Message 'OneDrive site admin update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($userPrincipalName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'SetOneDriveSiteAdmins' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'OneDrive site collection admin update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}




