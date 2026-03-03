<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-210000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Online.SharePoint.PowerShell

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M3208-Revoke-OneDriveExternalSharingLinks_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-NormalizedEmail {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Email
    )

    $value = ([string]$Email).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($value)) {
        return ''
    }

    return $value
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

function Get-ExternalSitePrincipals {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    $principals = [System.Collections.Generic.List[object]]::new()

    $siteUsers = @(Invoke-WithRetry -OperationName "Load site users $SiteUrl" -ScriptBlock {
        Get-SPOUser -Site $SiteUrl -Limit All -ErrorAction Stop
    })

    foreach ($siteUser in $siteUsers) {
        $loginName = ([string]$siteUser.LoginName).Trim()
        if ([string]::IsNullOrWhiteSpace($loginName)) {
            continue
        }

        $userType = ''
        if ($siteUser.PSObject.Properties.Name -contains 'UserType') {
            $userType = ([string]$siteUser.UserType).Trim()
        }

        $loginLower = $loginName.ToLowerInvariant()
        $isExternal = $loginLower.Contains('#ext#') -or $userType.Equals('Guest', [System.StringComparison]::OrdinalIgnoreCase)
        if (-not $isExternal) {
            continue
        }

        $emailValue = ''
        if ($siteUser.PSObject.Properties.Name -contains 'Email') {
            $emailValue = ([string]$siteUser.Email).Trim()
        }

        if ([string]::IsNullOrWhiteSpace($emailValue) -and $siteUser.PSObject.Properties.Name -contains 'UserPrincipalName') {
            $emailValue = ([string]$siteUser.UserPrincipalName).Trim()
        }

        $displayName = ''
        if ($siteUser.PSObject.Properties.Name -contains 'DisplayName') {
            $displayName = ([string]$siteUser.DisplayName).Trim()
        }

        $isSiteAdmin = $false
        if ($siteUser.PSObject.Properties.Name -contains 'IsSiteAdmin') {
            $isSiteAdmin = [bool]$siteUser.IsSiteAdmin
        }
        elseif ($siteUser.PSObject.Properties.Name -contains 'IsSiteCollectionAdmin') {
            $isSiteAdmin = [bool]$siteUser.IsSiteCollectionAdmin
        }

        $principals.Add([PSCustomObject]@{
                LoginName       = $loginName
                NormalizedLogin = Get-NormalizedLogin -LoginName $loginName
                Email           = $emailValue
                NormalizedEmail = Get-NormalizedEmail -Email $emailValue
                DisplayName     = $displayName
                UserType        = $userType
                IsSiteAdmin     = $isSiteAdmin
            })
    }

    return @($principals | Sort-Object -Property NormalizedEmail, NormalizedLogin)
}

$requiredHeaders = @(
    'UserPrincipalName',
    'ExternalUserEmails',
    'RevokeAllExternalUsers',
    'IncludeSiteCollectionAdmins'
)

Write-Status -Message 'Starting OneDrive external sharing link revocation script.'
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
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'RevokeOneDriveExternalSharingLinks' -Status 'NotFound' -Message "$($resolution.Message) Expected URL: $($resolution.ExpectedSiteUrl)"))
            $rowNumber++
            continue
        }

        if ($resolution.Status -eq 'Ambiguous') {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'RevokeOneDriveExternalSharingLinks' -Status 'Failed' -Message "$($resolution.Message) Expected URL: $($resolution.ExpectedSiteUrl)"))
            $rowNumber++
            continue
        }

        $site = $resolution.Site
        $siteUrl = ([string]$site.Url).Trim()

        $revokeAllExternalUsers = ConvertTo-Bool -Value $row.RevokeAllExternalUsers -Default $false
        $includeSiteCollectionAdmins = ConvertTo-Bool -Value $row.IncludeSiteCollectionAdmins -Default $false

        $targetIdentities = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in (ConvertTo-Array -Value ([string]$row.ExternalUserEmails))) {
            $normalized = Get-NormalizedEmail -Email $entry
            if (-not [string]::IsNullOrWhiteSpace($normalized)) {
                $null = $targetIdentities.Add($normalized)
            }

            $normalizedLogin = Get-NormalizedLogin -LoginName $entry
            if (-not [string]::IsNullOrWhiteSpace($normalizedLogin)) {
                $null = $targetIdentities.Add($normalizedLogin)
            }
        }

        if (-not $revokeAllExternalUsers -and $targetIdentities.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'RevokeOneDriveExternalSharingLinks' -Status 'Skipped' -Message 'No external identities were provided. Set RevokeAllExternalUsers=TRUE or populate ExternalUserEmails.'))
            $rowNumber++
            continue
        }

        $externalPrincipals = @(Get-ExternalSitePrincipals -SiteUrl $siteUrl)
        if ($externalPrincipals.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'RevokeOneDriveExternalSharingLinks' -Status 'Completed' -Message 'No external principals were found on this OneDrive site.'))
            $rowNumber++
            continue
        }

        $candidates = [System.Collections.Generic.List[object]]::new()
        foreach ($principal in $externalPrincipals) {
            if ($revokeAllExternalUsers) {
                $candidates.Add($principal)
                continue
            }

            if ($targetIdentities.Contains($principal.NormalizedEmail) -or $targetIdentities.Contains($principal.NormalizedLogin)) {
                $candidates.Add($principal)
            }
        }

        if (-not $includeSiteCollectionAdmins) {
            $filtered = [System.Collections.Generic.List[object]]::new()
            foreach ($candidate in $candidates) {
                if (-not $candidate.IsSiteAdmin) {
                    $filtered.Add($candidate)
                }
            }
            $candidates = $filtered
        }

        if ($candidates.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'RevokeOneDriveExternalSharingLinks' -Status 'NotFound' -Message 'No matching external principals were found for revocation.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($siteUrl, 'Remove OneDrive external principals and revoke associated sharing access')) {
            foreach ($candidate in $candidates) {
                Invoke-WithRetry -OperationName "Remove OneDrive external principal $($candidate.LoginName) from $siteUrl" -ScriptBlock {
                    Remove-SPOUser -Site $siteUrl -LoginName $candidate.LoginName -ErrorAction Stop
                }
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'RevokeOneDriveExternalSharingLinks' -Status 'Completed' -Message "Removed $($candidates.Count) external principal(s)."))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'RevokeOneDriveExternalSharingLinks' -Status 'WhatIf' -Message 'External sharing revocation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($userPrincipalName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'RevokeOneDriveExternalSharingLinks' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'OneDrive external sharing link revocation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
