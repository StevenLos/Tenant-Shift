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

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [switch]$AllowLastAdminRemoval,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M3241-Set-SharePointSiteAdmins_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {


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

function Get-SpoSiteOrThrow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    try {
        return Invoke-WithRetry -OperationName "Lookup SharePoint site $SiteUrl" -ScriptBlock {
            Get-SPOSite -Identity $SiteUrl -Detailed -ErrorAction Stop
        }
    }
    catch {
        $message = ([string]$_.Exception.Message).ToLowerInvariant()
        if ($message -match 'cannot find|was not found|does not exist|not found') {
            throw "Site '$SiteUrl' was not found."
        }

        throw
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
    'SiteUrl',
    'AddSiteAdmins',
    'RemoveSiteAdmins',
    'EnsurePrimaryOwnerIsAdmin'
)

Write-Status -Message 'Starting SharePoint site admin update script.'
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
            throw 'SiteUrl is required.'
        }

        $site = Get-SpoSiteOrThrow -SiteUrl $siteUrl
        $primaryOwnerNormalized = Get-NormalizedLogin -LoginName ([string]$site.Owner)

        $ensurePrimaryOwnerIsAdmin = ConvertTo-Bool -Value $row.EnsurePrimaryOwnerIsAdmin -Default $true
        $addAdminsRaw = ConvertTo-Array -Value ([string]$row.AddSiteAdmins)
        $removeAdminsRaw = ConvertTo-Array -Value ([string]$row.RemoveSiteAdmins)

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

        if ($ensurePrimaryOwnerIsAdmin -and -not [string]::IsNullOrWhiteSpace($primaryOwnerNormalized)) {
            $null = $addAdmins.Add($primaryOwnerNormalized)
            if ($removeAdmins.Contains($primaryOwnerNormalized)) {
                $null = $removeAdmins.Remove($primaryOwnerNormalized)
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
            throw 'Requested change would remove the last site collection administrator. Use -AllowLastAdminRemoval to override.'
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
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'SetSPOSiteAdmins' -Status 'Skipped' -Message 'Site admin membership already matches requested state.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($siteUrl, 'Update SharePoint site collection administrators')) {
            foreach ($adminToAdd in $adminsToAdd) {
                Invoke-WithRetry -OperationName "Grant site admin $adminToAdd on $siteUrl" -ScriptBlock {
                    Set-SPOUser -Site $siteUrl -LoginName $adminToAdd -IsSiteCollectionAdmin $true -ErrorAction Stop
                }
            }

            foreach ($adminToRemove in $adminsToRemove) {
                $loginToRemove = $currentAdminMap[$adminToRemove]
                Invoke-WithRetry -OperationName "Remove site admin $adminToRemove on $siteUrl" -ScriptBlock {
                    Set-SPOUser -Site $siteUrl -LoginName $loginToRemove -IsSiteCollectionAdmin $false -ErrorAction Stop
                }
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'SetSPOSiteAdmins' -Status 'Completed' -Message "Admins added: $($adminsToAdd.Count). Admins removed: $($adminsToRemove.Count)."))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'SetSPOSiteAdmins' -Status 'WhatIf' -Message 'Admin update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($siteUrl) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'SetSPOSiteAdmins' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint site admin update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







