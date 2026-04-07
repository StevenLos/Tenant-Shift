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
    Exports all SharePoint Online sites where specified users hold permissions.

.DESCRIPTION
    For each user in the input CSV, queries all SharePoint Online site collections
    and identifies where the user appears as a site collection admin, a member of
    a site group, or a direct permission assignment. Outputs one row per
    user-site-permission combination.
    Requires -SharePointAdminUrl to connect to the admin centre for site enumeration.
    All results — including users with no site access found — are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include UserPrincipalName.
    See the companion .input.csv template for the full column list.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).
    Required to enumerate all site collections when searching for user permissions.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-SPOL-0030-Get-SharePointSitesByUser.ps1 -InputCsvPath .\D-SPOL-0030-Get-SharePointSitesByUser.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com

    Export all sites where the listed users have permissions.

.NOTES
    Version:          1.0
    Required modules: PnP.PowerShell
    Required roles:   SharePoint Administrator
    Limitations:      Scans all site collections — may be slow on large tenants.
                      Only detects explicit group membership and site collection admin assignments.
                      Does not detect inherited permissions or sharing links.
                      Requires PnP.PowerShell and interactive authentication (or service principal).

    CSV Fields:
    Column              Type      Required  Description
    ------------------  --------  --------  -----------
    UserPrincipalName   String    Yes       UPN of the user to find site permissions for
#>
#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [ValidatePattern('^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$')]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-SPOL-0030-Get-SharePointSitesByUser_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'UserPrincipalName',
    'SiteUrl',
    'SiteTitle',
    'PermissionSource',
    'GroupName',
    'PermissionLevel'
)

$requiredHeaders = @('UserPrincipalName')

Write-Status -Message 'Starting SharePoint sites-by-user export script.'
Assert-ModuleCurrent -ModuleNames @('PnP.PowerShell')

$adminUrlTrimmed = $SharePointAdminUrl.TrimEnd('/')

# Connect to admin centre for site enumeration.
Write-Status -Message "Connecting to SharePoint admin centre: $adminUrlTrimmed"
Connect-PnPSite -Url $adminUrlTrimmed

Write-Status -Message 'Fetching all site collections.'
$allSites = Invoke-WithRetry -OperationName 'Get all tenant sites' -ScriptBlock {
    Get-PnPTenantSite -ErrorAction Stop | Select-Object Url, Title
}
Write-Status -Message "Found $($allSites.Count) site collections."

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $upn        = Get-TrimmedValue -Value $row.UserPrincipalName
    $primaryKey = $upn

    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Status -Message "Row $rowNumber skipped: UserPrincipalName is empty." -Level WARN
        $rowNumber++
        continue
    }

    $upnLower     = $upn.ToLowerInvariant()
    $matchesFound = 0

    try {
        foreach ($site in $allSites) {
            $siteUrl   = $site.Url
            $siteTitle = $site.Title

            try {
                Connect-PnPSite -Url $siteUrl

                # Check site collection admins.
                $admins = Invoke-WithRetry -OperationName "Get site collection admins for $siteUrl" -ScriptBlock {
                    Get-PnPSiteCollectionAdmin -ErrorAction Stop
                }

                foreach ($admin in $admins) {
                    $adminUpn = if ($admin.Email) { $admin.Email.Trim().ToLowerInvariant() } else { [string]$admin.LoginName }
                    if ($adminUpn -ieq $upnLower) {
                        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetSharePointSitesByUser' -Status 'Completed' -Message 'User found as site collection admin.' -Data ([ordered]@{
                            UserPrincipalName = $upn
                            SiteUrl           = $siteUrl
                            SiteTitle         = $siteTitle
                            PermissionSource  = 'SiteCollectionAdmin'
                            GroupName         = ''
                            PermissionLevel   = 'SiteCollectionAdmin'
                        })))
                        $matchesFound++
                    }
                }

                # Check site groups and their members.
                $groups = Invoke-WithRetry -OperationName "Get site groups for $siteUrl" -ScriptBlock {
                    Get-PnPGroup -ErrorAction Stop
                }

                foreach ($group in $groups) {
                    $members = Invoke-WithRetry -OperationName "Get members of group '$($group.Title)' on $siteUrl" -ScriptBlock {
                        Get-PnPGroupMember -Group $group.Title -ErrorAction Stop
                    }

                    foreach ($member in $members) {
                        $memberUpn = if ($member.Email) { $member.Email.Trim().ToLowerInvariant() } else { [string]$member.LoginName }
                        if ($memberUpn -ieq $upnLower) {
                            # Resolve permission levels for this group.
                            $permLevels = Invoke-WithRetry -OperationName "Get permissions for group '$($group.Title)' on $siteUrl" -ScriptBlock {
                                Get-PnPGroupPermissions -Identity $group.Title -ErrorAction Stop
                            }
                            $permLevel = if ($permLevels) { ($permLevels | Select-Object -ExpandProperty Name) -join '; ' } else { '' }

                            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetSharePointSitesByUser' -Status 'Completed' -Message "User found as member of group '$($group.Title)'." -Data ([ordered]@{
                                UserPrincipalName = $upn
                                SiteUrl           = $siteUrl
                                SiteTitle         = $siteTitle
                                PermissionSource  = 'GroupMember'
                                GroupName         = $group.Title
                                PermissionLevel   = $permLevel
                            })))
                            $matchesFound++
                        }
                    }
                }
            }
            catch {
                Write-Status -Message "  Site $siteUrl skipped for user $upn: $($_.Exception.Message)" -Level WARN
            }
        }

        if ($matchesFound -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetSharePointSitesByUser' -Status 'Completed' -Message 'No site permissions found for this user.' -Data ([ordered]@{
                UserPrincipalName = $upn
                SiteUrl           = ''
                SiteTitle         = ''
                PermissionSource  = ''
                GroupName         = ''
                PermissionLevel   = ''
            })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetSharePointSitesByUser' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            UserPrincipalName = $upn; SiteUrl = ''; SiteTitle = ''; PermissionSource = ''; GroupName = ''; PermissionLevel = ''
        })))
    }

    $rowNumber++
}

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint sites-by-user export script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
