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
    Exports site-level permissions for SharePoint Online site collections.

.DESCRIPTION
    Exports all site-level permissions for each site collection: unique permissions flag,
    site group memberships, site collection administrators, access request settings, and
    sharing capability. One row per principal-permission pair.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all site collections in the tenant (-DiscoverAll parameter set).
    All results — including sites that could not be queried — are written to the output CSV.
    Uses PnP.PowerShell for detailed permission access.

.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include SiteUrl.
    See the companion .input.csv template for the full column list.

.PARAMETER DiscoverAll
    Enumerate all site collections in the tenant rather than processing from an input CSV file.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).
    Required for DiscoverAll mode to enumerate all site collections.
    Also used for the initial PnP tenant connection.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-SPOL-0020-Get-SharePointSitePermissions.ps1 -InputCsvPath .\D-SPOL-0020-Get-SharePointSitePermissions.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com

    Export permissions for sites listed in the input CSV.

.EXAMPLE
    .\D-SPOL-0020-Get-SharePointSitePermissions.ps1 -DiscoverAll -SharePointAdminUrl https://los-admin.sharepoint.com

    Export permissions for all site collections in the tenant.

.NOTES
    Version:          1.0
    Required modules: PnP.PowerShell
    Required roles:   SharePoint Administrator
    Limitations:      Requires PnP.PowerShell and interactive authentication (or service principal).
                      Permission output covers site-level only — subweb-level unique permissions
                      are not reported by this script.
                      SCA = Site Collection Administrator.

    CSV Fields:
    Column      Type      Required  Description
    ----------  --------  --------  -----------
    SiteUrl     String    Yes       Absolute URL of the site collection (e.g. https://tenant.sharepoint.com/sites/Finance)
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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-SPOL-0020-Get-SharePointSitePermissions_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Connect-PnPSite {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Url)

    Connect-PnPOnline -Url $Url -Interactive -ErrorAction Stop
}

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

function New-EmptyPermissionData {
    [CmdletBinding()]
    param([string]$SiteUrlRequested = '')

    return [ordered]@{
        SiteUrlRequested      = $SiteUrlRequested
        SiteUrl               = ''
        SiteTitle             = ''
        HasUniquePermissions  = ''
        SharingCapability     = ''
        AccessRequestEnabled  = ''
        AccessRequestUrl      = ''
        PermissionType        = ''
        PrincipalType         = ''
        PrincipalLoginName    = ''
        PrincipalDisplayName  = ''
        PermissionLevel       = ''
        GroupName             = ''
        IsHiddenInUI          = ''
    }
}

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'SiteUrlRequested',
    'SiteUrl',
    'SiteTitle',
    'HasUniquePermissions',
    'SharingCapability',
    'AccessRequestEnabled',
    'AccessRequestUrl',
    'PermissionType',
    'PrincipalType',
    'PrincipalLoginName',
    'PrincipalDisplayName',
    'PermissionLevel',
    'GroupName',
    'IsHiddenInUI'
)

$requiredHeaders = @('SiteUrl')

Write-Status -Message 'Starting SharePoint site permissions discovery script.'
Assert-ModuleCurrent -ModuleNames @('PnP.PowerShell')

$adminUrlTrimmed = $SharePointAdminUrl.Trim().TrimEnd('/')
if ($adminUrlTrimmed -notmatch '^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com$') {
    throw "SharePointAdminUrl '$adminUrlTrimmed' is invalid. Use: https://<tenant>-admin.sharepoint.com"
}

$scopeMode = 'Csv'

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. Connecting to admin center to enumerate sites.' -Level WARN
    Connect-PnPSite -Url $adminUrlTrimmed

    $allSites = Invoke-WithRetry -OperationName 'Get all tenant sites' -ScriptBlock {
        Get-PnPTenantSite -IncludeOneDriveSites:$false -ErrorAction Stop | Select-Object Url
    }

    Write-Status -Message "Found $($allSites.Count) site collections."
    $rows = @($allSites | ForEach-Object { [PSCustomObject]@{ SiteUrl = [string]$_.Url } })
} else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $siteUrl    = ([string]$row.SiteUrl).Trim().TrimEnd('/')
    $primaryKey = $siteUrl

    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        Write-Status -Message "Row $rowNumber skipped: SiteUrl is empty." -Level WARN
        $rowNumber++
        continue
    }

    try {
        Write-Status -Message "Connecting to site: $siteUrl"
        Connect-PnPSite -Url $siteUrl

        # Get site properties.
        $pnpSite = Invoke-WithRetry -OperationName "Get site $siteUrl" -ScriptBlock {
            Get-PnPSite -Includes HasUniqueRoleAssignments, Title, Url, SharingCapability, AccessRequestEnabled, AccessRequestUrl -ErrorAction Stop
        }

        $siteTitle           = Get-TrimmedValue -Value $pnpSite.Title
        $hasUniquePerms      = [string]$pnpSite.HasUniqueRoleAssignments
        $sharingCapability   = [string]$pnpSite.SharingCapability
        $accessRequestEnabled = [string]$pnpSite.AccessRequestEnabled
        $accessRequestUrl    = Get-TrimmedValue -Value $pnpSite.AccessRequestUrl

        $siteRowData = [ordered]@{
            SiteUrlRequested     = [string]$row.SiteUrl
            SiteUrl              = $siteUrl
            SiteTitle            = $siteTitle
            HasUniquePermissions = $hasUniquePerms
            SharingCapability    = $sharingCapability
            AccessRequestEnabled = $accessRequestEnabled
            AccessRequestUrl     = $accessRequestUrl
        }

        # --- Site Collection Administrators ---
        $scaList = Invoke-WithRetry -OperationName "Get SCAs for $siteUrl" -ScriptBlock {
            Get-PnPSiteCollectionAdmin -ErrorAction Stop
        }

        foreach ($sca in $scaList) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "${siteUrl}|SCA|$([string]$sca.LoginName)" -Action 'GetSharePointSitePermission' -Status 'Completed' -Message 'Site collection administrator exported.' -Data ([ordered]@{
                SiteUrlRequested     = $siteRowData.SiteUrlRequested
                SiteUrl              = $siteRowData.SiteUrl
                SiteTitle            = $siteRowData.SiteTitle
                HasUniquePermissions = $siteRowData.HasUniquePermissions
                SharingCapability    = $siteRowData.SharingCapability
                AccessRequestEnabled = $siteRowData.AccessRequestEnabled
                AccessRequestUrl     = $siteRowData.AccessRequestUrl
                PermissionType       = 'SiteCollectionAdmin'
                PrincipalType        = [string]$sca.PrincipalType
                PrincipalLoginName   = Get-TrimmedValue -Value $sca.LoginName
                PrincipalDisplayName = Get-TrimmedValue -Value $sca.Title
                PermissionLevel      = 'SiteCollectionAdministrator'
                GroupName            = ''
                IsHiddenInUI         = ''
            })))
        }

        # --- Site Groups and their Members ---
        $groups = Invoke-WithRetry -OperationName "Get groups for $siteUrl" -ScriptBlock {
            Get-PnPGroup -ErrorAction Stop
        }

        foreach ($group in $groups) {
            $groupName    = Get-TrimmedValue -Value $group.Title
            $permLevel    = ''
            $isHiddenInUI = [string]$group.OnlyAllowMembersViewMembership

            # Get the permission level assigned to this group on the site root.
            try {
                $roleAssignments = Get-PnPGroupPermissions -Identity $group.Id -ErrorAction SilentlyContinue
                if ($roleAssignments) {
                    $permLevel = ($roleAssignments | ForEach-Object { [string]$_.Name } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join '; '
                }
            }
            catch { $permLevel = '' }

            # Get group members.
            $members = @()
            try {
                $members = @(Get-PnPGroupMember -Group $group.Id -ErrorAction SilentlyContinue)
            }
            catch { $members = @() }

            if ($members.Count -eq 0) {
                # Empty group row.
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "${siteUrl}|Group|${groupName}|(empty)" -Action 'GetSharePointSitePermission' -Status 'Completed' -Message 'Site group has no members.' -Data ([ordered]@{
                    SiteUrlRequested     = $siteRowData.SiteUrlRequested
                    SiteUrl              = $siteRowData.SiteUrl
                    SiteTitle            = $siteRowData.SiteTitle
                    HasUniquePermissions = $siteRowData.HasUniquePermissions
                    SharingCapability    = $siteRowData.SharingCapability
                    AccessRequestEnabled = $siteRowData.AccessRequestEnabled
                    AccessRequestUrl     = $siteRowData.AccessRequestUrl
                    PermissionType       = 'GroupMember'
                    PrincipalType        = ''
                    PrincipalLoginName   = ''
                    PrincipalDisplayName = ''
                    PermissionLevel      = $permLevel
                    GroupName            = $groupName
                    IsHiddenInUI         = $isHiddenInUI
                })))
            } else {
                foreach ($member in $members) {
                    $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "${siteUrl}|Group|${groupName}|$([string]$member.LoginName)" -Action 'GetSharePointSitePermission' -Status 'Completed' -Message 'Site group member exported.' -Data ([ordered]@{
                        SiteUrlRequested     = $siteRowData.SiteUrlRequested
                        SiteUrl              = $siteRowData.SiteUrl
                        SiteTitle            = $siteRowData.SiteTitle
                        HasUniquePermissions = $siteRowData.HasUniquePermissions
                        SharingCapability    = $siteRowData.SharingCapability
                        AccessRequestEnabled = $siteRowData.AccessRequestEnabled
                        AccessRequestUrl     = $siteRowData.AccessRequestUrl
                        PermissionType       = 'GroupMember'
                        PrincipalType        = [string]$member.PrincipalType
                        PrincipalLoginName   = Get-TrimmedValue -Value $member.LoginName
                        PrincipalDisplayName = Get-TrimmedValue -Value $member.Title
                        PermissionLevel      = $permLevel
                        GroupName            = $groupName
                        IsHiddenInUI         = $isHiddenInUI
                    })))
                }
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $emptyData = New-EmptyPermissionData -SiteUrlRequested ([string]$row.SiteUrl)
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetSharePointSitePermission' -Status 'Failed' -Message $_.Exception.Message -Data $emptyData))
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
Write-Status -Message 'SharePoint site permissions discovery script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
