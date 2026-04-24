<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260423-000000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Online.SharePoint.PowerShell
Microsoft.Graph.Authentication
Microsoft.Graph.Groups
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets SharePointSitesByUser and exports results to CSV.

.DESCRIPTION
    For each in-scope user, discovers every SharePoint Online site where they have access
    and identifies the access path that produced the entitlement. Multiple access paths to
    the same site are emitted as separate rows.
    Access paths evaluated (in order): SiteCollectionAdmin, SPGroupMember, M365GroupMember /
    TeamsGroupMember, and DirectPermission. Microsoft Graph is used inline to resolve M365
    group membership and check for an associated Microsoft Team.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all licensed users in the tenant (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all licensed users in the tenant rather than processing from an input CSV file.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).
    Required for both parameter sets; used to connect to SPO and enumerate all sites.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-SPOL-0500-Get-SharePointSitesByUser.ps1 -InputCsvPath .\Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
    Inventory SharePoint site access for users listed in the input CSV file.

.EXAMPLE
    .\D-SPOL-0500-Get-SharePointSitesByUser.ps1 -DiscoverAll -SharePointAdminUrl https://contoso-admin.sharepoint.com
    Discover and inventory SharePoint site access for all licensed users in the tenant.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Online.SharePoint.PowerShell, Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users
    Required roles:   SharePoint Administrator, Global Reader
    Graph scopes:     GroupMember.Read.All, Sites.Read.All
    Limitations:      None known.

    CSV Fields:
    Column              Type    Required  Description
    ------              ----    --------  -----------
    UserPrincipalName   String  Yes       UPN of the user whose SharePoint site access to inventory
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(Mandatory)]
    [ValidateScript({
        if ($_ -match '^https://' -and $_ -match 'sharepoint') { $true }
        else { throw "SharePointAdminUrl must start with https:// and contain 'sharepoint'." }
    })]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-SPOL-0500-Get-SharePointSitesByUser_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
        [Parameter(Mandatory)] [int]$RowNumber,
        [Parameter(Mandatory)] [string]$PrimaryKey,
        [Parameter(Mandatory)] [string]$Action,
        [Parameter(Mandatory)] [string]$Status,
        [Parameter(Mandatory)] [string]$Message,
        [Parameter(Mandatory)] [hashtable]$Data
    )

    $base = New-ResultObject -RowNumber $RowNumber -PrimaryKey $PrimaryKey -Action $Action -Status $Status -Message $Message
    $ordered = [ordered]@{}
    foreach ($prop in $base.PSObject.Properties.Name) { $ordered[$prop] = $base.$prop }
    foreach ($key in $Data.Keys) { $ordered[$key] = $Data[$key] }
    return [PSCustomObject]$ordered
}

function New-EmptySiteAccessData {
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName = '',
        [string]$SiteUrl           = '',
        [string]$SiteTitle         = '',
        [string]$SiteTemplate      = ''
    )

    return [ordered]@{
        UserPrincipalName   = $UserPrincipalName
        SiteUrl             = $SiteUrl
        SiteTitle           = $SiteTitle
        SiteTemplate        = $SiteTemplate
        AccessPath          = ''
        PermissionLevel     = ''
        SPGroupName         = ''
        M365GroupId         = ''
        M365GroupDisplayName = ''
        TeamConnected       = ''
        TeamId              = ''
        AssignmentChain     = ''
    }
}

$requiredHeaders = @('UserPrincipalName')

$reportPropertyOrder = @(
    'TimestampUtc', 'RowNumber', 'PrimaryKey', 'Action', 'Status', 'Message', 'ScopeMode',
    'UserPrincipalName', 'SiteUrl', 'SiteTitle', 'SiteTemplate',
    'AccessPath', 'PermissionLevel', 'SPGroupName',
    'M365GroupId', 'M365GroupDisplayName', 'TeamConnected', 'TeamId',
    'AssignmentChain'
)

Write-Status -Message 'Starting SharePoint Online sites by user discovery script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Online.SharePoint.PowerShell', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users')
Ensure-SharePointConnection -SharePointAdminUrl $SharePointAdminUrl
Ensure-GraphConnection -RequiredScopes @('GroupMember.Read.All', 'Sites.Read.All')

$scopeMode = 'Csv'
if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. Enumerating all licensed users in the tenant.'

    $licensedUsers = @(Invoke-WithRetry -OperationName 'Enumerate all licensed users' -ScriptBlock {
        Get-MgUser `
            -Filter "assignedLicenses/`$count ne 0" `
            -CountVariable lc `
            -ConsistencyLevel eventual `
            -All `
            -Property 'Id,UserPrincipalName,DisplayName' `
            -ErrorAction Stop
    })

    Write-Status -Message "DiscoverAll: found $($licensedUsers.Count) licensed users."
    $rows = @($licensedUsers | ForEach-Object { [PSCustomObject]@{ UserPrincipalName = $_.UserPrincipalName } })
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

# ── Enumerate and cache all SharePoint sites once before the user loop ────────────────────
Write-Status -Message 'Enumerating all SharePoint Online sites. This may take several minutes for large tenants.'

$allSites = @(Invoke-WithRetry -OperationName 'Enumerate all SPO sites' -ScriptBlock {
    Get-SPOSite -Limit All -IncludePersonalSite $false `
        -ErrorAction Stop |
        Select-Object Url, Title, Template, StorageQuota, GroupId
})

# Filter out personal/OneDrive sites
$allSites = @($allSites | Where-Object {
    $_.Template -notlike 'SPSPERS*' -and
    $_.Url -notlike '*-my.sharepoint.com*'
})

Write-Status -Message "Site enumeration complete. $($allSites.Count) SharePoint sites in scope."

# Cache for M365 group info: GroupId -> [DisplayName, TeamId (or empty)]
$m365GroupInfoCache = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $upn = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw 'UserPrincipalName is required and cannot be blank.'
        }

        Write-Status -Message "Processing user $rowNumber of $($rows.Count): $upn"

        # ── Pre-cache the user's Graph ObjectId ───────────────────────────────────
        $userObj = Invoke-WithRetry -OperationName "Resolve Entra user $upn" -ScriptBlock {
            Get-MgUser -UserId $upn -Property 'Id,UserPrincipalName,DisplayName' -ErrorAction Stop
        }
        $userId = ([string]$userObj.Id).Trim()

        $accessRows = [System.Collections.Generic.List[object]]::new()

        foreach ($site in $allSites) {
            $siteUrl      = ([string]$site.Url).Trim().TrimEnd('/')
            $siteTitle    = ([string]$site.Title).Trim()
            $siteTemplate = ([string]$site.Template).Trim()
            $siteGroupId  = ([string]$site.GroupId).Trim()

            # Track which access paths have already been recorded for this site/user
            # to avoid duplicate DirectPermission rows when a user was already captured
            # through SPGroupMember.
            $capturedPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

            # ── Path A — Site Collection Admin ────────────────────────────────────
            try {
                $spoUserSca = Invoke-WithRetry -OperationName "Get SPO user (SCA check) $upn on $siteUrl" -ScriptBlock {
                    Get-SPOUser -Site $siteUrl -LoginName $upn -ErrorAction SilentlyContinue
                }

                if ($null -ne $spoUserSca -and $spoUserSca.IsSiteAdmin -eq $true) {
                    $accessRows.Add([ordered]@{
                        UserPrincipalName    = $upn
                        SiteUrl              = $siteUrl
                        SiteTitle            = $siteTitle
                        SiteTemplate         = $siteTemplate
                        AccessPath           = 'SiteCollectionAdmin'
                        PermissionLevel      = 'Site Collection Administrator'
                        SPGroupName          = ''
                        M365GroupId          = ''
                        M365GroupDisplayName = ''
                        TeamConnected        = ''
                        TeamId               = ''
                        AssignmentChain      = "$upn > Site Collection Admin > $siteTitle"
                    })
                    [void]$capturedPaths.Add('SiteCollectionAdmin')
                }
            }
            catch {
                Write-Status -Message "Could not check SCA status for $upn on $siteUrl`: $($_.Exception.Message)" -Level WARN
            }

            # ── Path B — SharePoint Site Group Membership ─────────────────────────
            try {
                $siteGroups = @(Invoke-WithRetry -OperationName "Get SPO site groups for $siteUrl" -ScriptBlock {
                    Get-SPOSiteGroup -Site $siteUrl -ErrorAction Stop
                })

                foreach ($group in $siteGroups) {
                    $groupTitle = ([string]$group.Title).Trim()
                    $groupRoles = ([string[]]@($group.Roles) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join '; '

                    try {
                        $groupMembers = @(Invoke-WithRetry -OperationName "Get members of SPO group '$groupTitle' on $siteUrl" -ScriptBlock {
                            Get-SPOUser -Site $siteUrl -Group $groupTitle -ErrorAction Stop
                        })

                        $userInGroup = $groupMembers | Where-Object {
                            ([string]$_.LoginName).Trim() -eq $upn -or
                            ([string]$_.UserPrincipalName).Trim() -eq $upn
                        }

                        if ($null -ne $userInGroup) {
                            $accessRows.Add([ordered]@{
                                UserPrincipalName    = $upn
                                SiteUrl              = $siteUrl
                                SiteTitle            = $siteTitle
                                SiteTemplate         = $siteTemplate
                                AccessPath           = 'SPGroupMember'
                                PermissionLevel      = $groupRoles
                                SPGroupName          = $groupTitle
                                M365GroupId          = ''
                                M365GroupDisplayName = ''
                                TeamConnected        = ''
                                TeamId               = ''
                                AssignmentChain      = "$upn > SP Group: $groupTitle > $siteTitle"
                            })
                            [void]$capturedPaths.Add('SPGroupMember')
                        }
                    }
                    catch {
                        Write-Status -Message "Could not enumerate members of SPO group '$groupTitle' on $siteUrl`: $($_.Exception.Message)" -Level WARN
                    }
                }
            }
            catch {
                Write-Status -Message "Could not enumerate SPO site groups for $siteUrl`: $($_.Exception.Message)" -Level WARN
            }

            # ── Path C — M365 Group / Teams Connected Site ────────────────────────
            if (-not [string]::IsNullOrWhiteSpace($siteGroupId)) {
                try {
                    # Check if the user is a member of the M365 group
                    $groupMembers = @(Invoke-WithRetry -OperationName "Get M365 group members for $siteGroupId" -ScriptBlock {
                        Get-MgGroupMember -GroupId $siteGroupId -All -ErrorAction Stop
                    })

                    $userInM365Group = $groupMembers | Where-Object {
                        ([string]$_.Id).Trim() -eq $userId
                    }

                    if ($null -ne $userInM365Group) {
                        # Resolve M365 group display name and Team info (cached)
                        $m365GroupDisplayName = ''
                        $teamConnected        = 'False'
                        $teamId               = ''

                        if ($m365GroupInfoCache.ContainsKey($siteGroupId)) {
                            $cachedInfo           = $m365GroupInfoCache[$siteGroupId]
                            $m365GroupDisplayName = ([string]$cachedInfo.DisplayName).Trim()
                            $teamConnected        = ([string]$cachedInfo.TeamConnected).Trim()
                            $teamId               = ([string]$cachedInfo.TeamId).Trim()
                        }
                        else {
                            # Fetch group display name
                            try {
                                $m365GroupObj = Invoke-WithRetry -OperationName "Get M365 group info for $siteGroupId" -ScriptBlock {
                                    Get-MgGroup -GroupId $siteGroupId -Property 'Id,DisplayName' -ErrorAction Stop
                                }
                                $m365GroupDisplayName = ([string]$m365GroupObj.DisplayName).Trim()
                            }
                            catch {
                                Write-Status -Message "Could not resolve M365 group display name for $siteGroupId`: $($_.Exception.Message)" -Level WARN
                            }

                            # Check for an associated Team
                            try {
                                $teamObj = Invoke-WithRetry -OperationName "Get Team for M365 group $siteGroupId" -ScriptBlock {
                                    Get-MgGroupTeam -GroupId $siteGroupId -ErrorAction SilentlyContinue
                                }

                                if ($null -ne $teamObj -and -not [string]::IsNullOrWhiteSpace($teamObj.Id)) {
                                    $teamConnected = 'True'
                                    $teamId        = ([string]$teamObj.Id).Trim()
                                }
                            }
                            catch {
                                # No team or not accessible — leave teamConnected as 'False'
                            }

                            $m365GroupInfoCache[$siteGroupId] = [PSCustomObject]@{
                                DisplayName   = $m365GroupDisplayName
                                TeamConnected = $teamConnected
                                TeamId        = $teamId
                            }
                        }

                        if ($teamConnected -eq 'True') {
                            $accessPath     = 'TeamsGroupMember'
                            $assignmentChain = "$upn > Teams: $teamId > M365 Group: $m365GroupDisplayName > $siteTitle"
                        }
                        else {
                            $accessPath     = 'M365GroupMember'
                            $assignmentChain = "$upn > M365 Group: $m365GroupDisplayName > $siteTitle"
                        }

                        $accessRows.Add([ordered]@{
                            UserPrincipalName    = $upn
                            SiteUrl              = $siteUrl
                            SiteTitle            = $siteTitle
                            SiteTemplate         = $siteTemplate
                            AccessPath           = $accessPath
                            PermissionLevel      = ''
                            SPGroupName          = ''
                            M365GroupId          = $siteGroupId
                            M365GroupDisplayName = $m365GroupDisplayName
                            TeamConnected        = $teamConnected
                            TeamId               = $teamId
                            AssignmentChain      = $assignmentChain
                        })
                        [void]$capturedPaths.Add($accessPath)
                    }
                }
                catch {
                    Write-Status -Message "Could not check M365 group membership for group $siteGroupId on $siteUrl`: $($_.Exception.Message)" -Level WARN
                }
            }

            # ── Path D — Direct Permission ─────────────────────────────────────────
            # Only emit if the user has a SPO user entry, is not SCA, and was not
            # already captured via SPGroupMember (which covers group-based assignments).
            if (-not $capturedPaths.Contains('SPGroupMember')) {
                try {
                    $spoUserDirect = Invoke-WithRetry -OperationName "Get SPO user (direct check) $upn on $siteUrl" -ScriptBlock {
                        Get-SPOUser -Site $siteUrl -LoginName $upn -ErrorAction SilentlyContinue
                    }

                    if ($null -ne $spoUserDirect -and $spoUserDirect.IsSiteAdmin -ne $true) {
                        # Build permission level from Groups or Roles if available
                        $directPermLevel = ''
                        $groupsList = $spoUserDirect.Groups
                        if ($null -ne $groupsList -and @($groupsList).Count -gt 0) {
                            $directPermLevel = (@($groupsList) | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join '; '
                        }

                        if (-not [string]::IsNullOrWhiteSpace($directPermLevel) -or $capturedPaths.Count -eq 0) {
                            $accessRows.Add([ordered]@{
                                UserPrincipalName    = $upn
                                SiteUrl              = $siteUrl
                                SiteTitle            = $siteTitle
                                SiteTemplate         = $siteTemplate
                                AccessPath           = 'DirectPermission'
                                PermissionLevel      = $directPermLevel
                                SPGroupName          = ''
                                M365GroupId          = ''
                                M365GroupDisplayName = ''
                                TeamConnected        = ''
                                TeamId               = ''
                                AssignmentChain      = "$upn > Direct Permission > $siteTitle"
                            })
                        }
                    }
                }
                catch {
                    Write-Status -Message "Could not check direct permission for $upn on $siteUrl`: $($_.Exception.Message)" -Level WARN
                }
            }
        }

        # ── Emit results for this user ─────────────────────────────────────────────
        if ($accessRows.Count -eq 0) {
            $results.Add((New-InventoryResult `
                -RowNumber  $rowNumber `
                -PrimaryKey $upn `
                -Action     'GetSharePointSitesByUser' `
                -Status     'Success' `
                -Message    'No SharePoint site access found.' `
                -Data       (New-EmptySiteAccessData -UserPrincipalName $upn)))
        }
        else {
            foreach ($accessRow in $accessRows) {
                $pk = "$upn|$($accessRow.SiteUrl)|$($accessRow.AccessPath)"
                $results.Add((New-InventoryResult `
                    -RowNumber  $rowNumber `
                    -PrimaryKey $pk `
                    -Action     'GetSharePointSitesByUser' `
                    -Status     'Success' `
                    -Message    'SharePoint site access entitlement exported.' `
                    -Data       ([ordered]@{
                        UserPrincipalName    = $accessRow.UserPrincipalName
                        SiteUrl              = $accessRow.SiteUrl
                        SiteTitle            = $accessRow.SiteTitle
                        SiteTemplate         = $accessRow.SiteTemplate
                        AccessPath           = $accessRow.AccessPath
                        PermissionLevel      = $accessRow.PermissionLevel
                        SPGroupName          = $accessRow.SPGroupName
                        M365GroupId          = $accessRow.M365GroupId
                        M365GroupDisplayName = $accessRow.M365GroupDisplayName
                        TeamConnected        = $accessRow.TeamConnected
                        TeamId               = $accessRow.TeamId
                        AssignmentChain      = $accessRow.AssignmentChain
                    })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult `
            -RowNumber  $rowNumber `
            -PrimaryKey $upn `
            -Action     'GetSharePointSitesByUser' `
            -Status     'Failed' `
            -Message    $_.Exception.Message `
            -Data       (New-EmptySiteAccessData -UserPrincipalName $upn)))
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
Write-Status -Message 'SharePoint Online sites by user discovery script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
