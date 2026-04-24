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
Microsoft.Graph.Authentication
Microsoft.Graph.Groups
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets EntraAssignedGroupsByUser and exports results to CSV.

.DESCRIPTION
    Gets EntraAssignedGroupsByUser from Microsoft 365 and writes the results to a CSV file.
    For each in-scope user, finds all assigned (non-dynamic, non-M365) Entra ID security groups
    they belong to — directly or transitively through nested groups.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all licensed users in the tenant (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all licensed users in the tenant rather than processing from an input CSV file.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-MEID-0500-Get-EntraAssignedGroupsByUser.ps1 -InputCsvPath .\Scope-Users.input.csv
    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\D-MEID-0500-Get-EntraAssignedGroupsByUser.ps1 -DiscoverAll
    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users
    Required roles:   Global Reader
    Limitations:      None known.

    CSV Fields:
    Column              Type    Required  Description
    ------              ----    --------  -----------
    UserPrincipalName   String  Yes       UPN of the user whose group memberships to inventory
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-MEID-0500-Get-EntraAssignedGroupsByUser_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Test-IsAssignedSecurityGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Group
    )

    $groupTypes = @()
    if ($null -ne $Group.GroupTypes) {
        $groupTypes = @($Group.GroupTypes)
    }

    $isDynamic  = $groupTypes -contains 'DynamicMembership'
    $isUnified  = $groupTypes -contains 'Unified'
    $isAssigned = ($Group.SecurityEnabled -eq $true -and
                   $Group.MailEnabled    -eq $false -and
                   -not $isDynamic -and
                   -not $isUnified)
    return $isAssigned
}

function Get-MembershipPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Upn,
        [Parameter(Mandatory)] [string]$TargetGroupId,
        [Parameter(Mandatory)] [string]$TargetGroupDisplayName,
        [Parameter(Mandatory)] [System.Collections.Generic.HashSet[string]]$DirectMemberIds,
        [Parameter(Mandatory)] [System.Collections.Generic.Dictionary[string, object]]$ParentGroupCache
    )

    # If direct member, path is trivial.
    if ($DirectMemberIds.Contains($TargetGroupId)) {
        return "$Upn > $TargetGroupDisplayName"
    }

    # Walk upward from target group to find a parent that is in the direct membership set.
    # BFS limited to a reasonable depth to avoid infinite loops in circular nesting.
    $queue   = [System.Collections.Generic.Queue[object]]::new()
    $visited = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $queue.Enqueue([PSCustomObject]@{ GroupId = $TargetGroupId; DisplayName = $TargetGroupDisplayName; Path = $TargetGroupDisplayName })
    [void]$visited.Add($TargetGroupId)

    $maxDepth = 10
    $depth    = 0

    while ($queue.Count -gt 0 -and $depth -lt $maxDepth) {
        $depth++
        $batchSize = $queue.Count
        for ($i = 0; $i -lt $batchSize; $i++) {
            $current = $queue.Dequeue()

            # Fetch parent groups of current group
            if (-not $ParentGroupCache.ContainsKey($current.GroupId)) {
                $parentGroups = @(Invoke-WithRetry -OperationName "Get parent groups of $($current.GroupId)" -ScriptBlock {
                    Get-MgGroupMemberOf -GroupId $current.GroupId -All -Property 'Id,DisplayName' -ErrorAction Stop
                })
                $ParentGroupCache[$current.GroupId] = $parentGroups
            }

            $parents = @($ParentGroupCache[$current.GroupId])
            foreach ($parent in $parents) {
                $parentId   = ([string]$parent.Id).Trim()
                $parentName = ([string]$parent.DisplayName).Trim()
                $builtPath  = "$parentName > $($current.Path)"

                if ($DirectMemberIds.Contains($parentId)) {
                    return "$Upn > $builtPath"
                }

                if (-not $visited.Contains($parentId)) {
                    [void]$visited.Add($parentId)
                    $queue.Enqueue([PSCustomObject]@{ GroupId = $parentId; DisplayName = $parentName; Path = $builtPath })
                }
            }
        }
    }

    # Fallback if path cannot be resolved within depth limit
    return "$Upn > (nested) > $TargetGroupDisplayName"
}

$requiredHeaders = @('UserPrincipalName')

$reportPropertyOrder = @(
    'TimestampUtc', 'RowNumber', 'PrimaryKey', 'Action', 'Status', 'Message', 'ScopeMode',
    'UserPrincipalName', 'GroupId', 'GroupDisplayName', 'GroupMailNickname',
    'GroupType', 'MembershipType', 'InsertionMethod', 'MembershipDepth', 'MembershipPath'
)

Write-Status -Message 'Starting Entra ID assigned security group membership by user inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('GroupMember.Read.All', 'Directory.Read.All')

$scopeMode = 'Csv'
if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. Enumerating all licensed users in the tenant.'

    $licensedUsers = @(Invoke-WithRetry -OperationName 'Enumerate all licensed users' -ScriptBlock {
        Get-MgUser -All -Property 'Id,DisplayName,UserPrincipalName' `
            -Filter 'assignedLicenses/$count ne 0' `
            -ConsistencyLevel eventual `
            -CountVariable licensedUserCount `
            -ErrorAction Stop
    })

    Write-Status -Message "DiscoverAll: found $($licensedUsers.Count) licensed users."
    $rows = @($licensedUsers | ForEach-Object { [PSCustomObject]@{ UserPrincipalName = $_.UserPrincipalName } })
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results          = [System.Collections.Generic.List[object]]::new()
$parentGroupCache = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$rowNumber        = 1

foreach ($row in $rows) {
    $upn = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw 'UserPrincipalName is required and cannot be blank.'
        }

        # Resolve user object
        $userObj = Invoke-WithRetry -OperationName "Resolve user $upn" -ScriptBlock {
            Get-MgUser -UserId $upn -Property 'Id,DisplayName,UserPrincipalName' -ErrorAction Stop
        }
        $userId = ([string]$userObj.Id).Trim()

        # Get all transitive memberships
        $transitiveMemberships = @(Invoke-WithRetry -OperationName "Get transitive memberships for $upn" -ScriptBlock {
            Get-MgUserTransitiveMemberOf -UserId $userId -All `
                -Property 'Id,DisplayName,GroupTypes,SecurityEnabled,MailEnabled,MailNickname,MembershipRule' `
                -ErrorAction Stop
        })

        # Filter to assigned security groups only
        $assignedGroups = @($transitiveMemberships | Where-Object { Test-IsAssignedSecurityGroup -Group $_ })

        if ($assignedGroups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $upn -Action 'GetEntraAssignedGroupsByUser' -Status 'Success' -Message 'No memberships found.' -Data ([ordered]@{
                UserPrincipalName = $upn
                GroupId           = ''
                GroupDisplayName  = ''
                GroupMailNickname = ''
                GroupType         = ''
                MembershipType    = ''
                InsertionMethod   = ''
                MembershipDepth   = ''
                MembershipPath    = ''
            })))
            $rowNumber++
            continue
        }

        # Get direct memberships for path-building
        $directMemberships = @(Invoke-WithRetry -OperationName "Get direct memberships for $upn" -ScriptBlock {
            Get-MgUserMemberOf -UserId $userId -All -Property 'Id' -ErrorAction Stop
        })
        $directMemberIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($dm in $directMemberships) {
            [void]$directMemberIds.Add(([string]$dm.Id).Trim())
        }

        foreach ($group in @($assignedGroups | Sort-Object -Property DisplayName, Id)) {
            $groupId          = ([string]$group.Id).Trim()
            $groupDisplayName = ([string]$group.DisplayName).Trim()
            $groupMailNick    = ([string]$group.MailNickname).Trim()

            $isDirect       = $directMemberIds.Contains($groupId)
            $membershipType = if ($isDirect) { 'Direct' } else { 'Transitive' }
            $depth          = if ($isDirect) { 1 } else { '' }

            $membershipPath = Get-MembershipPath `
                -Upn $upn `
                -TargetGroupId $groupId `
                -TargetGroupDisplayName $groupDisplayName `
                -DirectMemberIds $directMemberIds `
                -ParentGroupCache $parentGroupCache

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$upn|$groupId" -Action 'GetEntraAssignedGroupsByUser' -Status 'Success' -Message 'Group membership exported.' -Data ([ordered]@{
                UserPrincipalName = $upn
                GroupId           = $groupId
                GroupDisplayName  = $groupDisplayName
                GroupMailNickname = $groupMailNick
                GroupType         = 'AssignedSecurity'
                MembershipType    = $membershipType
                InsertionMethod   = 'Direct'
                MembershipDepth   = $depth
                MembershipPath    = $membershipPath
            })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $upn -Action 'GetEntraAssignedGroupsByUser' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            UserPrincipalName = $upn
            GroupId           = ''
            GroupDisplayName  = ''
            GroupMailNickname = ''
            GroupType         = ''
            MembershipType    = ''
            InsertionMethod   = ''
            MembershipDepth   = ''
            MembershipPath    = ''
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
Write-Status -Message 'Entra ID assigned security group membership by user inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
