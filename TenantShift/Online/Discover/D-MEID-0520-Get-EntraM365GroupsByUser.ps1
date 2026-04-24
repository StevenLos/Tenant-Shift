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
    Gets EntraM365GroupsByUser and exports results to CSV.

.DESCRIPTION
    Gets EntraM365GroupsByUser from Microsoft 365 and writes the results to a CSV file.
    For each in-scope user, finds all Microsoft 365 groups (Unified groups) where the user
    is a member or owner. Captures both relationships with a Relationship column
    (Member / Owner / Member+Owner).
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
    .\D-MEID-0520-Get-EntraM365GroupsByUser.ps1 -InputCsvPath .\Scope-Users.input.csv
    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\D-MEID-0520-Get-EntraM365GroupsByUser.ps1 -DiscoverAll
    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users
    Required roles:   Global Reader
    Limitations:      None known.

    CSV Fields:
    Column              Type    Required  Description
    ------              ----    --------  -----------
    UserPrincipalName   String  Yes       UPN of the user whose M365 group memberships to inventory
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-MEID-0520-Get-EntraM365GroupsByUser_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Test-IsM365Group {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Group
    )

    $groupTypes = @()
    if ($null -ne $Group.GroupTypes) {
        $groupTypes = @($Group.GroupTypes)
    }

    return ($groupTypes -contains 'Unified')
}

function Get-ODataType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$DirectoryObject
    )

    if ($DirectoryObject.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $additional = $DirectoryObject.AdditionalProperties
        if ($additional) {
            try {
                if ($additional.ContainsKey('@odata.type')) {
                    return ([string]$additional['@odata.type']).Trim()
                }
            }
            catch { }
        }
    }
    return ''
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

    # Direct member path is trivial
    if ($DirectMemberIds.Contains($TargetGroupId)) {
        return "$Upn > $TargetGroupDisplayName"
    }

    # BFS upward from the target group to find a direct member ancestor
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
    'UserPrincipalName', 'GroupId', 'GroupDisplayName', 'GroupMailNickname', 'GroupMail',
    'Relationship', 'MembershipType', 'InsertionMethod', 'MembershipDepth', 'MembershipPath'
)

Write-Status -Message 'Starting Entra ID Microsoft 365 group membership by user inventory script.'
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
                -Property 'Id,DisplayName,GroupTypes,SecurityEnabled,MailEnabled,MailNickname,Mail' `
                -ErrorAction Stop
        })

        # Filter to M365 (Unified) groups only
        $m365Groups = @($transitiveMemberships | Where-Object { Test-IsM365Group -Group $_ })

        # Build a set of M365 group IDs from transitive membership
        $m365GroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($g in $m365Groups) { [void]$m365GroupIds.Add(([string]$g.Id).Trim()) }

        # Get direct memberships for path-building
        $directMemberships = @(Invoke-WithRetry -OperationName "Get direct memberships for $upn" -ScriptBlock {
            Get-MgUserMemberOf -UserId $userId -All -Property 'Id' -ErrorAction Stop
        })
        $directMemberIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($dm in $directMemberships) {
            [void]$directMemberIds.Add(([string]$dm.Id).Trim())
        }

        # Get owned objects and filter to groups
        $ownedObjects = @(Invoke-WithRetry -OperationName "Get owned objects for $upn" -ScriptBlock {
            Get-MgUserOwnedObject -UserId $userId -All -Property 'Id,DisplayName' -ErrorAction Stop
        })
        $ownedGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($obj in $ownedObjects) {
            $odataType = Get-ODataType -DirectoryObject $obj
            if ($odataType -eq '#microsoft.graph.group') {
                [void]$ownedGroupIds.Add(([string]$obj.Id).Trim())
            }
        }

        # Build a unified set of all M365 group IDs (member or owner)
        $allGroupIds = [System.Collections.Generic.HashSet[string]]::new($m365GroupIds, [System.StringComparer]::OrdinalIgnoreCase)
        foreach ($oid in $ownedGroupIds) { [void]$allGroupIds.Add($oid) }

        if ($allGroupIds.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $upn -Action 'GetEntraM365GroupsByUser' -Status 'Success' -Message 'No memberships found.' -Data ([ordered]@{
                UserPrincipalName = $upn
                GroupId           = ''
                GroupDisplayName  = ''
                GroupMailNickname = ''
                GroupMail         = ''
                Relationship      = ''
                MembershipType    = ''
                InsertionMethod   = ''
                MembershipDepth   = ''
                MembershipPath    = ''
            })))
            $rowNumber++
            continue
        }

        # Build lookup of group objects from transitive membership results
        $groupObjectById = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($g in $m365Groups) {
            $gid = ([string]$g.Id).Trim()
            if (-not $groupObjectById.ContainsKey($gid)) {
                $groupObjectById[$gid] = $g
            }
        }

        # Process each unique M365 group ID
        foreach ($groupId in ($allGroupIds | Sort-Object)) {
            $isMember = $m365GroupIds.Contains($groupId)
            $isOwner  = $ownedGroupIds.Contains($groupId)

            $relationship = if ($isMember -and $isOwner) { 'Member+Owner' }
                            elseif ($isMember)            { 'Member' }
                            else                          { 'Owner' }

            # Get group details — prefer cached transitive object, fall back to Graph call
            $groupObj = $null
            if ($groupObjectById.ContainsKey($groupId)) {
                $groupObj = $groupObjectById[$groupId]
            }
            else {
                # Owner-only group: fetch details directly
                $groupObj = Invoke-WithRetry -OperationName "Get group details for $groupId" -ScriptBlock {
                    Get-MgGroup -GroupId $groupId -Property 'Id,DisplayName,MailNickname,Mail,GroupTypes' -ErrorAction Stop
                }
            }

            $groupDisplayName = ([string]$groupObj.DisplayName).Trim()
            $groupMailNick    = ([string]$groupObj.MailNickname).Trim()
            $groupMail        = ([string]$groupObj.Mail).Trim()

            # Determine membership type and path
            if ($isOwner -and -not $isMember) {
                # Ownership is always direct
                $membershipType = 'Direct'
                $insertionMethod = 'Direct'
                $depth           = 1
                $membershipPath  = "$upn > $groupDisplayName"
            }
            elseif ($directMemberIds.Contains($groupId)) {
                $membershipType  = 'Direct'
                $insertionMethod = 'Direct'
                $depth           = 1
                $membershipPath  = "$upn > $groupDisplayName"
            }
            else {
                $membershipType  = 'Transitive'
                $insertionMethod = 'Direct'
                $depth           = ''
                $membershipPath  = Get-MembershipPath `
                    -Upn $upn `
                    -TargetGroupId $groupId `
                    -TargetGroupDisplayName $groupDisplayName `
                    -DirectMemberIds $directMemberIds `
                    -ParentGroupCache $parentGroupCache
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$upn|$groupId|$relationship" -Action 'GetEntraM365GroupsByUser' -Status 'Success' -Message 'M365 group membership exported.' -Data ([ordered]@{
                UserPrincipalName = $upn
                GroupId           = $groupId
                GroupDisplayName  = $groupDisplayName
                GroupMailNickname = $groupMailNick
                GroupMail         = $groupMail
                Relationship      = $relationship
                MembershipType    = $membershipType
                InsertionMethod   = $insertionMethod
                MembershipDepth   = $depth
                MembershipPath    = $membershipPath
            })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $upn -Action 'GetEntraM365GroupsByUser' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            UserPrincipalName = $upn
            GroupId           = ''
            GroupDisplayName  = ''
            GroupMailNickname = ''
            GroupMail         = ''
            Relationship      = ''
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
Write-Status -Message 'Entra ID Microsoft 365 group membership by user inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
