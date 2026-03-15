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
Microsoft.Graph.Authentication
Microsoft.Graph.Groups
Microsoft.Graph.Users
Microsoft.Graph.Teams

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3310-Set-MicrosoftTeamMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {


$requiredHeaders = @(
    'TeamMailNickname',
    'UserPrincipalName',
    'Role'
)

Write-Status -Message 'Starting Microsoft Teams user membership script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users', 'Microsoft.Graph.Teams')
Ensure-GraphConnection -RequiredScopes @('Group.ReadWrite.All', 'User.Read.All', 'Team.ReadBasic.All', 'TeamMember.ReadWrite.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$teamByAlias = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$memberIdsByGroupId = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$ownerIdsByGroupId = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$userByUpn = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

function Resolve-TeamByAlias {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TeamMailNickname
    )

    if ($teamByAlias.ContainsKey($TeamMailNickname)) {
        return $teamByAlias[$TeamMailNickname]
    }

    $escapedAlias = Escape-ODataString -Value $TeamMailNickname
    $groups = @(Invoke-WithRetry -OperationName "Lookup Team group alias $TeamMailNickname" -ScriptBlock {
        Get-MgGroup -Filter "mailNickname eq '$escapedAlias'" -ConsistencyLevel eventual -ErrorAction Stop
    })

    if ($groups.Count -eq 0) {
        throw "No group was found with mailNickname '$TeamMailNickname'."
    }

    if ($groups.Count -gt 1) {
        throw "Multiple groups were found with mailNickname '$TeamMailNickname'. Resolve duplicate aliases before running this script."
    }

    $group = Invoke-WithRetry -OperationName "Load Team group details for alias $TeamMailNickname" -ScriptBlock {
        Get-MgGroup -GroupId $groups[0].Id -Property 'id,displayName,mailNickname,groupTypes,securityEnabled,mailEnabled' -ErrorAction Stop
    }

    $groupTypes = @($group.GroupTypes)
    $isMicrosoft365Group = ($groupTypes -contains 'Unified') -and ($group.MailEnabled -eq $true) -and ($group.SecurityEnabled -eq $false)
    if (-not $isMicrosoft365Group) {
        throw "Group '$TeamMailNickname' exists but is not a Microsoft 365 group."
    }

    $team = Invoke-WithRetry -OperationName "Verify Team exists for alias $TeamMailNickname" -ScriptBlock {
        Get-MgGroupTeam -GroupId $group.Id -ErrorAction SilentlyContinue
    }
    if (-not $team) {
        throw "Microsoft 365 group '$TeamMailNickname' does not currently have a Team."
    }

    $teamByAlias[$TeamMailNickname] = $group
    return $group
}

function Resolve-UserByUpn {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    if ($userByUpn.ContainsKey($UserPrincipalName)) {
        return $userByUpn[$UserPrincipalName]
    }

    $escapedUpn = Escape-ODataString -Value $UserPrincipalName
    $users = @(Invoke-WithRetry -OperationName "Lookup user $UserPrincipalName" -ScriptBlock {
        Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -ErrorAction Stop
    })

    if ($users.Count -eq 0) {
        throw "User '$UserPrincipalName' was not found."
    }

    if ($users.Count -gt 1) {
        throw "Multiple users were returned for UPN '$UserPrincipalName'. Resolve duplicate directory objects before retrying."
    }

    $user = $users[0]
    $userByUpn[$UserPrincipalName] = $user
    return $user
}

function Get-TeamMemberIdSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$GroupId
    )

    if ($memberIdsByGroupId.ContainsKey($GroupId)) {
        return [System.Collections.Generic.HashSet[string]]$memberIdsByGroupId[$GroupId]
    }

    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $members = @(Invoke-WithRetry -OperationName "Load Team members for group $GroupId" -ScriptBlock {
        Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop
    })

    foreach ($member in $members) {
        $memberId = ([string]$member.Id).Trim()
        if (-not [string]::IsNullOrWhiteSpace($memberId)) {
            $null = $set.Add($memberId)
        }
    }

    $memberIdsByGroupId[$GroupId] = $set
    return $set
}

function Get-TeamOwnerIdSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$GroupId
    )

    if ($ownerIdsByGroupId.ContainsKey($GroupId)) {
        return [System.Collections.Generic.HashSet[string]]$ownerIdsByGroupId[$GroupId]
    }

    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $owners = @(Invoke-WithRetry -OperationName "Load Team owners for group $GroupId" -ScriptBlock {
        Get-MgGroupOwner -GroupId $GroupId -All -ErrorAction Stop
    })

    foreach ($owner in $owners) {
        $ownerId = ([string]$owner.Id).Trim()
        if (-not [string]::IsNullOrWhiteSpace($ownerId)) {
            $null = $set.Add($ownerId)
        }
    }

    $ownerIdsByGroupId[$GroupId] = $set
    return $set
}

$rowNumber = 1
foreach ($row in $rows) {
    $teamMailNickname = ([string]$row.TeamMailNickname).Trim()
    $upn = ([string]$row.UserPrincipalName).Trim()
    $roleRaw = ([string]$row.Role).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($teamMailNickname) -or [string]::IsNullOrWhiteSpace($upn)) {
            throw 'TeamMailNickname and UserPrincipalName are required.'
        }

        $role = if ([string]::IsNullOrWhiteSpace($roleRaw)) { 'Member' } else { $roleRaw }
        if ($role -notin @('Member', 'Owner')) {
            throw "Role '$role' is invalid. Use Member or Owner."
        }

        $teamGroup = Resolve-TeamByAlias -TeamMailNickname $teamMailNickname
        $user = Resolve-UserByUpn -UserPrincipalName $upn

        $groupId = ([string]$teamGroup.Id).Trim()
        $userId = ([string]$user.Id).Trim()
        if ([string]::IsNullOrWhiteSpace($groupId) -or [string]::IsNullOrWhiteSpace($userId)) {
            throw 'Unable to resolve group or user object ID.'
        }

        if ($role -eq 'Member') {
            $memberIds = Get-TeamMemberIdSet -GroupId $groupId
            if ($memberIds.Contains($userId)) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$upn|$role" -Action 'AddUserToTeam' -Status 'Skipped' -Message 'User is already a Team member.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$teamMailNickname -> $upn", 'Add Team member')) {
                $memberRef = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$userId" }
                Invoke-WithRetry -OperationName "Add Team member $upn to $teamMailNickname" -ScriptBlock {
                    New-MgGroupMemberByRef -GroupId $groupId -BodyParameter $memberRef -ErrorAction Stop
                }
                $null = $memberIds.Add($userId)
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$upn|$role" -Action 'AddUserToTeam' -Status 'Added' -Message 'User added to Team as member.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$upn|$role" -Action 'AddUserToTeam' -Status 'WhatIf' -Message 'Membership update skipped due to WhatIf.'))
            }
        }
        else {
            $ownerIds = Get-TeamOwnerIdSet -GroupId $groupId
            if ($ownerIds.Contains($userId)) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$upn|$role" -Action 'AddUserToTeam' -Status 'Skipped' -Message 'User is already a Team owner.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$teamMailNickname -> $upn", 'Add Team owner')) {
                $ownerRef = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$userId" }
                Invoke-WithRetry -OperationName "Add Team owner $upn to $teamMailNickname" -ScriptBlock {
                    New-MgGroupOwnerByRef -GroupId $groupId -BodyParameter $ownerRef -ErrorAction Stop
                }
                $null = $ownerIds.Add($userId)
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$upn|$role" -Action 'AddUserToTeam' -Status 'Added' -Message 'User added to Team as owner.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$upn|$role" -Action 'AddUserToTeam' -Status 'WhatIf' -Message 'Owner update skipped due to WhatIf.'))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($teamMailNickname|$upn|$roleRaw) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$upn|$roleRaw" -Action 'AddUserToTeam' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Microsoft Teams user membership script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







