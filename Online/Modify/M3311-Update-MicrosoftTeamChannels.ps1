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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M3311-Update-MicrosoftTeamChannels_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'ChannelDisplayName',
    'Description',
    'MembershipType',
    'OwnerUserPrincipalNames',
    'IsFavoriteByDefault'
)

Write-Status -Message 'Starting Microsoft Teams channel creation script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users', 'Microsoft.Graph.Teams')
Ensure-GraphConnection -RequiredScopes @('Group.ReadWrite.All', 'Team.ReadBasic.All', 'User.Read.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$teamByAlias = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$channelsByGroupId = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
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

function Get-TeamChannelMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$GroupId
    )

    if ($channelsByGroupId.ContainsKey($GroupId)) {
        return [System.Collections.Generic.Dictionary[string, object]]$channelsByGroupId[$GroupId]
    }

    $channelMap = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $channels = @(Invoke-WithRetry -OperationName "Load Team channels for group $GroupId" -ScriptBlock {
        Get-MgTeamChannel -TeamId $GroupId -All -ErrorAction Stop
    })

    foreach ($channel in $channels) {
        $channelName = ([string]$channel.DisplayName).Trim()
        if ([string]::IsNullOrWhiteSpace($channelName)) {
            continue
        }

        if (-not $channelMap.ContainsKey($channelName)) {
            $channelMap[$channelName] = $channel
        }
    }

    $channelsByGroupId[$GroupId] = $channelMap
    return $channelMap
}

$rowNumber = 1
foreach ($row in $rows) {
    $teamMailNickname = ([string]$row.TeamMailNickname).Trim()
    $channelDisplayName = ([string]$row.ChannelDisplayName).Trim()
    $membershipTypeRaw = ([string]$row.MembershipType).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($teamMailNickname) -or [string]::IsNullOrWhiteSpace($channelDisplayName)) {
            throw 'TeamMailNickname and ChannelDisplayName are required.'
        }

        $description = ([string]$row.Description).Trim()
        $membershipType = if ([string]::IsNullOrWhiteSpace($membershipTypeRaw)) { 'Standard' } else { $membershipTypeRaw }
        if ($membershipType -notin @('Standard', 'Private', 'Shared')) {
            throw "MembershipType '$membershipType' is invalid. Use Standard, Private, or Shared."
        }

        $ownerUpns = ConvertTo-Array -Value ([string]$row.OwnerUserPrincipalNames)
        if ($membershipType -in @('Private', 'Shared') -and $ownerUpns.Count -eq 0) {
            throw "MembershipType '$membershipType' requires at least one owner in OwnerUserPrincipalNames."
        }

        $isFavoriteRaw = ([string]$row.IsFavoriteByDefault).Trim()
        $setIsFavorite = -not [string]::IsNullOrWhiteSpace($isFavoriteRaw)
        $isFavoriteByDefault = $false
        if ($setIsFavorite) {
            $isFavoriteByDefault = ConvertTo-Bool -Value $isFavoriteRaw
        }

        $teamGroup = Resolve-TeamByAlias -TeamMailNickname $teamMailNickname
        $groupId = ([string]$teamGroup.Id).Trim()
        if ([string]::IsNullOrWhiteSpace($groupId)) {
            throw "Team '$teamMailNickname' does not have a valid group ID."
        }

        $channelMap = Get-TeamChannelMap -GroupId $groupId
        if ($channelMap.ContainsKey($channelDisplayName)) {
            $existingChannel = $channelMap[$channelDisplayName]
            $existingMembershipType = ([string]$existingChannel.MembershipType).Trim()
            if ([string]::IsNullOrWhiteSpace($existingMembershipType)) {
                $existingMembershipType = 'Standard'
            }

            if ($existingMembershipType.Equals($membershipType, [System.StringComparison]::OrdinalIgnoreCase)) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$channelDisplayName" -Action 'AddTeamChannel' -Status 'Skipped' -Message 'Channel already exists with the requested membership type.'))
                $rowNumber++
                continue
            }

            throw "Channel '$channelDisplayName' already exists with membership type '$existingMembershipType', not '$membershipType'."
        }

        $channelBody = @{
            displayName    = $channelDisplayName
            membershipType = $membershipType.ToLowerInvariant()
        }

        if (-not [string]::IsNullOrWhiteSpace($description)) {
            $channelBody.description = $description
        }

        if ($setIsFavorite) {
            $channelBody.isFavoriteByDefault = $isFavoriteByDefault
        }

        if ($membershipType -in @('Private', 'Shared')) {
            $ownerMemberEntries = [System.Collections.Generic.List[hashtable]]::new()
            $seenOwnerUserIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

            foreach ($ownerUpn in $ownerUpns) {
                $ownerUser = Resolve-UserByUpn -UserPrincipalName $ownerUpn
                $ownerUserId = ([string]$ownerUser.Id).Trim()
                if ([string]::IsNullOrWhiteSpace($ownerUserId)) {
                    throw "Owner '$ownerUpn' does not have a valid object ID."
                }

                if ($seenOwnerUserIds.Add($ownerUserId)) {
                    $ownerMemberEntries.Add(@{
                        '@odata.type'     = '#microsoft.graph.aadUserConversationMember'
                        roles             = @('owner')
                        'user@odata.bind' = "https://graph.microsoft.com/v1.0/users('$ownerUserId')"
                    })
                }
            }

            if ($ownerMemberEntries.Count -eq 0) {
                throw "MembershipType '$membershipType' requires at least one unique owner entry."
            }

            $channelBody.members = $ownerMemberEntries.ToArray()
        }

        if ($PSCmdlet.ShouldProcess("$teamMailNickname -> $channelDisplayName", 'Create Team channel')) {
            $createdChannel = Invoke-WithRetry -OperationName "Create Team channel $teamMailNickname/$channelDisplayName" -ScriptBlock {
                New-MgTeamChannel -TeamId $groupId -BodyParameter $channelBody -ErrorAction Stop
            }

            if (-not $channelMap.ContainsKey($channelDisplayName)) {
                $channelMap[$channelDisplayName] = $createdChannel
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$channelDisplayName" -Action 'AddTeamChannel' -Status 'Added' -Message "Channel created successfully as '$membershipType'."))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$channelDisplayName" -Action 'AddTeamChannel' -Status 'WhatIf' -Message 'Channel creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($teamMailNickname|$channelDisplayName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$channelDisplayName" -Action 'AddTeamChannel' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Microsoft Teams channel creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







