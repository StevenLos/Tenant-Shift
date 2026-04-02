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

.SYNOPSIS
    Modifies MicrosoftTeamChannelMembers in Microsoft 365.

.DESCRIPTION
    Updates MicrosoftTeamChannelMembers in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M3312-Set-MicrosoftTeamChannelMembers.ps1 -InputCsvPath .\3312.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M3312-Set-MicrosoftTeamChannelMembers.ps1 -InputCsvPath .\3312.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users, Microsoft.Graph.Teams
    Required roles:   Global Administrator or appropriate workload-specific role
    Limitations:      None known.

    CSV Fields:
    Column                Type      Required  Description
    --------------------  ----      --------  -----------
    TeamMailNickname      String    Yes       <fill in description>
    ChannelDisplayName    String    Yes       <fill in description>
    UserPrincipalName     String    Yes       <fill in description>
    Role                  String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3312-Set-MicrosoftTeamChannelMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'UserPrincipalName',
    'Role'
)

Write-Status -Message 'Starting Microsoft Teams channel membership script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users', 'Microsoft.Graph.Teams')
Ensure-GraphConnection -RequiredScopes @('Group.ReadWrite.All', 'TeamMember.ReadWrite.All', 'User.Read.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$teamByAlias = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$channelMapByTeamId = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$channelMembersByKey = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$userByUpn = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

function ConvertTo-CanonicalRoles {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$Roles
    )

    if ($null -eq $Roles) {
        return ,@()
    }

    return ,@(
        @($Roles) |
            ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
}

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
        [string]$TeamId
    )

    if ($channelMapByTeamId.ContainsKey($TeamId)) {
        return [System.Collections.Generic.Dictionary[string, object]]$channelMapByTeamId[$TeamId]
    }

    $map = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $channels = @(Invoke-WithRetry -OperationName "Load Team channels for $TeamId" -ScriptBlock {
        Get-MgTeamChannel -TeamId $TeamId -All -ErrorAction Stop
    })

    foreach ($channel in $channels) {
        $name = ([string]$channel.DisplayName).Trim()
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        if (-not $map.ContainsKey($name)) {
            $map[$name] = $channel
        }
    }

    $channelMapByTeamId[$TeamId] = $map
    return $map
}

function Get-ChannelMembersMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TeamId,

        [Parameter(Mandatory)]
        [string]$ChannelId
    )

    $key = "$TeamId|$ChannelId"
    if ($channelMembersByKey.ContainsKey($key)) {
        return [System.Collections.Generic.Dictionary[string, object]]$channelMembersByKey[$key]
    }

    $map = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $members = @(Invoke-WithRetry -OperationName "Load channel members $key" -ScriptBlock {
        Get-MgTeamChannelMember -TeamId $TeamId -ChannelId $ChannelId -All -ErrorAction Stop
    })

    foreach ($member in $members) {
        $memberId = ([string]$member.Id).Trim()
        $roles = @()
        if ($member.PSObject.Properties.Name -contains 'Roles') {
            $roles = @($member.Roles)
        }

        $additional = $null
        if ($member.PSObject.Properties.Name -contains 'AdditionalProperties') {
            $additional = $member.AdditionalProperties
        }

        $memberUserId = ''
        if ($member.PSObject.Properties.Name -contains 'UserId') {
            $memberUserId = ([string]$member.UserId).Trim()
        }
        if ([string]::IsNullOrWhiteSpace($memberUserId) -and $additional) {
            try {
                if ($additional.ContainsKey('userId')) {
                    $memberUserId = ([string]$additional['userId']).Trim()
                }
            }
            catch {
                # Best effort.
            }
        }

        $memberEmail = ''
        if ($member.PSObject.Properties.Name -contains 'Email') {
            $memberEmail = ([string]$member.Email).Trim().ToLowerInvariant()
        }
        if ([string]::IsNullOrWhiteSpace($memberEmail) -and $additional) {
            try {
                if ($additional.ContainsKey('email')) {
                    $memberEmail = ([string]$additional['email']).Trim().ToLowerInvariant()
                }
            }
            catch {
                # Best effort.
            }
        }

        $entry = [PSCustomObject]@{
            ConversationMemberId = $memberId
            RolesCanonical       = ConvertTo-CanonicalRoles -Roles $roles
            MemberUserId         = $memberUserId
            MemberEmail          = $memberEmail
        }

        if (-not [string]::IsNullOrWhiteSpace($memberUserId)) {
            $map["id:$memberUserId"] = $entry
        }

        if (-not [string]::IsNullOrWhiteSpace($memberEmail)) {
            $map["email:$memberEmail"] = $entry
        }
    }

    $channelMembersByKey[$key] = $map
    return $map
}

$rowNumber = 1
foreach ($row in $rows) {
    $teamMailNickname = ([string]$row.TeamMailNickname).Trim()
    $channelDisplayName = ([string]$row.ChannelDisplayName).Trim()
    $upn = ([string]$row.UserPrincipalName).Trim()
    $roleRaw = ([string]$row.Role).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($teamMailNickname) -or [string]::IsNullOrWhiteSpace($channelDisplayName) -or [string]::IsNullOrWhiteSpace($upn)) {
            throw 'TeamMailNickname, ChannelDisplayName, and UserPrincipalName are required.'
        }

        $role = if ([string]::IsNullOrWhiteSpace($roleRaw)) { 'Member' } else { $roleRaw }
        if ($role -notin @('Member', 'Owner')) {
            throw "Role '$role' is invalid. Use Member or Owner."
        }

        $desiredRoles = if ($role -eq 'Owner') { @('owner') } else { @() }
        $desiredRolesCanonical = ConvertTo-CanonicalRoles -Roles $desiredRoles

        $teamGroup = Resolve-TeamByAlias -TeamMailNickname $teamMailNickname
        $teamId = ([string]$teamGroup.Id).Trim()
        if ([string]::IsNullOrWhiteSpace($teamId)) {
            throw "Team '$teamMailNickname' does not have a valid Team/group ID."
        }

        $channelMap = Get-TeamChannelMap -TeamId $teamId
        if (-not $channelMap.ContainsKey($channelDisplayName)) {
            throw "Channel '$channelDisplayName' was not found in Team '$teamMailNickname'."
        }

        $channel = $channelMap[$channelDisplayName]
        $channelId = ([string]$channel.Id).Trim()
        if ([string]::IsNullOrWhiteSpace($channelId)) {
            throw "Channel '$channelDisplayName' does not have a valid channel ID."
        }

        $membershipType = ([string]$channel.MembershipType).Trim()
        if ([string]::IsNullOrWhiteSpace($membershipType)) {
            $membershipType = 'Standard'
        }

        if ($membershipType.Equals('Standard', [System.StringComparison]::OrdinalIgnoreCase)) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$channelDisplayName|$upn" -Action 'SetTeamChannelMembership' -Status 'Skipped' -Message 'Standard channels inherit Team membership. Use SM-M3310-Set-MicrosoftTeamMembers.ps1 to manage access.'))
            $rowNumber++
            continue
        }

        $user = Resolve-UserByUpn -UserPrincipalName $upn
        $userId = ([string]$user.Id).Trim()
        if ([string]::IsNullOrWhiteSpace($userId)) {
            throw "User '$upn' does not have a valid object ID."
        }

        $channelMemberMap = Get-ChannelMembersMap -TeamId $teamId -ChannelId $channelId
        $existingEntry = $null
        if ($channelMemberMap.ContainsKey("id:$userId")) {
            $existingEntry = $channelMemberMap["id:$userId"]
        }
        else {
            $upnLower = $upn.ToLowerInvariant()
            if ($channelMemberMap.ContainsKey("email:$upnLower")) {
                $existingEntry = $channelMemberMap["email:$upnLower"]
            }
        }

        if ($existingEntry) {
            $currentRoles = @($existingEntry.RolesCanonical)
            $leftOnly = @($currentRoles | Where-Object { $_ -notin $desiredRolesCanonical })
            $rightOnly = @($desiredRolesCanonical | Where-Object { $_ -notin $currentRoles })
            $roleMatches = ($leftOnly.Count -eq 0 -and $rightOnly.Count -eq 0)

            if ($roleMatches) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$channelDisplayName|$upn" -Action 'SetTeamChannelMembership' -Status 'Skipped' -Message "User is already in channel with role '$role'."))
                $rowNumber++
                continue
            }

            $conversationMemberId = ([string]$existingEntry.ConversationMemberId).Trim()
            if ([string]::IsNullOrWhiteSpace($conversationMemberId)) {
                throw "Unable to update role for '$upn' because the existing channel member ID was not available."
            }

            if ($PSCmdlet.ShouldProcess("$teamMailNickname/$channelDisplayName -> $upn", "Update channel role to $role")) {
                Invoke-WithRetry -OperationName "Update channel role $teamMailNickname/$channelDisplayName -> $upn" -ScriptBlock {
                    Update-MgTeamChannelMember -TeamId $teamId -ChannelId $channelId -ConversationMemberId $conversationMemberId -Roles $desiredRoles -ErrorAction Stop | Out-Null
                }

                $existingEntry.RolesCanonical = $desiredRolesCanonical
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$channelDisplayName|$upn" -Action 'SetTeamChannelMembership' -Status 'Updated' -Message "User role updated to '$role'."))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$channelDisplayName|$upn" -Action 'SetTeamChannelMembership' -Status 'WhatIf' -Message 'Role update skipped due to WhatIf.'))
            }
        }
        else {
            if ($PSCmdlet.ShouldProcess("$teamMailNickname/$channelDisplayName -> $upn", "Add channel user as $role")) {
                $memberBody = @{
                    '@odata.type'     = '#microsoft.graph.aadUserConversationMember'
                    roles             = $desiredRoles
                    'user@odata.bind' = "https://graph.microsoft.com/v1.0/users('$userId')"
                }

                $createdMember = Invoke-WithRetry -OperationName "Add channel user $teamMailNickname/$channelDisplayName -> $upn" -ScriptBlock {
                    New-MgTeamChannelMember -TeamId $teamId -ChannelId $channelId -BodyParameter $memberBody -ErrorAction Stop
                }

                $memberId = ''
                if ($createdMember -and $createdMember.PSObject.Properties.Name -contains 'Id') {
                    $memberId = ([string]$createdMember.Id).Trim()
                }

                $entry = [PSCustomObject]@{
                    ConversationMemberId = $memberId
                    RolesCanonical       = $desiredRolesCanonical
                    MemberUserId         = $userId
                    MemberEmail          = $upn.ToLowerInvariant()
                }
                $channelMemberMap["id:$userId"] = $entry
                $channelMemberMap["email:$($upn.ToLowerInvariant())"] = $entry

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$channelDisplayName|$upn" -Action 'SetTeamChannelMembership' -Status 'Added' -Message "User added to channel with role '$role'."))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$channelDisplayName|$upn" -Action 'SetTeamChannelMembership' -Status 'WhatIf' -Message 'Channel membership update skipped due to WhatIf.'))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($teamMailNickname|$channelDisplayName|$upn|$roleRaw) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$channelDisplayName|$upn|$roleRaw" -Action 'SetTeamChannelMembership' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Microsoft Teams channel membership script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







