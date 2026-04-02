<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-194500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Groups
Microsoft.Graph.Teams

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets MicrosoftTeamChannelMembers and exports results to CSV.

.DESCRIPTION
    Gets MicrosoftTeamChannelMembers from Microsoft 365 and writes the results to a CSV file.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all objects in scope (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all objects in scope rather than processing from an input CSV file. Uses the DiscoverAll parameter set.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-IR3312-Get-MicrosoftTeamChannelMembers.ps1 -InputCsvPath .\3312.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3312-Get-MicrosoftTeamChannelMembers.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Teams
    Required roles:   Global Administrator or appropriate workload-specific role
    Limitations:      None known.

    CSV Fields:
    Column                Type      Required  Description
    --------------------  ----      --------  -----------
    TeamMailNickname      String    Yes       <fill in description>
    ChannelDisplayName    String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-TEAM-0040-Get-MicrosoftTeamChannelMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
        [Parameter(Mandatory)]
        [int]$RowNumber,

        [Parameter(Mandatory)]
        [string]$PrimaryKey,

        [Parameter(Mandatory)]
        [string]$Action,

        [Parameter(Mandatory)]
        [string]$Status,

        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory)]
        [hashtable]$Data
    )

    $base = New-ResultObject -RowNumber $RowNumber -PrimaryKey $PrimaryKey -Action $Action -Status $Status -Message $Message
    $ordered = [ordered]@{}

    foreach ($prop in $base.PSObject.Properties.Name) {
        $ordered[$prop] = $base.$prop
    }

    foreach ($key in $Data.Keys) {
        $ordered[$key] = $Data[$key]
    }

    return [PSCustomObject]$ordered
}

function Test-IsMicrosoft365Group {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Group
    )

    $groupTypes = @($Group.GroupTypes)
    $isM365 = ($groupTypes -contains 'Unified') -and ($Group.MailEnabled -eq $true) -and ($Group.SecurityEnabled -eq $false)
    return $isM365
}

function Get-GraphPropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Object,

        [Parameter(Mandatory)]
        [string[]]$PropertyNames
    )

    if ($null -eq $Object) {
        return ''
    }

    foreach ($name in $PropertyNames) {
        if ($Object.PSObject.Properties.Name -contains $name) {
            return [string]$Object.$name
        }
    }

    if ($Object.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $additional = $Object.AdditionalProperties
        if ($additional) {
            foreach ($name in $PropertyNames) {
                try {
                    if ($additional.ContainsKey($name)) {
                        return [string]$additional[$name]
                    }
                }
                catch {
                    # Best effort only.
                }
            }
        }
    }

    return ''
}

function Get-ChannelMemberRoles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Member
    )

    $roles = @()

    if ($Member.PSObject.Properties.Name -contains 'Roles') {
        $roles = @($Member.Roles)
    }

    if ($roles.Count -eq 0 -and $Member.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $additional = $Member.AdditionalProperties
        if ($additional) {
            try {
                if ($additional.ContainsKey('roles')) {
                    $roles = @($additional['roles'])
                }
            }
            catch {
                # Best effort only.
            }
        }
    }

    $normalized = @(
        $roles |
            ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )

    $role = if ($normalized -contains 'owner') { 'Owner' } else { 'Member' }

    return [PSCustomObject]@{
        Role     = $role
        RolesRaw = ($normalized -join ';')
    }
}

$requiredHeaders = @(
    'TeamMailNickname',
    'ChannelDisplayName'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'TeamGroupId',
    'TeamDisplayName',
    'TeamMailNickname',
    'ChannelId',
    'ChannelDisplayName',
    'ChannelMembershipType',
    'ChannelEmail',
    'ChannelWebUrl',
    'RequestedChannelDisplayName',
    'MemberConversationId',
    'MemberObjectId',
    'MemberEmail',
    'MemberDisplayName',
    'MemberODataType',
    'Role',
    'RolesRaw'
)

Write-Status -Message 'Starting Microsoft Teams channel member inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Teams')
Ensure-GraphConnection -RequiredScopes @('Group.Read.All', 'Team.ReadBasic.All', 'TeamSettings.Read.All', 'TeamMember.Read.All', 'ChannelMember.Read.All', 'Directory.Read.All')

$scopeMode = 'Csv'
if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. CSV input is bypassed.' -Level WARN

    $discoverRow = [ordered]@{}
    foreach ($header in $requiredHeaders) {
        $discoverRow[$header] = '*'
    }

    $rows = @([PSCustomObject]$discoverRow)
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}
$results = [System.Collections.Generic.List[object]]::new()

$groupSelect = 'id,displayName,mailNickname,groupTypes,mailEnabled,securityEnabled'
$allM365GroupsCache = $null
$teamByGroupId = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$channelsByGroupId = [System.Collections.Generic.Dictionary[string, object[]]]::new([System.StringComparer]::OrdinalIgnoreCase)

$rowNumber = 1
foreach ($row in $rows) {
    $teamMailNickname = ([string]$row.TeamMailNickname).Trim()
    if ([string]::IsNullOrWhiteSpace($teamMailNickname) -and $row.PSObject.Properties.Name -contains 'GroupMailNickname') {
        $teamMailNickname = ([string]$row.GroupMailNickname).Trim()
    }

    $requestedChannelDisplayName = ([string]$row.ChannelDisplayName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($teamMailNickname)) {
            throw 'TeamMailNickname is required. Use * to inventory channel members for all Teams.'
        }

        if ([string]::IsNullOrWhiteSpace($requestedChannelDisplayName)) {
            $requestedChannelDisplayName = '*'
        }

        $candidateGroups = @()
        if ($teamMailNickname -eq '*') {
            if ($null -eq $allM365GroupsCache) {
                $allGroups = @(Invoke-WithRetry -OperationName 'Load all groups for Team channel member inventory' -ScriptBlock {
                    Get-MgGroup -All -Property $groupSelect -ErrorAction Stop
                })
                $allM365GroupsCache = @($allGroups | Where-Object { Test-IsMicrosoft365Group -Group $_ } | Sort-Object -Property MailNickname, DisplayName, Id)
            }

            $candidateGroups = @($allM365GroupsCache)
        }
        else {
            $escapedAlias = Escape-ODataString -Value $teamMailNickname
            $groups = @(Invoke-WithRetry -OperationName "Lookup group alias $teamMailNickname" -ScriptBlock {
                Get-MgGroup -Filter "mailNickname eq '$escapedAlias'" -ConsistencyLevel eventual -Property $groupSelect -ErrorAction Stop
            })
            $candidateGroups = @($groups | Where-Object { Test-IsMicrosoft365Group -Group $_ })
        }

        if ($candidateGroups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$requestedChannelDisplayName" -Action 'GetMicrosoftTeamChannelMember' -Status 'NotFound' -Message 'No matching Microsoft 365 groups were found.' -Data ([ordered]@{
                        TeamGroupId                  = ''
                        TeamDisplayName              = ''
                        TeamMailNickname             = $teamMailNickname
                        ChannelId                    = ''
                        ChannelDisplayName           = ''
                        ChannelMembershipType        = ''
                        ChannelEmail                 = ''
                        ChannelWebUrl                = ''
                        RequestedChannelDisplayName  = $requestedChannelDisplayName
                        MemberConversationId         = ''
                        MemberObjectId               = ''
                        MemberEmail                  = ''
                        MemberDisplayName            = ''
                        MemberODataType              = ''
                        Role                         = ''
                        RolesRaw                     = ''
                    })))
            $rowNumber++
            continue
        }

        $rowsAddedForInput = 0
        foreach ($group in @($candidateGroups | Sort-Object -Property MailNickname, DisplayName, Id)) {
            $groupId = ([string]$group.Id).Trim()
            $groupAlias = ([string]$group.MailNickname).Trim()
            $groupName = ([string]$group.DisplayName).Trim()

            $team = $null
            if ($teamByGroupId.ContainsKey($groupId)) {
                $team = $teamByGroupId[$groupId]
            }
            else {
                $team = Invoke-WithRetry -OperationName "Lookup Team for group $groupAlias" -ScriptBlock {
                    Get-MgGroupTeam -GroupId $groupId -ErrorAction SilentlyContinue
                }
                $teamByGroupId[$groupId] = $team
            }

            if (-not $team) {
                continue
            }

            $channels = @()
            if ($channelsByGroupId.ContainsKey($groupId)) {
                $channels = @($channelsByGroupId[$groupId])
            }
            else {
                $channels = @(Invoke-WithRetry -OperationName "Load channels for Team $groupAlias" -ScriptBlock {
                    Get-MgTeamChannel -TeamId $groupId -All -ErrorAction Stop
                })
                $channelsByGroupId[$groupId] = @($channels)
            }

            $selectedChannels = @()
            if ($requestedChannelDisplayName -eq '*') {
                $selectedChannels = @($channels)
            }
            else {
                $selectedChannels = @($channels | Where-Object {
                        ([string]$_.DisplayName).Trim().Equals($requestedChannelDisplayName, [System.StringComparison]::OrdinalIgnoreCase)
                    })
            }

            if ($selectedChannels.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$groupAlias|$requestedChannelDisplayName" -Action 'GetMicrosoftTeamChannelMember' -Status 'NotFound' -Message 'No matching channels were found for the Team.' -Data ([ordered]@{
                            TeamGroupId                  = $groupId
                            TeamDisplayName              = $groupName
                            TeamMailNickname             = $groupAlias
                            ChannelId                    = ''
                            ChannelDisplayName           = ''
                            ChannelMembershipType        = ''
                            ChannelEmail                 = ''
                            ChannelWebUrl                = ''
                            RequestedChannelDisplayName  = $requestedChannelDisplayName
                            MemberConversationId         = ''
                            MemberObjectId               = ''
                            MemberEmail                  = ''
                            MemberDisplayName            = ''
                            MemberODataType              = ''
                            Role                         = ''
                            RolesRaw                     = ''
                        })))
                $rowsAddedForInput++
                continue
            }

            foreach ($channel in @($selectedChannels | Sort-Object -Property DisplayName, Id)) {
                $channelId = ([string]$channel.Id).Trim()
                $channelName = ([string]$channel.DisplayName).Trim()
                $membershipType = ([string]$channel.MembershipType).Trim()
                if ([string]::IsNullOrWhiteSpace($membershipType)) {
                    $membershipType = 'Standard'
                }

                $channelEmail = Get-GraphPropertyValue -Object $channel -PropertyNames @('Email', 'email')
                $channelWebUrl = Get-GraphPropertyValue -Object $channel -PropertyNames @('WebUrl', 'webUrl')

                if ($membershipType.Equals('Standard', [System.StringComparison]::OrdinalIgnoreCase)) {
                    $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$groupAlias|$channelId" -Action 'GetMicrosoftTeamChannelMember' -Status 'Completed' -Message 'Standard channels inherit Team membership; use IR3310 for Team-level members.' -Data ([ordered]@{
                                TeamGroupId                  = $groupId
                                TeamDisplayName              = $groupName
                                TeamMailNickname             = $groupAlias
                                ChannelId                    = $channelId
                                ChannelDisplayName           = $channelName
                                ChannelMembershipType        = $membershipType
                                ChannelEmail                 = $channelEmail
                                ChannelWebUrl                = $channelWebUrl
                                RequestedChannelDisplayName  = $requestedChannelDisplayName
                                MemberConversationId         = ''
                                MemberObjectId               = ''
                                MemberEmail                  = ''
                                MemberDisplayName            = ''
                                MemberODataType              = ''
                                Role                         = ''
                                RolesRaw                     = ''
                            })))
                    $rowsAddedForInput++
                    continue
                }

                $members = @(Invoke-WithRetry -OperationName "Load channel members $groupAlias/$channelName" -ScriptBlock {
                    Get-MgTeamChannelMember -TeamId $groupId -ChannelId $channelId -All -ErrorAction Stop
                })

                if ($members.Count -eq 0) {
                    $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$groupAlias|$channelId" -Action 'GetMicrosoftTeamChannelMember' -Status 'Completed' -Message 'No explicit members were returned for this private/shared channel.' -Data ([ordered]@{
                                TeamGroupId                  = $groupId
                                TeamDisplayName              = $groupName
                                TeamMailNickname             = $groupAlias
                                ChannelId                    = $channelId
                                ChannelDisplayName           = $channelName
                                ChannelMembershipType        = $membershipType
                                ChannelEmail                 = $channelEmail
                                ChannelWebUrl                = $channelWebUrl
                                RequestedChannelDisplayName  = $requestedChannelDisplayName
                                MemberConversationId         = ''
                                MemberObjectId               = ''
                                MemberEmail                  = ''
                                MemberDisplayName            = ''
                                MemberODataType              = ''
                                Role                         = ''
                                RolesRaw                     = ''
                            })))
                    $rowsAddedForInput++
                    continue
                }

                foreach ($member in @($members | Sort-Object -Property Id)) {
                    $conversationMemberId = ([string]$member.Id).Trim()
                    $memberObjectId = Get-GraphPropertyValue -Object $member -PropertyNames @('UserId', 'userId')
                    $memberEmail = Get-GraphPropertyValue -Object $member -PropertyNames @('Email', 'email')
                    $memberDisplayName = Get-GraphPropertyValue -Object $member -PropertyNames @('DisplayName', 'displayName')
                    $memberODataType = Get-GraphPropertyValue -Object $member -PropertyNames @('@odata.type')

                    $roleInfo = Get-ChannelMemberRoles -Member $member

                    $memberKey = if (-not [string]::IsNullOrWhiteSpace($conversationMemberId)) {
                        $conversationMemberId
                    }
                    elseif (-not [string]::IsNullOrWhiteSpace($memberObjectId)) {
                        $memberObjectId
                    }
                    elseif (-not [string]::IsNullOrWhiteSpace($memberEmail)) {
                        $memberEmail.ToLowerInvariant()
                    }
                    else {
                        [guid]::NewGuid().ToString()
                    }

                    $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$groupAlias|$channelId|$memberKey" -Action 'GetMicrosoftTeamChannelMember' -Status 'Completed' -Message 'Channel member exported.' -Data ([ordered]@{
                                TeamGroupId                  = $groupId
                                TeamDisplayName              = $groupName
                                TeamMailNickname             = $groupAlias
                                ChannelId                    = $channelId
                                ChannelDisplayName           = $channelName
                                ChannelMembershipType        = $membershipType
                                ChannelEmail                 = $channelEmail
                                ChannelWebUrl                = $channelWebUrl
                                RequestedChannelDisplayName  = $requestedChannelDisplayName
                                MemberConversationId         = $conversationMemberId
                                MemberObjectId               = $memberObjectId
                                MemberEmail                  = $memberEmail
                                MemberDisplayName            = $memberDisplayName
                                MemberODataType              = $memberODataType
                                Role                         = $roleInfo.Role
                                RolesRaw                     = $roleInfo.RolesRaw
                            })))
                    $rowsAddedForInput++
                }
            }
        }

        if ($rowsAddedForInput -eq 0) {
            $message = if ($teamMailNickname -eq '*') { 'No Teams were found for the selected scope.' } else { "Group '$teamMailNickname' exists, but no Team is provisioned for it." }
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$requestedChannelDisplayName" -Action 'GetMicrosoftTeamChannelMember' -Status 'NotFound' -Message $message -Data ([ordered]@{
                        TeamGroupId                  = ''
                        TeamDisplayName              = ''
                        TeamMailNickname             = $teamMailNickname
                        ChannelId                    = ''
                        ChannelDisplayName           = ''
                        ChannelMembershipType        = ''
                        ChannelEmail                 = ''
                        ChannelWebUrl                = ''
                        RequestedChannelDisplayName  = $requestedChannelDisplayName
                        MemberConversationId         = ''
                        MemberObjectId               = ''
                        MemberEmail                  = ''
                        MemberDisplayName            = ''
                        MemberODataType              = ''
                        Role                         = ''
                        RolesRaw                     = ''
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($teamMailNickname|$requestedChannelDisplayName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$teamMailNickname|$requestedChannelDisplayName" -Action 'GetMicrosoftTeamChannelMember' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    TeamGroupId                  = ''
                    TeamDisplayName              = ''
                    TeamMailNickname             = $teamMailNickname
                    ChannelId                    = ''
                    ChannelDisplayName           = ''
                    ChannelMembershipType        = ''
                    ChannelEmail                 = ''
                    ChannelWebUrl                = ''
                    RequestedChannelDisplayName  = $requestedChannelDisplayName
                    MemberConversationId         = ''
                    MemberObjectId               = ''
                    MemberEmail                  = ''
                    MemberDisplayName            = ''
                    MemberODataType              = ''
                    Role                         = ''
                    RolesRaw                     = ''
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
Write-Status -Message 'Microsoft Teams channel member inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}




