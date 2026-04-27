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
Microsoft.Graph.Teams
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets MicrosoftTeamChannelsByUser and exports results to CSV.

.DESCRIPTION
    Gets MicrosoftTeamChannelsByUser from Microsoft 365 and writes the results to a CSV file.
    For each in-scope user, discovers every Microsoft Teams channel they have access to across
    all teams, with channel type distinction (Standard, Private, Shared) and the mechanism
    that provides access (TeamMembership, ExplicitChannelMember, CrossTenantShared).
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
    .\D-TEAM-0510-Get-MicrosoftTeamChannelsByUser.ps1 -InputCsvPath .\Scope-Users.input.csv
    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\D-TEAM-0510-Get-MicrosoftTeamChannelsByUser.ps1 -DiscoverAll
    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Teams, Microsoft.Graph.Users
    Required roles:   Teams Administrator or Global Reader
    Limitations:      None known.

    CSV Fields:
    Column              Type    Required  Description
    ------              ----    --------  -----------
    UserPrincipalName   String  Yes       UPN of the user whose Teams channel access to inventory
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-TEAM-0510-Get-MicrosoftTeamChannelsByUser_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-AdditionalPropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Object,

        [Parameter(Mandatory)]
        [string]$Key
    )

    if ($null -eq $Object) { return '' }

    if ($Object.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $additional = $Object.AdditionalProperties
        if ($null -ne $additional) {
            try {
                if ($additional.ContainsKey($Key)) {
                    return [string]$additional[$Key]
                }
            }
            catch {
                # Best effort only.
            }
        }
    }

    return ''
}

function Get-ChannelMemberUserId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Member
    )

    # userId is typically in AdditionalProperties for aadUserConversationMember
    $userId = Get-AdditionalPropertyValue -Object $Member -Key 'userId'
    if (-not [string]::IsNullOrWhiteSpace($userId)) { return $userId.Trim() }

    # Fallback: try top-level property
    if ($Member.PSObject.Properties.Name -contains 'UserId') {
        return ([string]$Member.UserId).Trim()
    }

    return ''
}

function Get-ChannelMemberTenantId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Member
    )

    $tenantId = Get-AdditionalPropertyValue -Object $Member -Key 'tenantId'
    if (-not [string]::IsNullOrWhiteSpace($tenantId)) { return $tenantId.Trim() }

    return ''
}

function Get-ChannelMemberRole {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Member
    )

    $roles = @()

    if ($Member.PSObject.Properties.Name -contains 'Roles') {
        $roles = @($Member.Roles)
    }

    if ($roles.Count -eq 0) {
        $rolesRaw = Get-AdditionalPropertyValue -Object $Member -Key 'roles'
        if (-not [string]::IsNullOrWhiteSpace($rolesRaw)) {
            $roles = @($rolesRaw -split ',')
        }
    }

    $normalized = @(
        $roles |
            ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    return if ($normalized -contains 'owner') { 'Owner' } else { 'Member' }
}

$requiredHeaders = @('UserPrincipalName')

$reportPropertyOrder = @(
    'TimestampUtc', 'RowNumber', 'PrimaryKey', 'Action', 'Status', 'Message', 'ScopeMode',
    'UserPrincipalName', 'TeamId', 'TeamDisplayName',
    'ChannelId', 'ChannelDisplayName', 'ChannelType',
    'ChannelMemberRole', 'AccessMechanism', 'ExternalTenantId', 'AccessExplained'
)

Write-Status -Message 'Starting Microsoft Teams channel access by user inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Teams', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('Team.ReadBasic.All', 'Channel.ReadBasic.All', 'ChannelMember.Read.All', 'TeamMember.Read.All')

# Resolve host tenant ID once — used to detect cross-tenant shared channel members
$hostTenantId = Invoke-WithRetry -OperationName 'Resolve host tenant ID' -ScriptBlock {
    Get-MgOrganization -All -Property 'Id' -ErrorAction Stop |
        Select-Object -First 1 -ExpandProperty Id
}
Write-Status -Message "Host tenant ID: $hostTenantId"

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

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $upn = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw 'UserPrincipalName is required and cannot be blank.'
        }

        Write-Status -Message "Processing user $rowNumber of $($rows.Count): $upn"

        # Resolve user object
        $userObj = Invoke-WithRetry -OperationName "Resolve user $upn" -ScriptBlock {
            Get-MgUser -UserId $upn -Property 'Id,DisplayName,UserPrincipalName,UserType' -ErrorAction Stop
        }
        $userId = ([string]$userObj.Id).Trim()

        # Get all joined teams for the user
        $joinedTeams = @(Invoke-WithRetry -OperationName "Get joined teams for $upn" -ScriptBlock {
            Get-MgUserJoinedTeam -UserId $userId -All -Property 'Id,DisplayName' -ErrorAction Stop
        })
        Write-Status -Message "  Found $($joinedTeams.Count) team(s) for $upn."

        if ($joinedTeams.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $upn -Action 'GetMicrosoftTeamChannelsByUser' -Status 'Success' -Message 'No Teams channels found.' -Data ([ordered]@{
                UserPrincipalName = $upn
                TeamId            = ''
                TeamDisplayName   = ''
                ChannelId         = ''
                ChannelDisplayName = ''
                ChannelType       = ''
                ChannelMemberRole = ''
                AccessMechanism   = ''
                ExternalTenantId  = ''
                AccessExplained   = ''
            })))
            $rowNumber++
            continue
        }

        # Determine if the user is an owner at team level — used for standard channel role attribution
        # Cache owners per team to avoid redundant Graph calls within a user's team set
        $teamOwnerCache = [System.Collections.Generic.Dictionary[string, bool]]::new([System.StringComparer]::OrdinalIgnoreCase)

        $channelsAddedForUser = 0

        $sortedTeams = @($joinedTeams | Sort-Object -Property DisplayName, Id)
        $teamIndex   = 0
        foreach ($team in $sortedTeams) {
            $teamIndex++
            $teamId          = ([string]$team.Id).Trim()
            $teamDisplayName = ([string]$team.DisplayName).Trim()

            # Determine whether user is a team owner (used for standard channel role)
            $isTeamOwner = $false
            if ($teamOwnerCache.ContainsKey($teamId)) {
                $isTeamOwner = $teamOwnerCache[$teamId]
            }
            else {
                try {
                    $owners = @(Invoke-WithRetry -OperationName "Get owners of team $teamId" -ScriptBlock {
                        Get-MgGroupOwner -GroupId $teamId -All -Property 'Id' -ErrorAction Stop
                    })
                    $ownerIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                    foreach ($owner in $owners) {
                        [void]$ownerIds.Add(([string]$owner.Id).Trim())
                    }
                    $isTeamOwner = $ownerIds.Contains($userId)
                }
                catch {
                    Write-Status -Message "Could not retrieve owners for team $teamId ($teamDisplayName): $($_.Exception.Message)" -Level WARN
                    $isTeamOwner = $false
                }
                $teamOwnerCache[$teamId] = $isTeamOwner
            }

            # Enumerate all channels for the team
            $channels = @(Invoke-WithRetry -OperationName "Get channels for team $teamId" -ScriptBlock {
                Get-MgTeamChannel -TeamId $teamId -All -Property 'Id,DisplayName,MembershipType,Description' -ErrorAction Stop
            })
            Write-Status -Message "  Team $teamIndex of $($sortedTeams.Count): '$teamDisplayName' ($($channels.Count) channel(s))."

            $sortedChannels = @($channels | Sort-Object -Property DisplayName, Id)
            $channelIndex   = 0
            foreach ($channel in $sortedChannels) {
                $channelIndex++
                $channelId          = ([string]$channel.Id).Trim()
                $channelDisplayName = ([string]$channel.DisplayName).Trim()
                $membershipTypeRaw  = ([string]$channel.MembershipType).Trim()

                # Normalise MembershipType to the ChannelType vocabulary
                $channelType = switch ($membershipTypeRaw.ToLowerInvariant()) {
                    'private' { 'Private'  }
                    'shared'  { 'Shared'   }
                    default   { 'Standard' }
                }

                Write-Status -Message "    Channel $channelIndex of $($sortedChannels.Count): '$channelDisplayName' [$channelType]."

                if ($channelType -eq 'Standard') {
                    # Access is inherited from team membership
                    $channelMemberRole = if ($isTeamOwner) { 'Owner' } else { 'Member' }
                    $accessMechanism   = 'TeamMembership'
                    $externalTenantId  = ''
                    $accessExplained   = 'Access inherited from team membership.'

                    $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$upn|$teamId|$channelId" -Action 'GetMicrosoftTeamChannelsByUser' -Status 'Success' -Message 'Standard channel access exported.' -Data ([ordered]@{
                        UserPrincipalName  = $upn
                        TeamId             = $teamId
                        TeamDisplayName    = $teamDisplayName
                        ChannelId          = $channelId
                        ChannelDisplayName = $channelDisplayName
                        ChannelType        = $channelType
                        ChannelMemberRole  = $channelMemberRole
                        AccessMechanism    = $accessMechanism
                        ExternalTenantId   = $externalTenantId
                        AccessExplained    = $accessExplained
                    })))
                    $channelsAddedForUser++
                }
                elseif ($channelType -eq 'Private') {
                    # Must check explicit channel membership
                    $channelMembers = @(Invoke-WithRetry -OperationName "Get members of private channel $channelId in team $teamId" -ScriptBlock {
                        Get-MgTeamChannelMember -TeamId $teamId -ChannelId $channelId -All -Property 'Id,Roles,DisplayName' -ErrorAction Stop
                    })

                    # Find the member entry matching the current user
                    $userMemberEntry = $null
                    foreach ($member in $channelMembers) {
                        $memberId = Get-ChannelMemberUserId -Member $member
                        if ($memberId -eq $userId) {
                            $userMemberEntry = $member
                            break
                        }
                    }

                    if ($null -eq $userMemberEntry) {
                        # User is not explicitly listed in this private channel — skip
                        continue
                    }

                    $channelMemberRole = Get-ChannelMemberRole -Member $userMemberEntry
                    $accessMechanism   = 'ExplicitChannelMember'
                    $externalTenantId  = ''
                    $accessExplained   = 'Explicitly added to private channel.'

                    $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$upn|$teamId|$channelId" -Action 'GetMicrosoftTeamChannelsByUser' -Status 'Success' -Message 'Private channel access exported.' -Data ([ordered]@{
                        UserPrincipalName  = $upn
                        TeamId             = $teamId
                        TeamDisplayName    = $teamDisplayName
                        ChannelId          = $channelId
                        ChannelDisplayName = $channelDisplayName
                        ChannelType        = $channelType
                        ChannelMemberRole  = $channelMemberRole
                        AccessMechanism    = $accessMechanism
                        ExternalTenantId   = $externalTenantId
                        AccessExplained    = $accessExplained
                    })))
                    $channelsAddedForUser++
                }
                elseif ($channelType -eq 'Shared') {
                    # Must check explicit channel membership; also detect cross-tenant access
                    $channelMembers = @(Invoke-WithRetry -OperationName "Get members of shared channel $channelId in team $teamId" -ScriptBlock {
                        Get-MgTeamChannelMember -TeamId $teamId -ChannelId $channelId -All -Property 'Id,Roles,DisplayName' -ErrorAction Stop
                    })

                    # Find the member entry matching the current user
                    $userMemberEntry = $null
                    foreach ($member in $channelMembers) {
                        $memberId = Get-ChannelMemberUserId -Member $member
                        if ($memberId -eq $userId) {
                            $userMemberEntry = $member
                            break
                        }
                    }

                    if ($null -eq $userMemberEntry) {
                        # User is not explicitly listed in this shared channel — skip
                        continue
                    }

                    $channelMemberRole = Get-ChannelMemberRole -Member $userMemberEntry

                    # Detect cross-tenant access by comparing member tenantId to the host tenant
                    $memberTenantId = Get-ChannelMemberTenantId -Member $userMemberEntry
                    $isCrossTenant  = (
                        -not [string]::IsNullOrWhiteSpace($memberTenantId) -and
                        -not [string]::IsNullOrWhiteSpace($hostTenantId) -and
                        -not $memberTenantId.Equals($hostTenantId, [System.StringComparison]::OrdinalIgnoreCase)
                    )

                    if ($isCrossTenant) {
                        $accessMechanism  = 'CrossTenantShared'
                        $externalTenantId = $memberTenantId
                        $accessExplained  = 'Cross-tenant shared channel access.'
                    }
                    else {
                        $accessMechanism  = 'ExplicitChannelMember'
                        $externalTenantId = ''
                        $accessExplained  = 'Explicitly added to shared channel.'
                    }

                    $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$upn|$teamId|$channelId" -Action 'GetMicrosoftTeamChannelsByUser' -Status 'Success' -Message 'Shared channel access exported.' -Data ([ordered]@{
                        UserPrincipalName  = $upn
                        TeamId             = $teamId
                        TeamDisplayName    = $teamDisplayName
                        ChannelId          = $channelId
                        ChannelDisplayName = $channelDisplayName
                        ChannelType        = $channelType
                        ChannelMemberRole  = $channelMemberRole
                        AccessMechanism    = $accessMechanism
                        ExternalTenantId   = $externalTenantId
                        AccessExplained    = $accessExplained
                    })))
                    $channelsAddedForUser++
                }
            }
        }

        if ($channelsAddedForUser -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $upn -Action 'GetMicrosoftTeamChannelsByUser' -Status 'Success' -Message 'No Teams channels found.' -Data ([ordered]@{
                UserPrincipalName  = $upn
                TeamId             = ''
                TeamDisplayName    = ''
                ChannelId          = ''
                ChannelDisplayName = ''
                ChannelType        = ''
                ChannelMemberRole  = ''
                AccessMechanism    = ''
                ExternalTenantId   = ''
                AccessExplained    = ''
            })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $upn -Action 'GetMicrosoftTeamChannelsByUser' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            UserPrincipalName  = $upn
            TeamId             = ''
            TeamDisplayName    = ''
            ChannelId          = ''
            ChannelDisplayName = ''
            ChannelType        = ''
            ChannelMemberRole  = ''
            AccessMechanism    = ''
            ExternalTenantId   = ''
            AccessExplained    = ''
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
Write-Status -Message 'Script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
