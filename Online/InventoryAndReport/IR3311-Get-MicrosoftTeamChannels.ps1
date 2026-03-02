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
Microsoft.Graph.Teams

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3311-Get-MicrosoftTeamChannels_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$requiredHeaders = @(
    'TeamMailNickname'
)

Write-Status -Message 'Starting Microsoft Teams channel inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Teams')
Ensure-GraphConnection -RequiredScopes @('Group.Read.All', 'Team.ReadBasic.All', 'TeamSettings.Read.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
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

    try {
        if ([string]::IsNullOrWhiteSpace($teamMailNickname)) {
            throw 'TeamMailNickname is required. Use * to inventory channels for all Teams.'
        }

        $candidateGroups = @()
        if ($teamMailNickname -eq '*') {
            if ($null -eq $allM365GroupsCache) {
                $allGroups = @(Invoke-WithRetry -OperationName 'Load all groups for Team channel inventory' -ScriptBlock {
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
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $teamMailNickname -Action 'GetMicrosoftTeamChannel' -Status 'NotFound' -Message 'No matching Microsoft 365 groups were found.' -Data ([ordered]@{
                        TeamGroupId         = ''
                        TeamDisplayName     = ''
                        TeamMailNickname    = $teamMailNickname
                        ChannelId           = ''
                        ChannelDisplayName  = ''
                        MembershipType      = ''
                        Description         = ''
                        IsFavoriteByDefault = ''
                        Email               = ''
                        WebUrl              = ''
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

            if ($channels.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$groupAlias|$groupId" -Action 'GetMicrosoftTeamChannel' -Status 'Completed' -Message 'Team has no channels.' -Data ([ordered]@{
                            TeamGroupId         = $groupId
                            TeamDisplayName     = $groupName
                            TeamMailNickname    = $groupAlias
                            ChannelId           = ''
                            ChannelDisplayName  = ''
                            MembershipType      = ''
                            Description         = ''
                            IsFavoriteByDefault = ''
                            Email               = ''
                            WebUrl              = ''
                        })))
                $rowsAddedForInput++
                continue
            }

            foreach ($channel in @($channels | Sort-Object -Property DisplayName, Id)) {
                $channelId = ([string]$channel.Id).Trim()
                $channelName = ([string]$channel.DisplayName).Trim()
                $membershipType = ([string]$channel.MembershipType).Trim()
                if ([string]::IsNullOrWhiteSpace($membershipType)) {
                    $membershipType = 'Standard'
                }

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$groupAlias|$channelId" -Action 'GetMicrosoftTeamChannel' -Status 'Completed' -Message 'Team channel exported.' -Data ([ordered]@{
                            TeamGroupId         = $groupId
                            TeamDisplayName     = $groupName
                            TeamMailNickname    = $groupAlias
                            ChannelId           = $channelId
                            ChannelDisplayName  = $channelName
                            MembershipType      = $membershipType
                            Description         = ([string]$channel.Description).Trim()
                            IsFavoriteByDefault = Get-GraphPropertyValue -Object $channel -PropertyNames @('IsFavoriteByDefault', 'isFavoriteByDefault')
                            Email               = Get-GraphPropertyValue -Object $channel -PropertyNames @('Email', 'email')
                            WebUrl              = Get-GraphPropertyValue -Object $channel -PropertyNames @('WebUrl', 'webUrl')
                        })))
                $rowsAddedForInput++
            }
        }

        if ($rowsAddedForInput -eq 0) {
            $message = if ($teamMailNickname -eq '*') { 'No Teams were found for the selected scope.' } else { "Group '$teamMailNickname' exists, but no Team is provisioned for it." }
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $teamMailNickname -Action 'GetMicrosoftTeamChannel' -Status 'NotFound' -Message $message -Data ([ordered]@{
                        TeamGroupId         = ''
                        TeamDisplayName     = ''
                        TeamMailNickname    = $teamMailNickname
                        ChannelId           = ''
                        ChannelDisplayName  = ''
                        MembershipType      = ''
                        Description         = ''
                        IsFavoriteByDefault = ''
                        Email               = ''
                        WebUrl              = ''
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($teamMailNickname) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $teamMailNickname -Action 'GetMicrosoftTeamChannel' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    TeamGroupId         = ''
                    TeamDisplayName     = ''
                    TeamMailNickname    = $teamMailNickname
                    ChannelId           = ''
                    ChannelDisplayName  = ''
                    MembershipType      = ''
                    Description         = ''
                    IsFavoriteByDefault = ''
                    Email               = ''
                    WebUrl              = ''
                })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Microsoft Teams channel inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







