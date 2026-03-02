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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3309-Get-MicrosoftTeams_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-FirstPropertyValue {
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

    return ''
}

$requiredHeaders = @(
    'TeamMailNickname'
)

Write-Status -Message 'Starting Microsoft Teams inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Teams')
Ensure-GraphConnection -RequiredScopes @('Group.Read.All', 'Team.ReadBasic.All', 'TeamSettings.Read.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$groupSelect = 'id,displayName,description,mailNickname,visibility,groupTypes,mailEnabled,securityEnabled,createdDateTime'
$allM365GroupsCache = $null
$teamByGroupId = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

$rowNumber = 1
foreach ($row in $rows) {
    $teamMailNickname = ([string]$row.TeamMailNickname).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($teamMailNickname)) {
            throw 'TeamMailNickname is required. Use * to inventory all Teams.'
        }

        $candidateGroups = @()
        if ($teamMailNickname -eq '*') {
            if ($null -eq $allM365GroupsCache) {
                $allGroups = @(Invoke-WithRetry -OperationName 'Load all groups for Teams inventory' -ScriptBlock {
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
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $teamMailNickname -Action 'GetMicrosoftTeam' -Status 'NotFound' -Message 'No matching Microsoft 365 groups were found.' -Data ([ordered]@{
                        TeamGroupId                            = ''
                        TeamDisplayName                        = ''
                        TeamMailNickname                       = $teamMailNickname
                        Description                            = ''
                        Visibility                             = ''
                        CreatedDateTime                        = ''
                        IsArchived                             = ''
                        AllowCreateUpdateChannels              = ''
                        AllowDeleteChannels                    = ''
                        AllowAddRemoveApps                     = ''
                        AllowCreateUpdateRemoveTabs            = ''
                        AllowCreateUpdateRemoveConnectors      = ''
                        AllowUserEditMessages                  = ''
                        AllowUserDeleteMessages                = ''
                        AllowOwnerDeleteMessages               = ''
                        AllowTeamMentions                      = ''
                        AllowChannelMentions                   = ''
                        AllowGiphy                             = ''
                        GiphyContentRating                     = ''
                        AllowStickersAndMemes                  = ''
                        AllowCustomMemes                       = ''
                    })))
            $rowNumber++
            continue
        }

        $teamRowsAdded = 0
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

            $memberSettings = $null
            if ($team.PSObject.Properties.Name -contains 'MemberSettings') {
                $memberSettings = $team.MemberSettings
            }

            $messagingSettings = $null
            if ($team.PSObject.Properties.Name -contains 'MessagingSettings') {
                $messagingSettings = $team.MessagingSettings
            }

            $funSettings = $null
            if ($team.PSObject.Properties.Name -contains 'FunSettings') {
                $funSettings = $team.FunSettings
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$groupAlias|$groupId" -Action 'GetMicrosoftTeam' -Status 'Completed' -Message 'Team exported.' -Data ([ordered]@{
                        TeamGroupId                            = $groupId
                        TeamDisplayName                        = $groupName
                        TeamMailNickname                       = $groupAlias
                        Description                            = ([string]$group.Description).Trim()
                        Visibility                             = ([string]$group.Visibility).Trim()
                        CreatedDateTime                        = [string]$group.CreatedDateTime
                        IsArchived                             = Get-FirstPropertyValue -Object $team -PropertyNames @('IsArchived', 'isArchived')
                        AllowCreateUpdateChannels              = Get-FirstPropertyValue -Object $memberSettings -PropertyNames @('AllowCreateUpdateChannels', 'allowCreateUpdateChannels')
                        AllowDeleteChannels                    = Get-FirstPropertyValue -Object $memberSettings -PropertyNames @('AllowDeleteChannels', 'allowDeleteChannels')
                        AllowAddRemoveApps                     = Get-FirstPropertyValue -Object $memberSettings -PropertyNames @('AllowAddRemoveApps', 'allowAddRemoveApps')
                        AllowCreateUpdateRemoveTabs            = Get-FirstPropertyValue -Object $memberSettings -PropertyNames @('AllowCreateUpdateRemoveTabs', 'allowCreateUpdateRemoveTabs')
                        AllowCreateUpdateRemoveConnectors      = Get-FirstPropertyValue -Object $memberSettings -PropertyNames @('AllowCreateUpdateRemoveConnectors', 'allowCreateUpdateRemoveConnectors')
                        AllowUserEditMessages                  = Get-FirstPropertyValue -Object $messagingSettings -PropertyNames @('AllowUserEditMessages', 'allowUserEditMessages')
                        AllowUserDeleteMessages                = Get-FirstPropertyValue -Object $messagingSettings -PropertyNames @('AllowUserDeleteMessages', 'allowUserDeleteMessages')
                        AllowOwnerDeleteMessages               = Get-FirstPropertyValue -Object $messagingSettings -PropertyNames @('AllowOwnerDeleteMessages', 'allowOwnerDeleteMessages')
                        AllowTeamMentions                      = Get-FirstPropertyValue -Object $messagingSettings -PropertyNames @('AllowTeamMentions', 'allowTeamMentions')
                        AllowChannelMentions                   = Get-FirstPropertyValue -Object $messagingSettings -PropertyNames @('AllowChannelMentions', 'allowChannelMentions')
                        AllowGiphy                             = Get-FirstPropertyValue -Object $funSettings -PropertyNames @('AllowGiphy', 'allowGiphy')
                        GiphyContentRating                     = Get-FirstPropertyValue -Object $funSettings -PropertyNames @('GiphyContentRating', 'giphyContentRating')
                        AllowStickersAndMemes                  = Get-FirstPropertyValue -Object $funSettings -PropertyNames @('AllowStickersAndMemes', 'allowStickersAndMemes')
                        AllowCustomMemes                       = Get-FirstPropertyValue -Object $funSettings -PropertyNames @('AllowCustomMemes', 'allowCustomMemes')
                    })))

            $teamRowsAdded++
        }

        if ($teamRowsAdded -eq 0) {
            $message = if ($teamMailNickname -eq '*') { 'No Teams were found for the selected scope.' } else { "Group '$teamMailNickname' exists, but no Team is provisioned for it." }
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $teamMailNickname -Action 'GetMicrosoftTeam' -Status 'NotFound' -Message $message -Data ([ordered]@{
                        TeamGroupId                            = ''
                        TeamDisplayName                        = ''
                        TeamMailNickname                       = $teamMailNickname
                        Description                            = ''
                        Visibility                             = ''
                        CreatedDateTime                        = ''
                        IsArchived                             = ''
                        AllowCreateUpdateChannels              = ''
                        AllowDeleteChannels                    = ''
                        AllowAddRemoveApps                     = ''
                        AllowCreateUpdateRemoveTabs            = ''
                        AllowCreateUpdateRemoveConnectors      = ''
                        AllowUserEditMessages                  = ''
                        AllowUserDeleteMessages                = ''
                        AllowOwnerDeleteMessages               = ''
                        AllowTeamMentions                      = ''
                        AllowChannelMentions                   = ''
                        AllowGiphy                             = ''
                        GiphyContentRating                     = ''
                        AllowStickersAndMemes                  = ''
                        AllowCustomMemes                       = ''
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($teamMailNickname) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $teamMailNickname -Action 'GetMicrosoftTeam' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    TeamGroupId                            = ''
                    TeamDisplayName                        = ''
                    TeamMailNickname                       = $teamMailNickname
                    Description                            = ''
                    Visibility                             = ''
                    CreatedDateTime                        = ''
                    IsArchived                             = ''
                    AllowCreateUpdateChannels              = ''
                    AllowDeleteChannels                    = ''
                    AllowAddRemoveApps                     = ''
                    AllowCreateUpdateRemoveTabs            = ''
                    AllowCreateUpdateRemoveConnectors      = ''
                    AllowUserEditMessages                  = ''
                    AllowUserDeleteMessages                = ''
                    AllowOwnerDeleteMessages               = ''
                    AllowTeamMentions                      = ''
                    AllowChannelMentions                   = ''
                    AllowGiphy                             = ''
                    GiphyContentRating                     = ''
                    AllowStickersAndMemes                  = ''
                    AllowCustomMemes                       = ''
                })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Microsoft Teams inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







