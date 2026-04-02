<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-193000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Groups
Microsoft.Graph.Teams

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets MicrosoftTeams and exports results to CSV.

.DESCRIPTION
    Gets MicrosoftTeams from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-IR3309-Get-MicrosoftTeams.ps1 -InputCsvPath .\3309.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3309-Get-MicrosoftTeams.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Teams
    Required roles:   Global Administrator or appropriate workload-specific role
    Limitations:      None known.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-TEAM-0010-Get-MicrosoftTeams_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'Description',
    'Visibility',
    'CreatedDateTime',
    'IsArchived',
    'AllowCreateUpdateChannels',
    'AllowDeleteChannels',
    'AllowAddRemoveApps',
    'AllowCreateUpdateRemoveTabs',
    'AllowCreateUpdateRemoveConnectors',
    'AllowUserEditMessages',
    'AllowUserDeleteMessages',
    'AllowOwnerDeleteMessages',
    'AllowTeamMentions',
    'AllowChannelMentions',
    'AllowGiphy',
    'GiphyContentRating',
    'AllowStickersAndMemes',
    'AllowCustomMemes'
)

Write-Status -Message 'Starting Microsoft Teams inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Teams')
Ensure-GraphConnection -RequiredScopes @('Group.Read.All', 'Team.ReadBasic.All', 'TeamSettings.Read.All', 'Directory.Read.All')

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

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Microsoft Teams inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}











