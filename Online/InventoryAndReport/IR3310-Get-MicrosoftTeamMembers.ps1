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

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3310-Get-MicrosoftTeamMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-DirectoryObjectType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$DirectoryObject
    )

    $odataType = ''
    if ($DirectoryObject.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $additional = $DirectoryObject.AdditionalProperties
        if ($additional) {
            try {
                if ($additional.ContainsKey('@odata.type')) {
                    $odataType = ([string]$additional['@odata.type']).Trim()
                }
            }
            catch {
                # Best effort only.
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($odataType)) {
        return 'unknown'
    }

    $normalized = $odataType.TrimStart('#').ToLowerInvariant()
    if ($normalized.StartsWith('microsoft.graph.')) {
        $normalized = $normalized.Substring('microsoft.graph.'.Length)
    }

    return $normalized
}

function Get-AdditionalPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$DirectoryObject,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    if ($DirectoryObject.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $additional = $DirectoryObject.AdditionalProperties
        if ($additional) {
            try {
                if ($additional.ContainsKey($PropertyName)) {
                    return ([string]$additional[$PropertyName]).Trim()
                }
            }
            catch {
                # Best effort only.
            }
        }
    }

    return ''
}

$requiredHeaders = @(
    'TeamMailNickname'
)

Write-Status -Message 'Starting Microsoft Teams member inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users', 'Microsoft.Graph.Teams')
Ensure-GraphConnection -RequiredScopes @('Group.Read.All', 'User.Read.All', 'Directory.Read.All', 'TeamMember.Read.All')

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
$userById = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

$rowNumber = 1
foreach ($row in $rows) {
    $teamMailNickname = ([string]$row.TeamMailNickname).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($teamMailNickname)) {
            throw 'TeamMailNickname is required. Use * to inventory all Team memberships.'
        }

        $candidateGroups = @()
        if ($teamMailNickname -eq '*') {
            if ($null -eq $allM365GroupsCache) {
                $allGroups = @(Invoke-WithRetry -OperationName 'Load all groups for Team member inventory' -ScriptBlock {
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
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $teamMailNickname -Action 'GetMicrosoftTeamMember' -Status 'NotFound' -Message 'No matching Microsoft 365 groups were found.' -Data ([ordered]@{
                        TeamGroupId               = ''
                        TeamDisplayName           = ''
                        TeamMailNickname          = $teamMailNickname
                        MemberObjectId            = ''
                        MemberObjectType          = ''
                        MemberUserPrincipalName   = ''
                        MemberDisplayName         = ''
                        MemberUserType            = ''
                        MemberAccountEnabled      = ''
                        Role                      = ''
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

            $owners = @(Invoke-WithRetry -OperationName "Load Team owners for $groupAlias" -ScriptBlock {
                Get-MgGroupOwner -GroupId $groupId -All -ErrorAction Stop
            })

            $members = @(Invoke-WithRetry -OperationName "Load Team members for $groupAlias" -ScriptBlock {
                Get-MgGroupMember -GroupId $groupId -All -ErrorAction Stop
            })

            $principalMap = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $entryOrder = [System.Collections.Generic.List[string]]::new()

            $addPrincipal = {
                param(
                    [Parameter(Mandatory)]
                    [object]$PrincipalObject,

                    [Parameter(Mandatory)]
                    [string]$Role
                )

                $principalId = ([string]$PrincipalObject.Id).Trim()
                if ([string]::IsNullOrWhiteSpace($principalId)) {
                    return
                }

                $principalType = Get-DirectoryObjectType -DirectoryObject $PrincipalObject
                $principalDisplayName = Get-AdditionalPropertyValue -DirectoryObject $PrincipalObject -PropertyName 'displayName'
                $principalUpn = Get-AdditionalPropertyValue -DirectoryObject $PrincipalObject -PropertyName 'userPrincipalName'
                $principalUserType = ''
                $principalAccountEnabled = ''

                if ($principalType -eq 'user') {
                    try {
                        $user = $null
                        if ($userById.ContainsKey($principalId)) {
                            $user = $userById[$principalId]
                        }
                        else {
                            $user = Invoke-WithRetry -OperationName "Load user details for Team member $principalId" -ScriptBlock {
                                Get-MgUser -UserId $principalId -Property 'id,userPrincipalName,displayName,userType,accountEnabled' -ErrorAction Stop
                            }
                            $userById[$principalId] = $user
                        }

                        $principalUpn = ([string]$user.UserPrincipalName).Trim()
                        $principalDisplayName = ([string]$user.DisplayName).Trim()
                        $principalUserType = ([string]$user.UserType).Trim()
                        $principalAccountEnabled = [string]$user.AccountEnabled
                    }
                    catch {
                        Write-Status -Message "User detail lookup failed for Team principal '$principalId' in Team '$groupAlias': $($_.Exception.Message)" -Level WARN
                    }
                }

                if ($principalMap.ContainsKey($principalId)) {
                    $existing = $principalMap[$principalId]
                    $roleSet = [System.Collections.Generic.HashSet[string]]$existing.RoleSet
                    $null = $roleSet.Add($Role)

                    if ([string]::IsNullOrWhiteSpace(([string]$existing.MemberDisplayName)) -and -not [string]::IsNullOrWhiteSpace($principalDisplayName)) {
                        $existing.MemberDisplayName = $principalDisplayName
                    }
                    if ([string]::IsNullOrWhiteSpace(([string]$existing.MemberUserPrincipalName)) -and -not [string]::IsNullOrWhiteSpace($principalUpn)) {
                        $existing.MemberUserPrincipalName = $principalUpn
                    }
                    if ([string]::IsNullOrWhiteSpace(([string]$existing.MemberUserType)) -and -not [string]::IsNullOrWhiteSpace($principalUserType)) {
                        $existing.MemberUserType = $principalUserType
                    }
                    if ([string]::IsNullOrWhiteSpace(([string]$existing.MemberAccountEnabled)) -and -not [string]::IsNullOrWhiteSpace($principalAccountEnabled)) {
                        $existing.MemberAccountEnabled = $principalAccountEnabled
                    }

                    return
                }

                $roleSetNew = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $null = $roleSetNew.Add($Role)

                $entry = [PSCustomObject]@{
                    MemberObjectId           = $principalId
                    MemberObjectType         = $principalType
                    MemberUserPrincipalName  = $principalUpn
                    MemberDisplayName        = $principalDisplayName
                    MemberUserType           = $principalUserType
                    MemberAccountEnabled     = $principalAccountEnabled
                    RoleSet                  = $roleSetNew
                }

                $principalMap[$principalId] = $entry
                $entryOrder.Add($principalId) | Out-Null
            }

            foreach ($owner in $owners) {
                & $addPrincipal -PrincipalObject $owner -Role 'Owner'
            }

            foreach ($member in $members) {
                & $addPrincipal -PrincipalObject $member -Role 'Member'
            }

            if ($principalMap.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$groupAlias|$groupId" -Action 'GetMicrosoftTeamMember' -Status 'Completed' -Message 'Team has no members.' -Data ([ordered]@{
                            TeamGroupId               = $groupId
                            TeamDisplayName           = $groupName
                            TeamMailNickname          = $groupAlias
                            MemberObjectId            = ''
                            MemberObjectType          = ''
                            MemberUserPrincipalName   = ''
                            MemberDisplayName         = ''
                            MemberUserType            = ''
                            MemberAccountEnabled      = ''
                            Role                      = ''
                        })))
                $rowsAddedForInput++
                continue
            }

            $sortedIds = @(
                $entryOrder |
                    Sort-Object {
                        $entry = $principalMap[$_]
                        if (-not [string]::IsNullOrWhiteSpace(([string]$entry.MemberUserPrincipalName))) {
                            return ([string]$entry.MemberUserPrincipalName).ToLowerInvariant()
                        }
                        if (-not [string]::IsNullOrWhiteSpace(([string]$entry.MemberDisplayName))) {
                            return ([string]$entry.MemberDisplayName).ToLowerInvariant()
                        }
                        return ([string]$entry.MemberObjectId).ToLowerInvariant()
                    }
            )

            foreach ($principalId in $sortedIds) {
                $entry = $principalMap[$principalId]
                $roleSet = [System.Collections.Generic.HashSet[string]]$entry.RoleSet
                $role = if ($roleSet.Contains('Owner')) { 'Owner' } else { 'Member' }

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$groupAlias|$principalId" -Action 'GetMicrosoftTeamMember' -Status 'Completed' -Message 'Team member exported.' -Data ([ordered]@{
                            TeamGroupId               = $groupId
                            TeamDisplayName           = $groupName
                            TeamMailNickname          = $groupAlias
                            MemberObjectId            = $entry.MemberObjectId
                            MemberObjectType          = $entry.MemberObjectType
                            MemberUserPrincipalName   = $entry.MemberUserPrincipalName
                            MemberDisplayName         = $entry.MemberDisplayName
                            MemberUserType            = $entry.MemberUserType
                            MemberAccountEnabled      = $entry.MemberAccountEnabled
                            Role                      = $role
                        })))
                $rowsAddedForInput++
            }
        }

        if ($rowsAddedForInput -eq 0) {
            $message = if ($teamMailNickname -eq '*') { 'No Teams were found for the selected scope.' } else { "Group '$teamMailNickname' exists, but no Team is provisioned for it." }
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $teamMailNickname -Action 'GetMicrosoftTeamMember' -Status 'NotFound' -Message $message -Data ([ordered]@{
                        TeamGroupId               = ''
                        TeamDisplayName           = ''
                        TeamMailNickname          = $teamMailNickname
                        MemberObjectId            = ''
                        MemberObjectType          = ''
                        MemberUserPrincipalName   = ''
                        MemberDisplayName         = ''
                        MemberUserType            = ''
                        MemberAccountEnabled      = ''
                        Role                      = ''
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($teamMailNickname) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $teamMailNickname -Action 'GetMicrosoftTeamMember' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    TeamGroupId               = ''
                    TeamDisplayName           = ''
                    TeamMailNickname          = $teamMailNickname
                    MemberObjectId            = ''
                    MemberObjectType          = ''
                    MemberUserPrincipalName   = ''
                    MemberDisplayName         = ''
                    MemberUserType            = ''
                    MemberAccountEnabled      = ''
                    Role                      = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Microsoft Teams member inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}












