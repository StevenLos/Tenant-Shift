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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR3007-Get-EntraSecurityGroupMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Test-IsAssignedSecurityGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Group
    )

    $membershipRule = ([string]$Group.MembershipRule).Trim()
    $isAssigned = ($Group.SecurityEnabled -eq $true -and $Group.MailEnabled -eq $false -and [string]::IsNullOrWhiteSpace($membershipRule))
    return $isAssigned
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
    'GroupDisplayName'
)

Write-Status -Message 'Starting Entra ID security group membership inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('Group.Read.All', 'User.Read.All', 'Directory.Read.All')

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

$groupSelect = 'id,displayName,mailNickname,securityEnabled,mailEnabled,membershipRule'
$allAssignedGroupsCache = $null
$userById = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

$rowNumber = 1
foreach ($row in $rows) {
    $groupDisplayName = ([string]$row.GroupDisplayName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($groupDisplayName)) {
            throw 'GroupDisplayName is required. Use * to inventory all assigned security group memberships.'
        }

        $groups = @()
        if ($groupDisplayName -eq '*') {
            if ($null -eq $allAssignedGroupsCache) {
                $allGroups = @(Invoke-WithRetry -OperationName 'Load all groups for security group membership inventory' -ScriptBlock {
                    Get-MgGroup -All -Property $groupSelect -ErrorAction Stop
                })
                $allAssignedGroupsCache = @($allGroups | Where-Object { Test-IsAssignedSecurityGroup -Group $_ } | Sort-Object -Property DisplayName, Id)
            }

            $groups = @($allAssignedGroupsCache)
        }
        else {
            $escapedName = Escape-ODataString -Value $groupDisplayName
            $candidateGroups = @(Invoke-WithRetry -OperationName "Lookup group $groupDisplayName" -ScriptBlock {
                Get-MgGroup -Filter "displayName eq '$escapedName'" -ConsistencyLevel eventual -Property $groupSelect -ErrorAction Stop
            })

            $groups = @($candidateGroups | Where-Object { Test-IsAssignedSecurityGroup -Group $_ })
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'GetEntraSecurityGroupMember' -Status 'NotFound' -Message 'No matching assigned security groups were found.' -Data ([ordered]@{
                        GroupId                  = ''
                        GroupDisplayName         = $groupDisplayName
                        MemberObjectId           = ''
                        MemberObjectType         = ''
                        MemberUserPrincipalName  = ''
                        MemberDisplayName        = ''
                        MemberUserType           = ''
                        MemberAccountEnabled     = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Id)) {
            $groupId = ([string]$group.Id).Trim()
            $resolvedGroupName = ([string]$group.DisplayName).Trim()
            $members = @(Invoke-WithRetry -OperationName "Load members for group $resolvedGroupName" -ScriptBlock {
                Get-MgGroupMember -GroupId $groupId -All -ErrorAction Stop
            })

            if ($members.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$resolvedGroupName|$groupId" -Action 'GetEntraSecurityGroupMember' -Status 'Completed' -Message 'Group has no members.' -Data ([ordered]@{
                            GroupId                  = $groupId
                            GroupDisplayName         = $resolvedGroupName
                            MemberObjectId           = ''
                            MemberObjectType         = ''
                            MemberUserPrincipalName  = ''
                            MemberDisplayName        = ''
                            MemberUserType           = ''
                            MemberAccountEnabled     = ''
                        })))
                continue
            }

            foreach ($member in @($members | Sort-Object -Property Id)) {
                $memberId = ([string]$member.Id).Trim()
                $memberType = Get-DirectoryObjectType -DirectoryObject $member
                $memberUpn = ''
                $memberDisplayName = ''
                $memberUserType = ''
                $memberAccountEnabled = ''
                $message = 'Group member exported.'

                if ($memberType -eq 'user' -and -not [string]::IsNullOrWhiteSpace($memberId)) {
                    try {
                        $user = $null
                        if ($userById.ContainsKey($memberId)) {
                            $user = $userById[$memberId]
                        }
                        else {
                            $user = Invoke-WithRetry -OperationName "Load user details for member $memberId" -ScriptBlock {
                                Get-MgUser -UserId $memberId -Property 'id,userPrincipalName,displayName,userType,accountEnabled' -ErrorAction Stop
                            }
                            $userById[$memberId] = $user
                        }

                        $memberUpn = ([string]$user.UserPrincipalName).Trim()
                        $memberDisplayName = ([string]$user.DisplayName).Trim()
                        $memberUserType = ([string]$user.UserType).Trim()
                        $memberAccountEnabled = [string]$user.AccountEnabled
                    }
                    catch {
                        $memberUpn = Get-AdditionalPropertyValue -DirectoryObject $member -PropertyName 'userPrincipalName'
                        $memberDisplayName = Get-AdditionalPropertyValue -DirectoryObject $member -PropertyName 'displayName'
                        $message = "Group member exported. User detail lookup failed: $($_.Exception.Message)"
                    }
                }
                else {
                    $memberUpn = Get-AdditionalPropertyValue -DirectoryObject $member -PropertyName 'userPrincipalName'
                    $memberDisplayName = Get-AdditionalPropertyValue -DirectoryObject $member -PropertyName 'displayName'
                }

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$resolvedGroupName|$memberId" -Action 'GetEntraSecurityGroupMember' -Status 'Completed' -Message $message -Data ([ordered]@{
                            GroupId                  = $groupId
                            GroupDisplayName         = $resolvedGroupName
                            MemberObjectId           = $memberId
                            MemberObjectType         = $memberType
                            MemberUserPrincipalName  = $memberUpn
                            MemberDisplayName        = $memberDisplayName
                            MemberUserType           = $memberUserType
                            MemberAccountEnabled     = $memberAccountEnabled
                        })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupDisplayName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'GetEntraSecurityGroupMember' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    GroupId                  = ''
                    GroupDisplayName         = $groupDisplayName
                    MemberObjectId           = ''
                    MemberObjectType         = ''
                    MemberUserPrincipalName  = ''
                    MemberDisplayName        = ''
                    MemberUserType           = ''
                    MemberAccountEnabled     = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID security group membership inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}












