<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-163500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Groups
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets EntraMicrosoft365GroupMembers and exports results to CSV.

.DESCRIPTION
    Gets EntraMicrosoft365GroupMembers from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-IR3010-Get-EntraMicrosoft365GroupMembers.ps1 -InputCsvPath .\3010.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3010-Get-EntraMicrosoft365GroupMembers.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users
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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-MEID-0110-Get-EntraMicrosoft365GroupMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    return (($groupTypes -contains 'Unified') -and ($Group.MailEnabled -eq $true) -and ($Group.SecurityEnabled -eq $false))
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
            }
        }
    }

    return ''
}

$requiredHeaders = @(
    'GroupMailNickname'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'GroupDisplayName',
    'GroupMailNickname',
    'GroupMail',
    'GroupId',
    'MemberDisplayName',
    'MemberObjectType',
    'MemberObjectId',
    'MemberUserPrincipalName',
    'MemberUserType',
    'MemberAccountEnabled'
)

Write-Status -Message 'Starting Entra ID Microsoft 365 group member inventory script.'
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

$groupSelect = 'id,displayName,mail,mailNickname,groupTypes,mailEnabled,securityEnabled'
$allM365GroupsCache = $null
$userById = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

$rowNumber = 1
foreach ($row in $rows) {
    $groupMailNickname = Get-TrimmedValue -Value $row.GroupMailNickname

    try {
        if ([string]::IsNullOrWhiteSpace($groupMailNickname)) {
            throw 'GroupMailNickname is required. Use * to inventory members for all Microsoft 365 groups.'
        }

        $groups = @()
        if ($groupMailNickname -eq '*') {
            if ($null -eq $allM365GroupsCache) {
                $allGroups = @(Invoke-WithRetry -OperationName 'Load all groups for Microsoft 365 group member inventory' -ScriptBlock {
                    Get-MgGroup -All -Property $groupSelect -ErrorAction Stop
                })
                $allM365GroupsCache = @($allGroups | Where-Object { Test-IsMicrosoft365Group -Group $_ } | Sort-Object -Property MailNickname, DisplayName, Id)
            }

            $groups = @($allM365GroupsCache)
        }
        else {
            $escapedAlias = Escape-ODataString -Value $groupMailNickname
            $candidateGroups = @(Invoke-WithRetry -OperationName "Lookup group alias $groupMailNickname" -ScriptBlock {
                Get-MgGroup -Filter "mailNickname eq '$escapedAlias'" -ConsistencyLevel eventual -Property $groupSelect -ErrorAction Stop
            })

            $groups = @($candidateGroups | Where-Object { Test-IsMicrosoft365Group -Group $_ })
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'GetEntraMicrosoft365GroupMember' -Status 'NotFound' -Message 'No matching Microsoft 365 groups were found.' -Data ([ordered]@{
                            GroupDisplayName         = ''
                            GroupMailNickname        = $groupMailNickname
                            GroupMail                = ''
                            GroupId                  = ''
                            MemberDisplayName        = ''
                            MemberObjectType         = ''
                            MemberObjectId           = ''
                            MemberUserPrincipalName  = ''
                            MemberUserType           = ''
                            MemberAccountEnabled     = ''
                        })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property MailNickname, DisplayName, Id)) {
            $groupId = Get-TrimmedValue -Value $group.Id
            $groupDisplayName = Get-TrimmedValue -Value $group.DisplayName
            $resolvedAlias = Get-TrimmedValue -Value $group.MailNickname
            $groupMail = Get-TrimmedValue -Value $group.Mail

            $members = @(Invoke-WithRetry -OperationName "Load members for group $resolvedAlias" -ScriptBlock {
                Get-MgGroupMember -GroupId $groupId -All -ErrorAction Stop
            })

            if ($members.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$resolvedAlias|$groupId" -Action 'GetEntraMicrosoft365GroupMember' -Status 'Completed' -Message 'Microsoft 365 group has no members.' -Data ([ordered]@{
                                GroupDisplayName         = $groupDisplayName
                                GroupMailNickname        = $resolvedAlias
                                GroupMail                = $groupMail
                                GroupId                  = $groupId
                                MemberDisplayName        = ''
                                MemberObjectType         = ''
                                MemberObjectId           = ''
                                MemberUserPrincipalName  = ''
                                MemberUserType           = ''
                                MemberAccountEnabled     = ''
                            })))
                continue
            }

            foreach ($member in @($members | Sort-Object -Property Id)) {
                $memberId = Get-TrimmedValue -Value $member.Id
                $memberType = Get-DirectoryObjectType -DirectoryObject $member
                $memberUpn = Get-AdditionalPropertyValue -DirectoryObject $member -PropertyName 'userPrincipalName'
                $memberDisplayName = Get-AdditionalPropertyValue -DirectoryObject $member -PropertyName 'displayName'
                $memberUserType = ''
                $memberAccountEnabled = ''
                $message = 'Microsoft 365 group member exported.'

                if ($memberType -eq 'user' -and -not [string]::IsNullOrWhiteSpace($memberId)) {
                    try {
                        $memberUser = $null
                        if ($userById.ContainsKey($memberId)) {
                            $memberUser = $userById[$memberId]
                        }
                        else {
                            $memberUser = Invoke-WithRetry -OperationName "Load member user details $memberId" -ScriptBlock {
                                Get-MgUser -UserId $memberId -Property 'id,userPrincipalName,displayName,userType,accountEnabled' -ErrorAction Stop
                            }
                            $userById[$memberId] = $memberUser
                        }

                        $memberUpn = Get-TrimmedValue -Value $memberUser.UserPrincipalName
                        $memberDisplayName = Get-TrimmedValue -Value $memberUser.DisplayName
                        $memberUserType = Get-TrimmedValue -Value $memberUser.UserType
                        $memberAccountEnabled = [string]$memberUser.AccountEnabled
                    }
                    catch {
                        $message = "Microsoft 365 group member exported. User detail lookup failed: $($_.Exception.Message)"
                    }
                }

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$resolvedAlias|$memberId" -Action 'GetEntraMicrosoft365GroupMember' -Status 'Completed' -Message $message -Data ([ordered]@{
                                GroupDisplayName         = $groupDisplayName
                                GroupMailNickname        = $resolvedAlias
                                GroupMail                = $groupMail
                                GroupId                  = $groupId
                                MemberDisplayName        = $memberDisplayName
                                MemberObjectType         = $memberType
                                MemberObjectId           = $memberId
                                MemberUserPrincipalName  = $memberUpn
                                MemberUserType           = $memberUserType
                                MemberAccountEnabled     = $memberAccountEnabled
                            })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupMailNickname) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'GetEntraMicrosoft365GroupMember' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                        GroupDisplayName         = ''
                        GroupMailNickname        = $groupMailNickname
                        GroupMail                = ''
                        GroupId                  = ''
                        MemberDisplayName        = ''
                        MemberObjectType         = ''
                        MemberObjectId           = ''
                        MemberUserPrincipalName  = ''
                        MemberUserType           = ''
                        MemberAccountEnabled     = ''
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
Write-Status -Message 'Entra ID Microsoft 365 group member inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
