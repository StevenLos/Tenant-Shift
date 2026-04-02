<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-163000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Groups
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets EntraMicrosoft365GroupOwners and exports results to CSV.

.DESCRIPTION
    Gets EntraMicrosoft365GroupOwners from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-IR3009-Get-EntraMicrosoft365GroupOwners.ps1 -InputCsvPath .\3009.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3009-Get-EntraMicrosoft365GroupOwners.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-MEID-0120-Get-EntraMicrosoft365GroupOwners_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'OwnerDisplayName',
    'OwnerObjectType',
    'OwnerObjectId',
    'OwnerUserPrincipalName',
    'OwnerUserType',
    'OwnerAccountEnabled'
)

Write-Status -Message 'Starting Entra ID Microsoft 365 group owner inventory script.'
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
            throw 'GroupMailNickname is required. Use * to inventory owners for all Microsoft 365 groups.'
        }

        $groups = @()
        if ($groupMailNickname -eq '*') {
            if ($null -eq $allM365GroupsCache) {
                $allGroups = @(Invoke-WithRetry -OperationName 'Load all groups for Microsoft 365 group owner inventory' -ScriptBlock {
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
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'GetEntraMicrosoft365GroupOwner' -Status 'NotFound' -Message 'No matching Microsoft 365 groups were found.' -Data ([ordered]@{
                            GroupDisplayName         = ''
                            GroupMailNickname        = $groupMailNickname
                            GroupMail                = ''
                            GroupId                  = ''
                            OwnerDisplayName         = ''
                            OwnerObjectType          = ''
                            OwnerObjectId            = ''
                            OwnerUserPrincipalName   = ''
                            OwnerUserType            = ''
                            OwnerAccountEnabled      = ''
                        })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property MailNickname, DisplayName, Id)) {
            $groupId = Get-TrimmedValue -Value $group.Id
            $groupDisplayName = Get-TrimmedValue -Value $group.DisplayName
            $resolvedAlias = Get-TrimmedValue -Value $group.MailNickname
            $groupMail = Get-TrimmedValue -Value $group.Mail

            $owners = @(Invoke-WithRetry -OperationName "Load owners for group $resolvedAlias" -ScriptBlock {
                Get-MgGroupOwner -GroupId $groupId -All -ErrorAction Stop
            })

            if ($owners.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$resolvedAlias|$groupId" -Action 'GetEntraMicrosoft365GroupOwner' -Status 'Completed' -Message 'Microsoft 365 group has no owners.' -Data ([ordered]@{
                                GroupDisplayName         = $groupDisplayName
                                GroupMailNickname        = $resolvedAlias
                                GroupMail                = $groupMail
                                GroupId                  = $groupId
                                OwnerDisplayName         = ''
                                OwnerObjectType          = ''
                                OwnerObjectId            = ''
                                OwnerUserPrincipalName   = ''
                                OwnerUserType            = ''
                                OwnerAccountEnabled      = ''
                            })))
                continue
            }

            foreach ($owner in @($owners | Sort-Object -Property Id)) {
                $ownerId = Get-TrimmedValue -Value $owner.Id
                $ownerType = Get-DirectoryObjectType -DirectoryObject $owner
                $ownerUpn = Get-AdditionalPropertyValue -DirectoryObject $owner -PropertyName 'userPrincipalName'
                $ownerDisplayName = Get-AdditionalPropertyValue -DirectoryObject $owner -PropertyName 'displayName'
                $ownerUserType = ''
                $ownerAccountEnabled = ''
                $message = 'Microsoft 365 group owner exported.'

                if ($ownerType -eq 'user' -and -not [string]::IsNullOrWhiteSpace($ownerId)) {
                    try {
                        $ownerUser = $null
                        if ($userById.ContainsKey($ownerId)) {
                            $ownerUser = $userById[$ownerId]
                        }
                        else {
                            $ownerUser = Invoke-WithRetry -OperationName "Load owner user details $ownerId" -ScriptBlock {
                                Get-MgUser -UserId $ownerId -Property 'id,userPrincipalName,displayName,userType,accountEnabled' -ErrorAction Stop
                            }
                            $userById[$ownerId] = $ownerUser
                        }

                        $ownerUpn = Get-TrimmedValue -Value $ownerUser.UserPrincipalName
                        $ownerDisplayName = Get-TrimmedValue -Value $ownerUser.DisplayName
                        $ownerUserType = Get-TrimmedValue -Value $ownerUser.UserType
                        $ownerAccountEnabled = [string]$ownerUser.AccountEnabled
                    }
                    catch {
                        $message = "Microsoft 365 group owner exported. User detail lookup failed: $($_.Exception.Message)"
                    }
                }

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$resolvedAlias|$ownerId" -Action 'GetEntraMicrosoft365GroupOwner' -Status 'Completed' -Message $message -Data ([ordered]@{
                                GroupDisplayName         = $groupDisplayName
                                GroupMailNickname        = $resolvedAlias
                                GroupMail                = $groupMail
                                GroupId                  = $groupId
                                OwnerDisplayName         = $ownerDisplayName
                                OwnerObjectType          = $ownerType
                                OwnerObjectId            = $ownerId
                                OwnerUserPrincipalName   = $ownerUpn
                                OwnerUserType            = $ownerUserType
                                OwnerAccountEnabled      = $ownerAccountEnabled
                            })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupMailNickname) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'GetEntraMicrosoft365GroupOwner' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                        GroupDisplayName         = ''
                        GroupMailNickname        = $groupMailNickname
                        GroupMail                = ''
                        GroupId                  = ''
                        OwnerDisplayName         = ''
                        OwnerObjectType          = ''
                        OwnerObjectId            = ''
                        OwnerUserPrincipalName   = ''
                        OwnerUserType            = ''
                        OwnerAccountEnabled      = ''
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
Write-Status -Message 'Entra ID Microsoft 365 group owner inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
