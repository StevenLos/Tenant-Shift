<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-162000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Groups
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets EntraMicrosoft365Groups and exports results to CSV.

.DESCRIPTION
    Gets EntraMicrosoft365Groups from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-D3008-Get-EntraMicrosoft365Groups.ps1 -InputCsvPath .\3008.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D3008-Get-EntraMicrosoft365Groups.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-MEID-0100-Get-EntraMicrosoft365Groups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'GroupId',
    'Description',
    'Mail',
    'ProxyAddresses',
    'Visibility',
    'Classification',
    'PreferredDataLocation',
    'GroupTypes',
    'MailEnabled',
    'SecurityEnabled',
    'IsAssignableToRole',
    'OwnersCount',
    'OwnersUserPrincipalNames',
    'OwnersObjectIds',
    'MembersCount',
    'MembersUserPrincipalNames',
    'MembersObjectIds',
    'OnPremisesSyncEnabled',
    'OnPremisesLastSyncDateTime',
    'OnPremisesDomainName',
    'OnPremisesNetBiosName',
    'OnPremisesSamAccountName',
    'OnPremisesSecurityIdentifier',
    'CreatedDateTime',
    'RenewedDateTime',
    'DeletedDateTime'
)

Write-Status -Message 'Starting Entra ID Microsoft 365 group inventory script.'
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

$groupSelect = 'id,displayName,description,mail,mailNickname,proxyAddresses,groupTypes,visibility,classification,preferredDataLocation,mailEnabled,securityEnabled,isAssignableToRole,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesDomainName,onPremisesNetBiosName,onPremisesSamAccountName,onPremisesSecurityIdentifier,createdDateTime,renewedDateTime,deletedDateTime'
$allM365GroupsCache = $null
$userById = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

$rowNumber = 1
foreach ($row in $rows) {
    $groupMailNickname = ([string]$row.GroupMailNickname).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($groupMailNickname)) {
            throw 'GroupMailNickname is required. Use * to inventory all Microsoft 365 groups.'
        }

        $groups = @()
        if ($groupMailNickname -eq '*') {
            if ($null -eq $allM365GroupsCache) {
                $allGroups = @(Invoke-WithRetry -OperationName 'Load all groups for Microsoft 365 group inventory' -ScriptBlock {
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
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'GetEntraMicrosoft365Group' -Status 'NotFound' -Message 'No matching Microsoft 365 groups were found.' -Data ([ordered]@{
                        GroupId                        = ''
                        GroupDisplayName               = ''
                        GroupMailNickname              = $groupMailNickname
                        Description                    = ''
                        Mail                           = ''
                        ProxyAddresses                 = ''
                        Visibility                     = ''
                        Classification                 = ''
                        PreferredDataLocation          = ''
                        GroupTypes                     = ''
                        MailEnabled                    = ''
                        SecurityEnabled                = ''
                        IsAssignableToRole             = ''
                        OnPremisesSyncEnabled          = ''
                        OnPremisesLastSyncDateTime     = ''
                        OnPremisesDomainName           = ''
                        OnPremisesNetBiosName          = ''
                        OnPremisesSamAccountName       = ''
                        OnPremisesSecurityIdentifier   = ''
                        CreatedDateTime                = ''
                        RenewedDateTime                = ''
                        DeletedDateTime                = ''
                        OwnersUserPrincipalNames       = ''
                        OwnersObjectIds                = ''
                        OwnersCount                    = ''
                        MembersUserPrincipalNames      = ''
                        MembersObjectIds               = ''
                        MembersCount                   = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property MailNickname, DisplayName, Id)) {
            $groupId = Get-TrimmedValue -Value $group.Id
            $groupDisplayName = Get-TrimmedValue -Value $group.DisplayName
            $resolvedAlias = Get-TrimmedValue -Value $group.MailNickname

            $ownerUpns = [System.Collections.Generic.List[string]]::new()
            $ownerObjectIds = [System.Collections.Generic.List[string]]::new()
            $memberUpns = [System.Collections.Generic.List[string]]::new()
            $memberObjectIds = [System.Collections.Generic.List[string]]::new()

            $owners = @(Invoke-WithRetry -OperationName "Load owners for group $resolvedAlias" -ScriptBlock {
                Get-MgGroupOwner -GroupId $groupId -All -ErrorAction Stop
            })

            foreach ($owner in @($owners | Sort-Object -Property Id)) {
                $ownerId = Get-TrimmedValue -Value $owner.Id
                if ([string]::IsNullOrWhiteSpace($ownerId)) {
                    continue
                }

                if (-not $ownerObjectIds.Contains($ownerId)) {
                    $ownerObjectIds.Add($ownerId)
                }

                if ((Get-DirectoryObjectType -DirectoryObject $owner) -ne 'user') {
                    continue
                }

                try {
                    $ownerUser = $null
                    if ($userById.ContainsKey($ownerId)) {
                        $ownerUser = $userById[$ownerId]
                    }
                    else {
                        $ownerUser = Invoke-WithRetry -OperationName "Load owner user details $ownerId" -ScriptBlock {
                            Get-MgUser -UserId $ownerId -Property 'id,userPrincipalName' -ErrorAction Stop
                        }
                        $userById[$ownerId] = $ownerUser
                    }

                    $ownerUpn = Get-TrimmedValue -Value $ownerUser.UserPrincipalName
                    if (-not [string]::IsNullOrWhiteSpace($ownerUpn) -and -not $ownerUpns.Contains($ownerUpn)) {
                        $ownerUpns.Add($ownerUpn)
                    }
                }
                catch {
                    Write-Status -Message "Owner detail lookup failed for owner ID '$ownerId' in group '$groupDisplayName': $($_.Exception.Message)" -Level WARN
                }
            }

            $members = @(Invoke-WithRetry -OperationName "Load members for group $resolvedAlias" -ScriptBlock {
                Get-MgGroupMember -GroupId $groupId -All -ErrorAction Stop
            })

            foreach ($member in @($members | Sort-Object -Property Id)) {
                $memberId = Get-TrimmedValue -Value $member.Id
                if ([string]::IsNullOrWhiteSpace($memberId)) {
                    continue
                }

                if (-not $memberObjectIds.Contains($memberId)) {
                    $memberObjectIds.Add($memberId)
                }

                if ((Get-DirectoryObjectType -DirectoryObject $member) -ne 'user') {
                    continue
                }

                try {
                    $memberUser = $null
                    if ($userById.ContainsKey($memberId)) {
                        $memberUser = $userById[$memberId]
                    }
                    else {
                        $memberUser = Invoke-WithRetry -OperationName "Load member user details $memberId" -ScriptBlock {
                            Get-MgUser -UserId $memberId -Property 'id,userPrincipalName' -ErrorAction Stop
                        }
                        $userById[$memberId] = $memberUser
                    }

                    $memberUpn = Get-TrimmedValue -Value $memberUser.UserPrincipalName
                    if (-not [string]::IsNullOrWhiteSpace($memberUpn) -and -not $memberUpns.Contains($memberUpn)) {
                        $memberUpns.Add($memberUpn)
                    }
                }
                catch {
                    Write-Status -Message "Member detail lookup failed for member ID '$memberId' in group '$groupDisplayName': $($_.Exception.Message)" -Level WARN
                }
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$resolvedAlias|$groupId" -Action 'GetEntraMicrosoft365Group' -Status 'Completed' -Message 'Microsoft 365 group exported.' -Data ([ordered]@{
                        GroupId                        = $groupId
                        GroupDisplayName               = $groupDisplayName
                        GroupMailNickname              = $resolvedAlias
                        Description                    = Get-TrimmedValue -Value $group.Description
                        Mail                           = Get-TrimmedValue -Value $group.Mail
                        ProxyAddresses                 = Convert-MultiValueToString -Value $group.ProxyAddresses
                        Visibility                     = Get-TrimmedValue -Value $group.Visibility
                        Classification                 = Get-TrimmedValue -Value $group.Classification
                        PreferredDataLocation          = Get-TrimmedValue -Value $group.PreferredDataLocation
                        GroupTypes                     = Convert-MultiValueToString -Value $group.GroupTypes
                        MailEnabled                    = [string]$group.MailEnabled
                        SecurityEnabled                = [string]$group.SecurityEnabled
                        IsAssignableToRole             = [string]$group.IsAssignableToRole
                        OnPremisesSyncEnabled          = [string]$group.OnPremisesSyncEnabled
                        OnPremisesLastSyncDateTime     = [string]$group.OnPremisesLastSyncDateTime
                        OnPremisesDomainName           = Get-TrimmedValue -Value $group.OnPremisesDomainName
                        OnPremisesNetBiosName          = Get-TrimmedValue -Value $group.OnPremisesNetBiosName
                        OnPremisesSamAccountName       = Get-TrimmedValue -Value $group.OnPremisesSamAccountName
                        OnPremisesSecurityIdentifier   = Get-TrimmedValue -Value $group.OnPremisesSecurityIdentifier
                        CreatedDateTime                = [string]$group.CreatedDateTime
                        RenewedDateTime                = [string]$group.RenewedDateTime
                        DeletedDateTime                = [string]$group.DeletedDateTime
                        OwnersUserPrincipalNames       = (@($ownerUpns | Sort-Object) -join ';')
                        OwnersObjectIds                = (@($ownerObjectIds | Sort-Object) -join ';')
                        OwnersCount                    = [string]$ownerObjectIds.Count
                        MembersUserPrincipalNames      = (@($memberUpns | Sort-Object) -join ';')
                        MembersObjectIds               = (@($memberObjectIds | Sort-Object) -join ';')
                        MembersCount                   = [string]$memberObjectIds.Count
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupMailNickname) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'GetEntraMicrosoft365Group' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    GroupId                        = ''
                    GroupDisplayName               = ''
                    GroupMailNickname              = $groupMailNickname
                    Description                    = ''
                    Mail                           = ''
                    ProxyAddresses                 = ''
                    Visibility                     = ''
                    Classification                 = ''
                    PreferredDataLocation          = ''
                    GroupTypes                     = ''
                    MailEnabled                    = ''
                    SecurityEnabled                = ''
                    IsAssignableToRole             = ''
                    OnPremisesSyncEnabled          = ''
                    OnPremisesLastSyncDateTime     = ''
                    OnPremisesDomainName           = ''
                    OnPremisesNetBiosName          = ''
                    OnPremisesSamAccountName       = ''
                    OnPremisesSecurityIdentifier   = ''
                    CreatedDateTime                = ''
                    RenewedDateTime                = ''
                    DeletedDateTime                = ''
                    OwnersUserPrincipalNames       = ''
                    OwnersObjectIds                = ''
                    OwnersCount                    = ''
                    MembersUserPrincipalNames      = ''
                    MembersObjectIds               = ''
                    MembersCount                   = ''
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
Write-Status -Message 'Entra ID Microsoft 365 group inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}




