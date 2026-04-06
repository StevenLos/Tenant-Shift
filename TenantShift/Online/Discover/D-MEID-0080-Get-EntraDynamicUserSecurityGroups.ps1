<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-161000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Groups

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets EntraDynamicUserSecurityGroups and exports results to CSV.

.DESCRIPTION
    Gets EntraDynamicUserSecurityGroups from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-D3006-Get-EntraDynamicUserSecurityGroups.ps1 -InputCsvPath .\3006.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D3006-Get-EntraDynamicUserSecurityGroups.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Groups
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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-MEID-0080-Get-EntraDynamicUserSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Test-IsDynamicUserSecurityGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Group
    )

    $membershipRule = Get-TrimmedValue -Value $Group.MembershipRule
    $groupTypes = @($Group.GroupTypes)

    $isDynamic = ($groupTypes -contains 'DynamicMembership')
    $isSecurityEnabled = ($Group.SecurityEnabled -eq $true)
    $isMailDisabled = ($Group.MailEnabled -eq $false)
    $hasRule = -not [string]::IsNullOrWhiteSpace($membershipRule)
    $looksLikeUserRule = $membershipRule.ToLowerInvariant().Contains('user.')

    return ($isDynamic -and $isSecurityEnabled -and $isMailDisabled -and $hasRule -and $looksLikeUserRule)
}

$requiredHeaders = @(
    'GroupDisplayName'
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
    'GroupId',
    'Description',
    'Mail',
    'MailNickname',
    'ProxyAddresses',
    'SecurityEnabled',
    'MailEnabled',
    'MembershipType',
    'MembershipRule',
    'MembershipRuleProcessingState',
    'GroupTypes',
    'Visibility',
    'Classification',
    'PreferredDataLocation',
    'IsAssignableToRole',
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

Write-Status -Message 'Starting Entra ID dynamic user security group inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')
Ensure-GraphConnection -RequiredScopes @('Group.Read.All', 'Directory.Read.All')

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

$groupSelect = 'id,displayName,description,mail,mailNickname,proxyAddresses,securityEnabled,mailEnabled,membershipRule,membershipRuleProcessingState,groupTypes,visibility,classification,preferredDataLocation,isAssignableToRole,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesDomainName,onPremisesNetBiosName,onPremisesSamAccountName,onPremisesSecurityIdentifier,createdDateTime,renewedDateTime,deletedDateTime'
$allDynamicUserGroupsCache = $null

$rowNumber = 1
foreach ($row in $rows) {
    $groupDisplayName = ([string]$row.GroupDisplayName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($groupDisplayName)) {
            throw 'GroupDisplayName is required. Use * to inventory all dynamic user security groups.'
        }

        $groups = @()
        if ($groupDisplayName -eq '*') {
            if ($null -eq $allDynamicUserGroupsCache) {
                $allGroups = @(Invoke-WithRetry -OperationName 'Load all groups for dynamic user security group inventory' -ScriptBlock {
                    Get-MgGroup -All -Property $groupSelect -ErrorAction Stop
                })
                $allDynamicUserGroupsCache = @($allGroups | Where-Object { Test-IsDynamicUserSecurityGroup -Group $_ } | Sort-Object -Property DisplayName, Id)
            }

            $groups = @($allDynamicUserGroupsCache)
        }
        else {
            $escapedName = Escape-ODataString -Value $groupDisplayName
            $candidateGroups = @(Invoke-WithRetry -OperationName "Lookup group $groupDisplayName" -ScriptBlock {
                Get-MgGroup -Filter "displayName eq '$escapedName'" -ConsistencyLevel eventual -Property $groupSelect -ErrorAction Stop
            })

            $groups = @($candidateGroups | Where-Object { Test-IsDynamicUserSecurityGroup -Group $_ })
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'GetEntraDynamicUserSecurityGroup' -Status 'NotFound' -Message 'No matching dynamic user security groups were found.' -Data ([ordered]@{
                        GroupId                       = ''
                        GroupDisplayName              = $groupDisplayName
                        Description                   = ''
                        Mail                          = ''
                        MailNickname                  = ''
                        ProxyAddresses                = ''
                        SecurityEnabled               = ''
                        MailEnabled                   = ''
                        MembershipType                = ''
                        MembershipRule                = ''
                        MembershipRuleProcessingState = ''
                        GroupTypes                    = ''
                        Visibility                    = ''
                        Classification                = ''
                        PreferredDataLocation         = ''
                        IsAssignableToRole            = ''
                        OnPremisesSyncEnabled         = ''
                        OnPremisesLastSyncDateTime    = ''
                        OnPremisesDomainName          = ''
                        OnPremisesNetBiosName         = ''
                        OnPremisesSamAccountName      = ''
                        OnPremisesSecurityIdentifier  = ''
                        CreatedDateTime               = ''
                        RenewedDateTime               = ''
                        DeletedDateTime               = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Id)) {
            $groupId = Get-TrimmedValue -Value $group.Id
            $displayName = Get-TrimmedValue -Value $group.DisplayName
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$displayName|$groupId" -Action 'GetEntraDynamicUserSecurityGroup' -Status 'Completed' -Message 'Dynamic user security group exported.' -Data ([ordered]@{
                        GroupId                       = $groupId
                        GroupDisplayName              = $displayName
                        Description                   = Get-TrimmedValue -Value $group.Description
                        Mail                          = Get-TrimmedValue -Value $group.Mail
                        MailNickname                  = Get-TrimmedValue -Value $group.MailNickname
                        ProxyAddresses                = Convert-MultiValueToString -Value $group.ProxyAddresses
                        SecurityEnabled               = [string]$group.SecurityEnabled
                        MailEnabled                   = [string]$group.MailEnabled
                        MembershipType                = 'Dynamic'
                        MembershipRule                = Get-TrimmedValue -Value $group.MembershipRule
                        MembershipRuleProcessingState = Get-TrimmedValue -Value $group.MembershipRuleProcessingState
                        GroupTypes                    = Convert-MultiValueToString -Value $group.GroupTypes
                        Visibility                    = Get-TrimmedValue -Value $group.Visibility
                        Classification                = Get-TrimmedValue -Value $group.Classification
                        PreferredDataLocation         = Get-TrimmedValue -Value $group.PreferredDataLocation
                        IsAssignableToRole            = [string]$group.IsAssignableToRole
                        OnPremisesSyncEnabled         = [string]$group.OnPremisesSyncEnabled
                        OnPremisesLastSyncDateTime    = [string]$group.OnPremisesLastSyncDateTime
                        OnPremisesDomainName          = Get-TrimmedValue -Value $group.OnPremisesDomainName
                        OnPremisesNetBiosName         = Get-TrimmedValue -Value $group.OnPremisesNetBiosName
                        OnPremisesSamAccountName      = Get-TrimmedValue -Value $group.OnPremisesSamAccountName
                        OnPremisesSecurityIdentifier  = Get-TrimmedValue -Value $group.OnPremisesSecurityIdentifier
                        CreatedDateTime               = [string]$group.CreatedDateTime
                        RenewedDateTime               = [string]$group.RenewedDateTime
                        DeletedDateTime               = [string]$group.DeletedDateTime
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupDisplayName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'GetEntraDynamicUserSecurityGroup' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    GroupId                       = ''
                    GroupDisplayName              = $groupDisplayName
                    Description                   = ''
                    Mail                          = ''
                    MailNickname                  = ''
                    ProxyAddresses                = ''
                    SecurityEnabled               = ''
                    MailEnabled                   = ''
                    MembershipType                = ''
                    MembershipRule                = ''
                    MembershipRuleProcessingState = ''
                    GroupTypes                    = ''
                    Visibility                    = ''
                    Classification                = ''
                    PreferredDataLocation         = ''
                    IsAssignableToRole            = ''
                    OnPremisesSyncEnabled         = ''
                    OnPremisesLastSyncDateTime    = ''
                    OnPremisesDomainName          = ''
                    OnPremisesNetBiosName         = ''
                    OnPremisesSamAccountName      = ''
                    OnPremisesSecurityIdentifier  = ''
                    CreatedDateTime               = ''
                    RenewedDateTime               = ''
                    DeletedDateTime               = ''
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
Write-Status -Message 'Entra ID dynamic user security group inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}




