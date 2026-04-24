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
Microsoft.Graph.Groups
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets EntraDynamicGroupsByUser and exports results to CSV.

.DESCRIPTION
    Gets EntraDynamicGroupsByUser from Microsoft 365 and writes the results to a CSV file.
    For each in-scope user, finds all dynamic Entra ID security groups where the user is an
    evaluated member. Results are based on actual evaluated membership via transitiveMemberOf,
    not rule text inspection.
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
    .\D-MEID-0510-Get-EntraDynamicGroupsByUser.ps1 -InputCsvPath .\Scope-Users.input.csv
    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\D-MEID-0510-Get-EntraDynamicGroupsByUser.ps1 -DiscoverAll
    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users
    Required roles:   Global Reader
    Limitations:      None known. Membership results reflect evaluated state at time of query.

    CSV Fields:
    Column              Type    Required  Description
    ------              ----    --------  -----------
    UserPrincipalName   String  Yes       UPN of the user whose dynamic group memberships to inventory
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-MEID-0510-Get-EntraDynamicGroupsByUser_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Test-IsDynamicSecurityGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Group
    )

    $groupTypes = @()
    if ($null -ne $Group.GroupTypes) {
        $groupTypes = @($Group.GroupTypes)
    }

    $isDynamic = $groupTypes -contains 'DynamicMembership'
    $isUnified = $groupTypes -contains 'Unified'
    $result    = ($Group.SecurityEnabled -eq $true -and $isDynamic -and -not $isUnified)
    return $result
}

$requiredHeaders = @('UserPrincipalName')

$reportPropertyOrder = @(
    'TimestampUtc', 'RowNumber', 'PrimaryKey', 'Action', 'Status', 'Message', 'ScopeMode',
    'UserPrincipalName', 'GroupId', 'GroupDisplayName',
    'MembershipType', 'InsertionMethod', 'DynamicMembershipRule', 'RuleProcessingState'
)

Write-Status -Message 'Starting Entra ID dynamic security group membership by user inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('GroupMember.Read.All', 'Directory.Read.All')

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

$results          = [System.Collections.Generic.List[object]]::new()
$groupDetailCache = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$rowNumber        = 1

foreach ($row in $rows) {
    $upn = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw 'UserPrincipalName is required and cannot be blank.'
        }

        # Resolve user object
        $userObj = Invoke-WithRetry -OperationName "Resolve user $upn" -ScriptBlock {
            Get-MgUser -UserId $upn -Property 'Id,DisplayName,UserPrincipalName' -ErrorAction Stop
        }
        $userId = ([string]$userObj.Id).Trim()

        # Get all transitive memberships — presence here IS evaluated membership confirmation
        $transitiveMemberships = @(Invoke-WithRetry -OperationName "Get transitive memberships for $upn" -ScriptBlock {
            Get-MgUserTransitiveMemberOf -UserId $userId -All `
                -Property 'Id,DisplayName,GroupTypes,SecurityEnabled,MailEnabled,MembershipRule,MembershipRuleProcessingState' `
                -ErrorAction Stop
        })

        # Filter to dynamic security groups only
        $dynamicGroups = @($transitiveMemberships | Where-Object { Test-IsDynamicSecurityGroup -Group $_ })

        if ($dynamicGroups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $upn -Action 'GetEntraDynamicGroupsByUser' -Status 'Success' -Message 'No memberships found.' -Data ([ordered]@{
                UserPrincipalName     = $upn
                GroupId               = ''
                GroupDisplayName      = ''
                MembershipType        = ''
                InsertionMethod       = ''
                DynamicMembershipRule = ''
                RuleProcessingState   = ''
            })))
            $rowNumber++
            continue
        }

        foreach ($group in @($dynamicGroups | Sort-Object -Property DisplayName, Id)) {
            $groupId          = ([string]$group.Id).Trim()
            $groupDisplayName = ([string]$group.DisplayName).Trim()

            # Fetch full group object to get MembershipRule and RuleProcessingState
            if (-not $groupDetailCache.ContainsKey($groupId)) {
                $fullGroup = Invoke-WithRetry -OperationName "Get full group details for $groupId" -ScriptBlock {
                    Get-MgGroup -GroupId $groupId -Property 'Id,DisplayName,MembershipRule,MembershipRuleProcessingState' -ErrorAction Stop
                }
                $groupDetailCache[$groupId] = $fullGroup
            }

            $groupDetail          = $groupDetailCache[$groupId]
            $membershipRule       = ([string]$groupDetail.MembershipRule).Trim()
            $ruleProcessingState  = ([string]$groupDetail.MembershipRuleProcessingState).Trim()

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$upn|$groupId" -Action 'GetEntraDynamicGroupsByUser' -Status 'Success' -Message 'Dynamic group membership exported.' -Data ([ordered]@{
                UserPrincipalName     = $upn
                GroupId               = $groupId
                GroupDisplayName      = $groupDisplayName
                MembershipType        = 'Dynamic'
                InsertionMethod       = 'Dynamic'
                DynamicMembershipRule = $membershipRule
                RuleProcessingState   = $ruleProcessingState
            })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $upn -Action 'GetEntraDynamicGroupsByUser' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            UserPrincipalName     = $upn
            GroupId               = ''
            GroupDisplayName      = ''
            MembershipType        = ''
            InsertionMethod       = ''
            DynamicMembershipRule = ''
            RuleProcessingState   = ''
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
Write-Status -Message 'Entra ID dynamic security group membership by user inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
