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
Microsoft.Graph.Teams
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets MicrosoftTeamsByUser and exports results to CSV.

.DESCRIPTION
    Gets MicrosoftTeamsByUser from Microsoft 365 and writes the results to a CSV file.
    For each in-scope user, discovers every Microsoft Team they belong to, with explicit
    role attribution (Member or Owner) and the access type (Direct or GuestInvite).
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
    .\D-TEAM-0500-Get-MicrosoftTeamsByUser.ps1 -InputCsvPath .\Scope-Users.input.csv
    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\D-TEAM-0500-Get-MicrosoftTeamsByUser.ps1 -DiscoverAll
    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Teams, Microsoft.Graph.Users
    Required roles:   Teams Administrator or Global Reader
    Limitations:      None known.

    CSV Fields:
    Column              Type    Required  Description
    ------              ----    --------  -----------
    UserPrincipalName   String  Yes       UPN of the user whose Teams memberships to inventory
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-TEAM-0500-Get-MicrosoftTeamsByUser_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$requiredHeaders = @('UserPrincipalName')

$reportPropertyOrder = @(
    'TimestampUtc', 'RowNumber', 'PrimaryKey', 'Action', 'Status', 'Message', 'ScopeMode',
    'UserPrincipalName', 'TeamId', 'TeamDisplayName', 'TeamVisibility',
    'TeamRole', 'AccessType', 'TeamArchived', 'LinkedM365GroupId'
)

Write-Status -Message 'Starting Microsoft Teams membership by user inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Teams', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('Team.ReadBasic.All', 'TeamMember.Read.All', 'GroupMember.Read.All')

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

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $upn = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw 'UserPrincipalName is required and cannot be blank.'
        }

        # Resolve user object
        $userObj = Invoke-WithRetry -OperationName "Resolve user $upn" -ScriptBlock {
            Get-MgUser -UserId $upn -Property 'Id,DisplayName,UserPrincipalName,UserType' -ErrorAction Stop
        }
        $userId   = ([string]$userObj.Id).Trim()
        $userType = ([string]$userObj.UserType).Trim()

        # Get all joined teams for the user
        $joinedTeams = @(Invoke-WithRetry -OperationName "Get joined teams for $upn" -ScriptBlock {
            Get-MgUserJoinedTeam -UserId $userId -All -Property 'Id,DisplayName,Description,Visibility,IsArchived' -ErrorAction Stop
        })

        if ($joinedTeams.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $upn -Action 'GetMicrosoftTeamsByUser' -Status 'Success' -Message 'No Teams memberships found.' -Data ([ordered]@{
                UserPrincipalName = $upn
                TeamId            = ''
                TeamDisplayName   = ''
                TeamVisibility    = ''
                TeamRole          = ''
                AccessType        = ''
                TeamArchived      = ''
                LinkedM365GroupId = ''
            })))
            $rowNumber++
            continue
        }

        # Determine access type based on user type
        $accessType = if ($userType -eq 'Guest') { 'GuestInvite' } else { 'Direct' }

        foreach ($team in @($joinedTeams | Sort-Object -Property DisplayName, Id)) {
            $teamId          = ([string]$team.Id).Trim()
            $teamDisplayName = ([string]$team.DisplayName).Trim()
            $teamVisibility  = ([string]$team.Visibility).Trim()
            $teamArchived    = if ($team.IsArchived -eq $true) { 'True' } else { 'False' }

            # Teams are built on M365 groups — TeamId == LinkedM365GroupId
            $linkedM365GroupId = $teamId

            # Determine role: check if user is an owner of the underlying M365 group
            $teamRole = 'Member'
            try {
                $owners = @(Invoke-WithRetry -OperationName "Get owners of team $teamId" -ScriptBlock {
                    Get-MgGroupOwner -GroupId $teamId -All -Property 'Id' -ErrorAction Stop
                })
                $ownerIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($owner in $owners) {
                    [void]$ownerIds.Add(([string]$owner.Id).Trim())
                }
                if ($ownerIds.Contains($userId)) {
                    $teamRole = 'Owner'
                }
            }
            catch {
                Write-Status -Message "Could not retrieve owners for team $teamId ($teamDisplayName): $($_.Exception.Message)" -Level WARN
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$upn|$teamId" -Action 'GetMicrosoftTeamsByUser' -Status 'Success' -Message 'Team membership exported.' -Data ([ordered]@{
                UserPrincipalName = $upn
                TeamId            = $teamId
                TeamDisplayName   = $teamDisplayName
                TeamVisibility    = $teamVisibility
                TeamRole          = $teamRole
                AccessType        = $accessType
                TeamArchived      = $teamArchived
                LinkedM365GroupId = $linkedM365GroupId
            })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $upn -Action 'GetMicrosoftTeamsByUser' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            UserPrincipalName = $upn
            TeamId            = ''
            TeamDisplayName   = ''
            TeamVisibility    = ''
            TeamRole          = ''
            AccessType        = ''
            TeamArchived      = ''
            LinkedM365GroupId = ''
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
Write-Status -Message 'Script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
