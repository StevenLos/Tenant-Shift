<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-154500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets EntraGuestUsers and exports results to CSV.

.DESCRIPTION
    Gets EntraGuestUsers from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-D3002-Get-EntraGuestUsers.ps1 -InputCsvPath .\3002.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D3002-Get-EntraGuestUsers.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Users
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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-MEID-0020-Get-EntraGuestUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$requiredHeaders = @(
    'UserPrincipalName'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'UserId',
    'UserPrincipalName',
    'Mail',
    'DisplayName',
    'GivenName',
    'Surname',
    'UserType',
    'AccountEnabled',
    'ExternalUserState',
    'ExternalUserStateChangeDateTime',
    'CreationType',
    'CreatedDateTime'
)

Write-Status -Message 'Starting Entra ID guest user inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('User.Read.All', 'Directory.Read.All')

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

$userSelect = 'id,userPrincipalName,displayName,givenName,surname,mail,accountEnabled,userType,externalUserState,externalUserStateChangeDateTime,createdDateTime,creationType'

$rowNumber = 1
foreach ($row in $rows) {
    $userPrincipalName = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
            throw 'UserPrincipalName is required. Use * to inventory all guest users.'
        }

        $guests = @()
        if ($userPrincipalName -eq '*') {
            $guests = @(Invoke-WithRetry -OperationName 'Load all guest users' -ScriptBlock {
                Get-MgUser -Filter "userType eq 'Guest'" -ConsistencyLevel eventual -All -Property $userSelect -ErrorAction Stop
            })
        }
        else {
            $escapedUpn = Escape-ODataString -Value $userPrincipalName
            $guests = @(Invoke-WithRetry -OperationName "Lookup guest user $userPrincipalName" -ScriptBlock {
                Get-MgUser -Filter "userPrincipalName eq '$escapedUpn' and userType eq 'Guest'" -ConsistencyLevel eventual -Property $userSelect -ErrorAction Stop
            })
        }

        if ($guests.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'GetEntraGuestUser' -Status 'NotFound' -Message 'No matching guest users were found.' -Data ([ordered]@{
                        UserId                          = ''
                        UserPrincipalName               = $userPrincipalName
                        DisplayName                     = ''
                        GivenName                       = ''
                        Surname                         = ''
                        Mail                            = ''
                        AccountEnabled                  = ''
                        UserType                        = ''
                        ExternalUserState               = ''
                        ExternalUserStateChangeDateTime = ''
                        CreationType                    = ''
                        CreatedDateTime                 = ''
                    })))
            $rowNumber++
            continue
        }

        $sortedGuests = @($guests | Sort-Object -Property UserPrincipalName, Id)
        foreach ($guest in $sortedGuests) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey ([string]$guest.UserPrincipalName) -Action 'GetEntraGuestUser' -Status 'Completed' -Message 'Guest user exported.' -Data ([ordered]@{
                        UserId                          = ([string]$guest.Id).Trim()
                        UserPrincipalName               = ([string]$guest.UserPrincipalName).Trim()
                        DisplayName                     = ([string]$guest.DisplayName).Trim()
                        GivenName                       = ([string]$guest.GivenName).Trim()
                        Surname                         = ([string]$guest.Surname).Trim()
                        Mail                            = ([string]$guest.Mail).Trim()
                        AccountEnabled                  = [string]$guest.AccountEnabled
                        UserType                        = ([string]$guest.UserType).Trim()
                        ExternalUserState               = ([string]$guest.ExternalUserState).Trim()
                        ExternalUserStateChangeDateTime = [string]$guest.ExternalUserStateChangeDateTime
                        CreationType                    = ([string]$guest.CreationType).Trim()
                        CreatedDateTime                 = [string]$guest.CreatedDateTime
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($userPrincipalName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'GetEntraGuestUser' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    UserId                          = ''
                    UserPrincipalName               = $userPrincipalName
                    DisplayName                     = ''
                    GivenName                       = ''
                    Surname                         = ''
                    Mail                            = ''
                    AccountEnabled                  = ''
                    UserType                        = ''
                    ExternalUserState               = ''
                    ExternalUserStateChangeDateTime = ''
                    CreationType                    = ''
                    CreatedDateTime                 = ''
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
Write-Status -Message 'Entra ID guest user inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}











