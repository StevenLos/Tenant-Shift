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
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3002-Get-EntraGuestUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

Write-Status -Message 'Starting Entra ID guest user inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('User.Read.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
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

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID guest user inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







