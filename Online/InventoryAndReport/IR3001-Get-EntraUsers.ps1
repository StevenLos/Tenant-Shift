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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3001-Get-EntraUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

Write-Status -Message 'Starting Entra ID user inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('User.Read.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$userSelect = 'id,userPrincipalName,displayName,givenName,surname,mailNickname,accountEnabled,userType,usageLocation,department,jobTitle,mail,mobilePhone,businessPhones,createdDateTime,lastPasswordChangeDateTime,onPremisesSyncEnabled'

$rowNumber = 1
foreach ($row in $rows) {
    $userPrincipalName = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
            throw 'UserPrincipalName is required. Use * to inventory all users.'
        }

        $users = @()
        if ($userPrincipalName -eq '*') {
            $users = @(Invoke-WithRetry -OperationName 'Load all users' -ScriptBlock {
                Get-MgUser -All -Property $userSelect -ErrorAction Stop
            })
        }
        else {
            $escapedUpn = Escape-ODataString -Value $userPrincipalName
            $users = @(Invoke-WithRetry -OperationName "Lookup user $userPrincipalName" -ScriptBlock {
                Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -Property $userSelect -ErrorAction Stop
            })
        }

        if ($users.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'GetEntraUser' -Status 'NotFound' -Message 'No matching users were found.' -Data ([ordered]@{
                        UserId                     = ''
                        UserPrincipalName          = $userPrincipalName
                        DisplayName                = ''
                        GivenName                  = ''
                        Surname                    = ''
                        MailNickname               = ''
                        AccountEnabled             = ''
                        UserType                   = ''
                        UsageLocation              = ''
                        Department                 = ''
                        JobTitle                   = ''
                        Mail                       = ''
                        MobilePhone                = ''
                        BusinessPhones             = ''
                        CreatedDateTime            = ''
                        LastPasswordChangeDateTime = ''
                        OnPremisesSyncEnabled      = ''
                    })))
            $rowNumber++
            continue
        }

        $sortedUsers = @($users | Sort-Object -Property UserPrincipalName, Id)
        foreach ($user in $sortedUsers) {
            $businessPhones = @($user.BusinessPhones)
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey ([string]$user.UserPrincipalName) -Action 'GetEntraUser' -Status 'Completed' -Message 'User exported.' -Data ([ordered]@{
                        UserId                     = ([string]$user.Id).Trim()
                        UserPrincipalName          = ([string]$user.UserPrincipalName).Trim()
                        DisplayName                = ([string]$user.DisplayName).Trim()
                        GivenName                  = ([string]$user.GivenName).Trim()
                        Surname                    = ([string]$user.Surname).Trim()
                        MailNickname               = ([string]$user.MailNickname).Trim()
                        AccountEnabled             = [string]$user.AccountEnabled
                        UserType                   = ([string]$user.UserType).Trim()
                        UsageLocation              = ([string]$user.UsageLocation).Trim()
                        Department                 = ([string]$user.Department).Trim()
                        JobTitle                   = ([string]$user.JobTitle).Trim()
                        Mail                       = ([string]$user.Mail).Trim()
                        MobilePhone                = ([string]$user.MobilePhone).Trim()
                        BusinessPhones             = ($businessPhones -join ';')
                        CreatedDateTime            = [string]$user.CreatedDateTime
                        LastPasswordChangeDateTime = [string]$user.LastPasswordChangeDateTime
                        OnPremisesSyncEnabled      = [string]$user.OnPremisesSyncEnabled
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($userPrincipalName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'GetEntraUser' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    UserId                     = ''
                    UserPrincipalName          = $userPrincipalName
                    DisplayName                = ''
                    GivenName                  = ''
                    Surname                    = ''
                    MailNickname               = ''
                    AccountEnabled             = ''
                    UserType                   = ''
                    UsageLocation              = ''
                    Department                 = ''
                    JobTitle                   = ''
                    Mail                       = ''
                    MobilePhone                = ''
                    BusinessPhones             = ''
                    CreatedDateTime            = ''
                    LastPasswordChangeDateTime = ''
                    OnPremisesSyncEnabled      = ''
                })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID user inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







