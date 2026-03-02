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

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_P3001-Create-EntraUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {


$requiredHeaders = @(
    'UserPrincipalName',
    'DisplayName',
    'GivenName',
    'Surname',
    'MailNickname',
    'Password',
    'ForceChangePasswordNextSignIn',
    'AccountEnabled',
    'UsageLocation',
    'Department',
    'JobTitle',
    'MobilePhone',
    'BusinessPhones'
)

Write-Status -Message 'Starting Entra ID user creation script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('User.ReadWrite.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $upn = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw 'UserPrincipalName is required.'
        }

        $escapedUpn = Escape-ODataString -Value $upn
        $existingUser = Invoke-WithRetry -OperationName "Lookup user $upn" -ScriptBlock {
            Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -ErrorAction Stop | Select-Object -First 1
        }
        if ($existingUser) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'CreateUser' -Status 'Skipped' -Message 'User already exists.'))
            $rowNumber++
            continue
        }

        $mailNickname = ([string]$row.MailNickname).Trim()
        if ([string]::IsNullOrWhiteSpace($mailNickname)) {
            $mailNickname = $upn.Split('@')[0]
        }

        $displayName = ([string]$row.DisplayName).Trim()
        $givenName = ([string]$row.GivenName).Trim()
        $surname = ([string]$row.Surname).Trim()
        $password = [string]$row.Password

        if ([string]::IsNullOrWhiteSpace($displayName) -or [string]::IsNullOrWhiteSpace($givenName) -or [string]::IsNullOrWhiteSpace($surname) -or [string]::IsNullOrWhiteSpace($password)) {
            throw 'DisplayName, GivenName, Surname, and Password are required.'
        }

        $businessPhones = ConvertTo-Array -Value ([string]$row.BusinessPhones)

        $body = @{
            accountEnabled    = ConvertTo-Bool -Value $row.AccountEnabled -Default $true
            displayName       = $displayName
            givenName         = $givenName
            surname           = $surname
            mailNickname      = $mailNickname
            userPrincipalName = $upn
            passwordProfile   = @{
                password                             = $password
                forceChangePasswordNextSignIn       = ConvertTo-Bool -Value $row.ForceChangePasswordNextSignIn -Default $true
            }
        }

        $usageLocation = ([string]$row.UsageLocation).Trim()
        if (-not [string]::IsNullOrWhiteSpace($usageLocation)) {
            $body.usageLocation = $usageLocation
        }

        $department = ([string]$row.Department).Trim()
        if (-not [string]::IsNullOrWhiteSpace($department)) {
            $body.department = $department
        }

        $jobTitle = ([string]$row.JobTitle).Trim()
        if (-not [string]::IsNullOrWhiteSpace($jobTitle)) {
            $body.jobTitle = $jobTitle
        }

        $mobilePhone = ([string]$row.MobilePhone).Trim()
        if (-not [string]::IsNullOrWhiteSpace($mobilePhone)) {
            $body.mobilePhone = $mobilePhone
        }

        if ($businessPhones.Count -gt 0) {
            $body.businessPhones = $businessPhones
        }

        if ($PSCmdlet.ShouldProcess($upn, 'Create Entra ID user')) {
            Invoke-WithRetry -OperationName "Create user $upn" -ScriptBlock {
                New-MgUser -BodyParameter $body -ErrorAction Stop | Out-Null
            }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'CreateUser' -Status 'Created' -Message 'User created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'CreateUser' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'CreateUser' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID user creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







