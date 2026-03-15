<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-131500

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3004-Set-EntraUserPasswordResets_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-TrimmedValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    return ([string]$Value).Trim()
}

function Get-NullableBool {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return (ConvertTo-Bool -Value $text)
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'UserPrincipalName',
    'Password',
    'ForceChangePasswordNextSignIn',
    'ForceChangePasswordNextSignInWithMfa'
)

Write-Status -Message 'Starting Entra ID user password reset script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('User.ReadWrite.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $upn = Get-TrimmedValue -Value $row.UserPrincipalName

    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw 'UserPrincipalName is required.'
        }

        $password = [string]$row.Password
        if ([string]::IsNullOrWhiteSpace($password)) {
            throw 'Password is required.'
        }

        $forceChangePasswordNextSignIn = Get-NullableBool -Value $row.ForceChangePasswordNextSignIn
        if ($null -eq $forceChangePasswordNextSignIn) {
            $forceChangePasswordNextSignIn = $true
        }

        $forceChangePasswordNextSignInWithMfa = Get-NullableBool -Value $row.ForceChangePasswordNextSignInWithMfa
        if ($null -eq $forceChangePasswordNextSignInWithMfa) {
            $forceChangePasswordNextSignInWithMfa = $false
        }

        if ($forceChangePasswordNextSignInWithMfa -and -not $forceChangePasswordNextSignIn) {
            throw 'ForceChangePasswordNextSignInWithMfa cannot be TRUE when ForceChangePasswordNextSignIn is FALSE.'
        }

        $escapedUpn = Escape-ODataString -Value $upn
        $users = @(Invoke-WithRetry -OperationName "Lookup user $upn" -ScriptBlock {
            Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -Property 'id,userPrincipalName' -ErrorAction Stop
        })

        if ($users.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'SetEntraUserPassword' -Status 'NotFound' -Message 'User not found.'))
            $rowNumber++
            continue
        }

        if ($users.Count -gt 1) {
            throw "Multiple users were returned for UPN '$upn'. Resolve duplicate directory objects before retrying."
        }

        $user = $users[0]
        $body = @{
            passwordProfile = @{
                password                             = $password
                forceChangePasswordNextSignIn        = $forceChangePasswordNextSignIn
                forceChangePasswordNextSignInWithMfa = $forceChangePasswordNextSignInWithMfa
            }
        }

        if ($PSCmdlet.ShouldProcess($upn, 'Reset Entra ID user password')) {
            Invoke-WithRetry -OperationName "Reset password for $upn" -ScriptBlock {
                Update-MgUser -UserId $user.Id -BodyParameter $body -ErrorAction Stop | Out-Null
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'SetEntraUserPassword' -Status 'Updated' -Message 'Password reset completed.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'SetEntraUserPassword' -Status 'WhatIf' -Message 'Password reset skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'SetEntraUserPassword' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID user password reset script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
