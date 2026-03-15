<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-131500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
ActiveDirectory

.MODULEVERSIONPOLICY
Best-effort PSGallery check (offline-safe)
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M0010-Set-ActiveDirectoryTemporaryPasswords_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Resolve-TargetAdUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IdentityType,

        [Parameter(Mandatory)]
        [string]$IdentityValue
    )

    $normalizedType = $IdentityType.Trim().ToLowerInvariant()
    switch ($normalizedType) {
        'samaccountname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADUser -Filter "SamAccountName -eq '$escaped'" -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'userprincipalname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADUser -Filter "UserPrincipalName -eq '$escaped'" -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'distinguishedname' {
            return Get-ADUser -Identity $IdentityValue -ErrorAction SilentlyContinue
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            return Get-ADUser -Identity $guid -ErrorAction SilentlyContinue
        }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use SamAccountName, UserPrincipalName, DistinguishedName, or ObjectGuid."
        }
    }
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'IdentityType',
    'IdentityValue',
    'AccountPassword',
    'ChangePasswordAtLogon',
    'UnlockAccount',
    'EnableAccount'
)

Write-Status -Message 'Starting Active Directory temporary password reset script.'
Ensure-ActiveDirectoryConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $identityType = Get-TrimmedValue -Value $row.IdentityType
    $identityValue = Get-TrimmedValue -Value $row.IdentityValue
    $primaryKey = "${identityType}:$identityValue"

    try {
        if ([string]::IsNullOrWhiteSpace($identityType) -or [string]::IsNullOrWhiteSpace($identityValue)) {
            throw 'IdentityType and IdentityValue are required.'
        }

        $targetUser = Invoke-WithRetry -OperationName "Resolve AD user $primaryKey" -ScriptBlock {
            Resolve-TargetAdUser -IdentityType $identityType -IdentityValue $identityValue
        }

        if (-not $targetUser) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryTemporaryPassword' -Status 'NotFound' -Message 'Target user was not found.'))
            $rowNumber++
            continue
        }

        $resolvedKey = if (-not [string]::IsNullOrWhiteSpace($targetUser.UserPrincipalName)) { $targetUser.UserPrincipalName } else { $targetUser.SamAccountName }
        $accountPassword = Get-TrimmedValue -Value $row.AccountPassword
        if ([string]::IsNullOrWhiteSpace($accountPassword)) {
            throw 'AccountPassword is required.'
        }

        $changePasswordAtLogon = Get-NullableBool -Value $row.ChangePasswordAtLogon
        if ($null -eq $changePasswordAtLogon) {
            $changePasswordAtLogon = $true
        }

        $unlockAccount = Get-NullableBool -Value $row.UnlockAccount
        if ($null -eq $unlockAccount) {
            $unlockAccount = $false
        }

        $enableAccount = Get-NullableBool -Value $row.EnableAccount
        $messages = [System.Collections.Generic.List[string]]::new()

        if ($PSCmdlet.ShouldProcess($resolvedKey, 'Reset AD account password')) {
            $securePassword = ConvertTo-SecureString -String $accountPassword -AsPlainText -Force
            Invoke-WithRetry -OperationName "Reset AD account password $resolvedKey" -ScriptBlock {
                Set-ADAccountPassword -Identity $targetUser.DistinguishedName -Reset -NewPassword $securePassword -ErrorAction Stop
            }
            $messages.Add('Password reset completed.')

            Invoke-WithRetry -OperationName "Set ChangePasswordAtLogon for $resolvedKey" -ScriptBlock {
                Set-ADUser -Identity $targetUser.DistinguishedName -ChangePasswordAtLogon $changePasswordAtLogon -ErrorAction Stop
            }
            $messages.Add("ChangePasswordAtLogon set to '$changePasswordAtLogon'.")

            if ($unlockAccount) {
                $userLockState = Invoke-WithRetry -OperationName "Check lockout state for $resolvedKey" -ScriptBlock {
                    Get-ADUser -Identity $targetUser.DistinguishedName -Properties LockedOut -ErrorAction Stop
                }

                if ([bool]$userLockState.LockedOut) {
                    Invoke-WithRetry -OperationName "Unlock AD account $resolvedKey" -ScriptBlock {
                        Unlock-ADAccount -Identity $targetUser.DistinguishedName -ErrorAction Stop
                    }
                    $messages.Add('Account unlocked.')
                }
                else {
                    $messages.Add('Account was not locked.')
                }
            }

            if ($null -ne $enableAccount) {
                if ($enableAccount -and -not [bool]$targetUser.Enabled) {
                    Invoke-WithRetry -OperationName "Enable AD account $resolvedKey" -ScriptBlock {
                        Enable-ADAccount -Identity $targetUser.DistinguishedName -ErrorAction Stop
                    }
                    $messages.Add('Account enabled.')
                }
                elseif ((-not $enableAccount) -and [bool]$targetUser.Enabled) {
                    Invoke-WithRetry -OperationName "Disable AD account $resolvedKey" -ScriptBlock {
                        Disable-ADAccount -Identity $targetUser.DistinguishedName -ErrorAction Stop
                    }
                    $messages.Add('Account disabled.')
                }
                else {
                    $messages.Add("Account enabled state already '$enableAccount'.")
                }
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'SetActiveDirectoryTemporaryPassword' -Status 'Completed' -Message ($messages -join ' ')))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'SetActiveDirectoryTemporaryPassword' -Status 'WhatIf' -Message 'Password reset skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryTemporaryPassword' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory temporary password reset script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
