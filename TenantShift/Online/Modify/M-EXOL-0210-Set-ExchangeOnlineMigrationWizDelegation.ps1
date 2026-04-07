<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260406-000000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Grants MigrationWiz delegated access to Exchange Online mailboxes.

.DESCRIPTION
    Grants MigrationWiz delegated access to Exchange Online mailboxes using one of two
    delegation types:
    - Impersonation: Grants the ApplicationImpersonation management role to a specified
      service account, giving it tenant-wide access to all mailboxes. This is the standard
      approach for MigrationWiz cloud-to-cloud and hybrid migrations. TENANT-WIDE EFFECT.
      Requires explicit -Confirm before applying.
    - FullAccess: Grants per-mailbox Full Access permission to a specified delegate account.
      Scoped to individual mailboxes listed in the input CSV.
    Supports -WhatIf for dry-run validation before committing changes.
    All results are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Required fields: UserPrincipalName, DelegationType.
    See the companion .input.csv template for the full column list.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\M-EXOL-0210-Set-ExchangeOnlineMigrationWizDelegation.ps1 -InputCsvPath .\M-EXOL-0210-Set-ExchangeOnlineMigrationWizDelegation.input.csv -WhatIf

    Dry-run: shows what delegation changes would be applied.

.EXAMPLE
    .\M-EXOL-0210-Set-ExchangeOnlineMigrationWizDelegation.ps1 -InputCsvPath .\M-EXOL-0210-Set-ExchangeOnlineMigrationWizDelegation.input.csv

    Apply delegation. Impersonation rows prompt for -Confirm before applying.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator (Organization Management)
    Limitations:      Impersonation delegation is TENANT-WIDE — it grants the service account
                      access to ALL mailboxes in the tenant. Use with care and review with the
                      security team before applying in production.
                      FullAccess delegation is per-mailbox.
                      This script does not remove delegation — run Remove-MailboxPermission manually
                      or via a separate cleanup script after migration completes.

    CSV Fields:
    Column              Type      Required  Description
    ------------------  --------  --------  -----------
    UserPrincipalName   String    Yes       For FullAccess: UPN of the mailbox to delegate access to.
                                            For Impersonation: the service account UPN receiving the role.
    DelegationType      String    Yes       Impersonation or FullAccess
    DelegateAccount     String    No        For FullAccess: UPN of the account to grant access to.
                                            Not required for Impersonation (role is assigned tenant-wide).
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M-EXOL-0210-Set-ExchangeOnlineMigrationWizDelegation_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function New-ModifyResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$RowNumber,
        [Parameter(Mandatory)][string]$PrimaryKey,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$Message,
        [Parameter(Mandatory)][hashtable]$Data
    )

    $base    = New-ResultObject -RowNumber $RowNumber -PrimaryKey $PrimaryKey -Action $Action -Status $Status -Message $Message
    $ordered = [ordered]@{}

    foreach ($prop in $base.PSObject.Properties.Name) {
        $ordered[$prop] = $base.$prop
    }

    foreach ($key in $Data.Keys) {
        $ordered[$key] = $Data[$key]
    }

    return [PSCustomObject]$ordered
}

$requiredHeaders  = @('UserPrincipalName', 'DelegationType')
$validDelegations = @('Impersonation', 'FullAccess')

Write-Status -Message 'Starting Exchange Online MigrationWiz delegation script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

# Track whether Impersonation has already been applied this run (it's tenant-wide — only needs doing once).
$impersonationApplied = $false

foreach ($row in $rows) {
    $upn            = Get-TrimmedValue -Value $row.UserPrincipalName
    $delegationType = Get-TrimmedValue -Value $row.DelegationType
    $delegateAcct   = if ($row.PSObject.Properties['DelegateAccount']) { Get-TrimmedValue -Value $row.DelegateAccount } else { '' }
    $primaryKey     = "${upn}|${delegationType}"

    try {
        if ([string]::IsNullOrWhiteSpace($upn))            { throw 'UserPrincipalName is required.' }
        if ([string]::IsNullOrWhiteSpace($delegationType)) { throw 'DelegationType is required.' }

        $normalizedType = $validDelegations | Where-Object { $_ -ieq $delegationType }
        if (-not $normalizedType) { throw "DelegationType '$delegationType' is invalid. Valid values: Impersonation, FullAccess." }
        $delegationType = $normalizedType

        switch ($delegationType) {
            'Impersonation' {
                # ApplicationImpersonation is tenant-wide — $upn is the service account receiving the role.
                $description = "TENANT-WIDE: Grant ApplicationImpersonation role to '$upn'. This gives '$upn' access to ALL mailboxes in the tenant."

                if (-not $PSCmdlet.ShouldProcess('All mailboxes (tenant-wide)', $description)) {
                    $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetMigrationWizDelegation' -Status 'WhatIf' -Message "WhatIf: would grant ApplicationImpersonation role to '$upn'." -Data ([ordered]@{
                        UserPrincipalName = $upn; DelegationType = $delegationType; DelegateAccount = ''; Timestamp = ''
                    })))
                    $rowNumber++
                    continue
                }

                if (-not $PSCmdlet.ShouldContinue("Granting ApplicationImpersonation to '$upn' gives it FULL ACCESS to ALL mailboxes in the Exchange Online tenant. This is a TENANT-WIDE change. Confirm with your security team before proceeding. Continue?", 'Confirm Impersonation Delegation')) {
                    $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetMigrationWizDelegation' -Status 'Skipped' -Message 'Impersonation delegation cancelled by user.' -Data ([ordered]@{
                        UserPrincipalName = $upn; DelegationType = $delegationType; DelegateAccount = ''; Timestamp = ''
                    })))
                    $rowNumber++
                    continue
                }

                if ($impersonationApplied) {
                    $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetMigrationWizDelegation' -Status 'Skipped' -Message 'ApplicationImpersonation role already granted this run (tenant-wide, only applied once).' -Data ([ordered]@{
                        UserPrincipalName = $upn; DelegationType = $delegationType; DelegateAccount = ''; Timestamp = ''
                    })))
                    $rowNumber++
                    continue
                }

                Invoke-WithRetry -OperationName "Grant ApplicationImpersonation to $upn" -ScriptBlock {
                    New-ManagementRoleAssignment -Name "MigrationWiz-$upn" -Role ApplicationImpersonation -User $upn -ErrorAction Stop
                }

                $impersonationApplied = $true
                $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetMigrationWizDelegation' -Status 'Completed' -Message "ApplicationImpersonation role granted to '$upn' (tenant-wide)." -Data ([ordered]@{
                    UserPrincipalName = $upn; DelegationType = $delegationType; DelegateAccount = ''; Timestamp = (Get-Date -Format 'o')
                })))
            }

            'FullAccess' {
                if ([string]::IsNullOrWhiteSpace($delegateAcct)) {
                    throw 'DelegateAccount is required for FullAccess delegation.'
                }

                $description = "Grant FullAccess on mailbox '$upn' to delegate account '$delegateAcct'"
                if (-not $PSCmdlet.ShouldProcess($upn, $description)) {
                    $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetMigrationWizDelegation' -Status 'WhatIf' -Message "WhatIf: would grant FullAccess on '$upn' to '$delegateAcct'." -Data ([ordered]@{
                        UserPrincipalName = $upn; DelegationType = $delegationType; DelegateAccount = $delegateAcct; Timestamp = ''
                    })))
                    $rowNumber++
                    continue
                }

                Invoke-WithRetry -OperationName "Grant FullAccess on $upn to $delegateAcct" -ScriptBlock {
                    Add-MailboxPermission -Identity $upn -User $delegateAcct -AccessRights FullAccess -InheritanceType All -AutoMapping $false -ErrorAction Stop | Out-Null
                }

                $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetMigrationWizDelegation' -Status 'Completed' -Message "FullAccess granted on '$upn' to '$delegateAcct'." -Data ([ordered]@{
                    UserPrincipalName = $upn; DelegationType = $delegationType; DelegateAccount = $delegateAcct; Timestamp = (Get-Date -Format 'o')
                })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetMigrationWizDelegation' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            UserPrincipalName = $upn; DelegationType = $delegationType; DelegateAccount = $delegateAcct; Timestamp = ''
        })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online MigrationWiz delegation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
