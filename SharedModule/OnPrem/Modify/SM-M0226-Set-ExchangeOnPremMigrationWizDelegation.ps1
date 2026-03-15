<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260305-081600

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$DefaultDelegateIdentity = 'svc_bittitan',

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M0226-Set-ExchangeOnPremMigrationWizDelegation_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'MailboxIdentity',
    'DelegateIdentity',
    'PermissionType',
    'PermissionAction',
    'AutoMapping',
    'Notes'
)

Write-Status -Message 'Starting Exchange on-prem MigrationWiz delegation script.'
Ensure-ExchangeOnPremConnection

$hasRecipientPermissionCmdlets = (Get-Command -Name Get-RecipientPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Add-RecipientPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Remove-RecipientPermission -ErrorAction SilentlyContinue)
$hasAdPermissionCmdlets = (Get-Command -Name Get-ADPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Add-ADPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Remove-ADPermission -ErrorAction SilentlyContinue)

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = Get-TrimmedValue -Value $row.MailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
            throw 'MailboxIdentity is required.'
        }

        $delegateIdentity = Get-TrimmedValue -Value $row.DelegateIdentity
        if ([string]::IsNullOrWhiteSpace($delegateIdentity)) {
            $delegateIdentity = Get-TrimmedValue -Value $DefaultDelegateIdentity
        }

        if ([string]::IsNullOrWhiteSpace($delegateIdentity)) {
            throw 'DelegateIdentity is required (or provide -DefaultDelegateIdentity).'
        }

        $permissionTypeRaw = Get-TrimmedValue -Value $row.PermissionType
        $permissionActionRaw = Get-TrimmedValue -Value $row.PermissionAction
        $permissionType = if ([string]::IsNullOrWhiteSpace($permissionTypeRaw)) { 'FullAccess' } else { $permissionTypeRaw }
        $permissionAction = if ([string]::IsNullOrWhiteSpace($permissionActionRaw)) { 'Add' } else { $permissionActionRaw }

        if ($permissionType -notin @('FullAccess', 'SendAs')) {
            throw "PermissionType '$permissionType' is invalid. Use FullAccess or SendAs."
        }

        if ($permissionAction -notin @('Add', 'Remove')) {
            throw "PermissionAction '$permissionAction' is invalid. Use Add or Remove."
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $mailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|$permissionType" -Action 'SetMigrationWizDelegation' -Status 'NotFound' -Message 'Mailbox not found.'))
            $rowNumber++
            continue
        }

        $mailboxType = Get-TrimmedValue -Value $mailbox.RecipientTypeDetails
        if ($mailboxType -notin @('UserMailbox', 'SharedMailbox')) {
            throw "Recipient '$mailboxIdentity' is '$mailboxType'. Expected UserMailbox or SharedMailbox."
        }

        $delegateRecipient = Invoke-WithRetry -OperationName "Lookup delegate recipient $delegateIdentity" -ScriptBlock {
            Get-Recipient -Identity $delegateIdentity -ErrorAction SilentlyContinue
        }

        if (-not $delegateRecipient) {
            throw "Delegate '$delegateIdentity' was not found."
        }

        if ($permissionType -eq 'FullAccess') {
            $autoMappingRaw = Get-TrimmedValue -Value $row.AutoMapping
            $autoMapping = if ([string]::IsNullOrWhiteSpace($autoMappingRaw)) { $true } else { ConvertTo-Bool -Value $autoMappingRaw }

            $existingPermissions = @(Invoke-WithRetry -OperationName "Load mailbox permissions for $mailboxIdentity" -ScriptBlock {
                Get-MailboxPermission -Identity $mailbox.Identity -ErrorAction Stop
            })

            $existing = $false
            foreach ($permission in $existingPermissions) {
                if ($permission.Deny -or $permission.IsInherited) { continue }

                $accessRights = @($permission.AccessRights | ForEach-Object { Get-TrimmedValue -Value $_ })
                if ($accessRights -notcontains 'FullAccess') { continue }

                $permUser = Get-TrimmedValue -Value $permission.User
                if ($permUser.Equals((Get-TrimmedValue -Value $delegateRecipient.Identity), [System.StringComparison]::OrdinalIgnoreCase) -or
                    $permUser.Equals((Get-TrimmedValue -Value $delegateRecipient.Name), [System.StringComparison]::OrdinalIgnoreCase) -or
                    $permUser.Equals((Get-TrimmedValue -Value $delegateRecipient.PrimarySmtpAddress), [System.StringComparison]::OrdinalIgnoreCase)) {
                    $existing = $true
                    break
                }
            }

            if ($permissionAction -eq 'Add') {
                if ($existing) {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|FullAccess" -Action 'SetMigrationWizDelegation' -Status 'Skipped' -Message 'FullAccess delegation already exists.'))
                    $rowNumber++
                    continue
                }

                if ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $delegateIdentity", 'Add FullAccess delegation')) {
                    Invoke-WithRetry -OperationName "Add FullAccess delegation $mailboxIdentity -> $delegateIdentity" -ScriptBlock {
                        Add-MailboxPermission -Identity $mailbox.Identity -User $delegateRecipient.Identity -AccessRights FullAccess -InheritanceType All -AutoMapping:$autoMapping -ErrorAction Stop | Out-Null
                    }

                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|FullAccess" -Action 'SetMigrationWizDelegation' -Status 'Added' -Message 'FullAccess delegation added.'))
                }
                else {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|FullAccess" -Action 'SetMigrationWizDelegation' -Status 'WhatIf' -Message 'Delegation change skipped due to WhatIf.'))
                }
            }
            else {
                if (-not $existing) {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|FullAccess" -Action 'SetMigrationWizDelegation' -Status 'Skipped' -Message 'FullAccess delegation does not exist.'))
                    $rowNumber++
                    continue
                }

                if ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $delegateIdentity", 'Remove FullAccess delegation')) {
                    Invoke-WithRetry -OperationName "Remove FullAccess delegation $mailboxIdentity -> $delegateIdentity" -ScriptBlock {
                        Remove-MailboxPermission -Identity $mailbox.Identity -User $delegateRecipient.Identity -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                    }

                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|FullAccess" -Action 'SetMigrationWizDelegation' -Status 'Removed' -Message 'FullAccess delegation removed.'))
                }
                else {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|FullAccess" -Action 'SetMigrationWizDelegation' -Status 'WhatIf' -Message 'Delegation change skipped due to WhatIf.'))
                }
            }

            $rowNumber++
            continue
        }

        if (-not $hasRecipientPermissionCmdlets -and -not $hasAdPermissionCmdlets) {
            throw 'No supported cmdlet set for SendAs delegations was found (Get/Add/Remove-RecipientPermission or Get/Add/Remove-ADPermission).'
        }

        if ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $delegateIdentity", "$permissionAction SendAs delegation")) {
            if ($hasRecipientPermissionCmdlets) {
                try {
                    if ($permissionAction -eq 'Add') {
                        Invoke-WithRetry -OperationName "Add SendAs delegation $mailboxIdentity -> $delegateIdentity" -ScriptBlock {
                            Add-RecipientPermission -Identity $mailbox.Identity -Trustee $delegateRecipient.Identity -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                        }
                        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|SendAs" -Action 'SetMigrationWizDelegation' -Status 'Added' -Message 'SendAs delegation add attempted via RecipientPermission cmdlets.'))
                    }
                    else {
                        Invoke-WithRetry -OperationName "Remove SendAs delegation $mailboxIdentity -> $delegateIdentity" -ScriptBlock {
                            Remove-RecipientPermission -Identity $mailbox.Identity -Trustee $delegateRecipient.Identity -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                        }
                        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|SendAs" -Action 'SetMigrationWizDelegation' -Status 'Removed' -Message 'SendAs delegation remove attempted via RecipientPermission cmdlets.'))
                    }
                }
                catch {
                    $messageLower = ([string]$_.Exception.Message).ToLowerInvariant()
                    if ($permissionAction -eq 'Add' -and $messageLower -match 'already|exists|duplicate') {
                        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|SendAs" -Action 'SetMigrationWizDelegation' -Status 'Skipped' -Message 'SendAs delegation already exists.'))
                    }
                    elseif ($permissionAction -eq 'Remove' -and $messageLower -match 'cannot find|not found|doesn''t exist') {
                        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|SendAs" -Action 'SetMigrationWizDelegation' -Status 'Skipped' -Message 'SendAs delegation does not exist.'))
                    }
                    else {
                        throw
                    }
                }
            }
            else {
                if ($permissionAction -eq 'Add') {
                    Invoke-WithRetry -OperationName "Add SendAs AD delegation $mailboxIdentity -> $delegateIdentity" -ScriptBlock {
                        Add-ADPermission -Identity $mailbox.Identity -User $delegateRecipient.Identity -ExtendedRights 'Send As' -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|SendAs" -Action 'SetMigrationWizDelegation' -Status 'Added' -Message 'SendAs delegation add attempted via ADPermission cmdlets.'))
                }
                else {
                    Invoke-WithRetry -OperationName "Remove SendAs AD delegation $mailboxIdentity -> $delegateIdentity" -ScriptBlock {
                        Remove-ADPermission -Identity $mailbox.Identity -User $delegateRecipient.Identity -ExtendedRights 'Send As' -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|SendAs" -Action 'SetMigrationWizDelegation' -Status 'Removed' -Message 'SendAs delegation remove attempted via ADPermission cmdlets.'))
                }
            }
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$delegateIdentity|SendAs" -Action 'SetMigrationWizDelegation' -Status 'WhatIf' -Message 'Delegation change skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$($row.DelegateIdentity)|$($row.PermissionType)|$($row.PermissionAction)" -Action 'SetMigrationWizDelegation' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem MigrationWiz delegation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
