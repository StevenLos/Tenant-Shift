<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-220000

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M0220-Set-ExchangeOnPremMailboxDelegations_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-RecipientKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Recipient
    )

    $primary = Get-TrimmedValue -Value $Recipient.PrimarySmtpAddress
    if (-not [string]::IsNullOrWhiteSpace($primary)) {
        return $primary.ToLowerInvariant()
    }

    $identity = Get-TrimmedValue -Value $Recipient.Identity
    if (-not [string]::IsNullOrWhiteSpace($identity)) {
        return $identity.ToLowerInvariant()
    }

    return ''
}

$requiredHeaders = @(
    'MailboxIdentity',
    'TrusteeIdentity',
    'PermissionType',
    'PermissionAction',
    'AutoMapping'
)

Write-Status -Message 'Starting Exchange on-prem mailbox delegation script.'
Ensure-ExchangeOnPremConnection

$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop
$supportsGrantSendOnBehalf = $setMailboxCommand.Parameters.ContainsKey('GrantSendOnBehalfTo')
$hasRecipientPermissionCmdlets = (Get-Command -Name Get-RecipientPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Add-RecipientPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Remove-RecipientPermission -ErrorAction SilentlyContinue)
$hasAdPermissionCmdlets = (Get-Command -Name Get-ADPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Add-ADPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Remove-ADPermission -ErrorAction SilentlyContinue)

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = Get-TrimmedValue -Value $row.MailboxIdentity
    $trusteeIdentity = Get-TrimmedValue -Value $row.TrusteeIdentity
    $permissionTypeRaw = Get-TrimmedValue -Value $row.PermissionType
    $permissionActionRaw = Get-TrimmedValue -Value $row.PermissionAction

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity) -or [string]::IsNullOrWhiteSpace($trusteeIdentity)) {
            throw 'MailboxIdentity and TrusteeIdentity are required.'
        }

        $permissionType = if ([string]::IsNullOrWhiteSpace($permissionTypeRaw)) { 'FullAccess' } else { $permissionTypeRaw }
        if ($permissionType -notin @('FullAccess', 'SendAs', 'SendOnBehalf')) {
            throw "PermissionType '$permissionType' is invalid. Use FullAccess, SendAs, or SendOnBehalf."
        }

        $permissionAction = if ([string]::IsNullOrWhiteSpace($permissionActionRaw)) { 'Add' } else { $permissionActionRaw }
        if ($permissionAction -notin @('Add', 'Remove')) {
            throw "PermissionAction '$permissionAction' is invalid. Use Add or Remove."
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $mailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|$permissionType" -Action 'SetMailboxDelegation' -Status 'NotFound' -Message 'Mailbox not found.'))
            $rowNumber++
            continue
        }

        $trusteeRecipient = Invoke-WithRetry -OperationName "Lookup trustee recipient $trusteeIdentity" -ScriptBlock {
            Get-Recipient -Identity $trusteeIdentity -ErrorAction SilentlyContinue
        }
        if (-not $trusteeRecipient) {
            throw "Trustee '$trusteeIdentity' was not found."
        }

        $trusteeKey = Get-RecipientKey -Recipient $trusteeRecipient

        if ($permissionType -eq 'FullAccess') {
            $autoMappingRaw = Get-TrimmedValue -Value $row.AutoMapping
            $autoMapping = if ([string]::IsNullOrWhiteSpace($autoMappingRaw)) { $true } else { ConvertTo-Bool -Value $autoMappingRaw }

            $existingPermissions = @(Invoke-WithRetry -OperationName "Load FullAccess delegations for $mailboxIdentity" -ScriptBlock {
                Get-MailboxPermission -Identity $mailbox.Identity -ErrorAction Stop
            })

            $existing = $false
            foreach ($perm in $existingPermissions) {
                if ($perm.IsInherited -or $perm.Deny) {
                    continue
                }

                $permUser = Get-TrimmedValue -Value $perm.User
                if ($permUser.Equals((Get-TrimmedValue -Value $trusteeRecipient.Identity), [System.StringComparison]::OrdinalIgnoreCase) -or $permUser.Equals((Get-TrimmedValue -Value $trusteeRecipient.Name), [System.StringComparison]::OrdinalIgnoreCase) -or $permUser.Equals((Get-TrimmedValue -Value $trusteeRecipient.PrimarySmtpAddress), [System.StringComparison]::OrdinalIgnoreCase)) {
                    if (@($perm.AccessRights) -contains 'FullAccess') {
                        $existing = $true
                        break
                    }
                }
            }

            if ($permissionAction -eq 'Add') {
                if ($existing) {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetMailboxDelegation' -Status 'Skipped' -Message 'FullAccess delegation already exists.'))
                    $rowNumber++
                    continue
                }

                if ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeIdentity", 'Add FullAccess delegation')) {
                    Invoke-WithRetry -OperationName "Add FullAccess delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                        Add-MailboxPermission -Identity $mailbox.Identity -User $trusteeRecipient.Identity -AccessRights FullAccess -InheritanceType All -AutoMapping:$autoMapping -ErrorAction Stop | Out-Null
                    }

                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetMailboxDelegation' -Status 'Added' -Message 'FullAccess delegation added.'))
                }
                else {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetMailboxDelegation' -Status 'WhatIf' -Message 'Delegation change skipped due to WhatIf.'))
                }
            }
            else {
                if (-not $existing) {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetMailboxDelegation' -Status 'Skipped' -Message 'FullAccess delegation does not exist.'))
                    $rowNumber++
                    continue
                }

                if ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeIdentity", 'Remove FullAccess delegation')) {
                    Invoke-WithRetry -OperationName "Remove FullAccess delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                        Remove-MailboxPermission -Identity $mailbox.Identity -User $trusteeRecipient.Identity -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                    }

                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetMailboxDelegation' -Status 'Removed' -Message 'FullAccess delegation removed.'))
                }
                else {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetMailboxDelegation' -Status 'WhatIf' -Message 'Delegation change skipped due to WhatIf.'))
                }
            }

            $rowNumber++
            continue
        }

        if ($permissionType -eq 'SendAs') {
            if (-not $hasRecipientPermissionCmdlets -and -not $hasAdPermissionCmdlets) {
                throw 'No supported cmdlet set for SendAs delegations was found (Get/Add/Remove-RecipientPermission or Get/Add/Remove-ADPermission).'
            }

            if ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeIdentity", "$permissionAction SendAs delegation")) {
                if ($hasRecipientPermissionCmdlets) {
                    try {
                        if ($permissionAction -eq 'Add') {
                            Invoke-WithRetry -OperationName "Add SendAs delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                                Add-RecipientPermission -Identity $mailbox.Identity -Trustee $trusteeRecipient.Identity -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                            }
                            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'Added' -Message 'SendAs delegation add attempted via RecipientPermission cmdlets.'))
                        }
                        else {
                            Invoke-WithRetry -OperationName "Remove SendAs delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                                Remove-RecipientPermission -Identity $mailbox.Identity -Trustee $trusteeRecipient.Identity -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                            }
                            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'Removed' -Message 'SendAs delegation remove attempted via RecipientPermission cmdlets.'))
                        }
                    }
                    catch {
                        $messageLower = ([string]$_.Exception.Message).ToLowerInvariant()
                        if ($permissionAction -eq 'Add' -and $messageLower -match 'already|exists|duplicate') {
                            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'Skipped' -Message 'SendAs delegation already exists.'))
                        }
                        elseif ($permissionAction -eq 'Remove' -and $messageLower -match 'cannot find|not found|doesn''t exist') {
                            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'Skipped' -Message 'SendAs delegation does not exist.'))
                        }
                        else {
                            throw
                        }
                    }
                }
                else {
                    if ($permissionAction -eq 'Add') {
                        Invoke-WithRetry -OperationName "Add SendAs AD delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                            Add-ADPermission -Identity $mailbox.Identity -User $trusteeRecipient.Identity -ExtendedRights 'Send As' -Confirm:$false -ErrorAction Stop | Out-Null
                        }
                        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'Added' -Message 'SendAs delegation add attempted via ADPermission cmdlets.'))
                    }
                    else {
                        Invoke-WithRetry -OperationName "Remove SendAs AD delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                            Remove-ADPermission -Identity $mailbox.Identity -User $trusteeRecipient.Identity -ExtendedRights 'Send As' -Confirm:$false -ErrorAction Stop | Out-Null
                        }
                        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'Removed' -Message 'SendAs delegation remove attempted via ADPermission cmdlets.'))
                    }
                }
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'WhatIf' -Message 'Delegation change skipped due to WhatIf.'))
            }

            $rowNumber++
            continue
        }

        if (-not $supportsGrantSendOnBehalf) {
            throw 'Set-Mailbox -GrantSendOnBehalfTo is not available in this session.'
        }

        $currentDelegates = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($delegate in @($mailbox.GrantSendOnBehalfTo)) {
            $delegateRecipient = Invoke-WithRetry -OperationName "Resolve current SendOnBehalf delegate $delegate" -ScriptBlock {
                Get-Recipient -Identity $delegate -ErrorAction SilentlyContinue
            }

            if ($delegateRecipient) {
                $key = Get-RecipientKey -Recipient $delegateRecipient
                if (-not [string]::IsNullOrWhiteSpace($key)) {
                    $null = $currentDelegates.Add($key)
                }
            }
            else {
                $raw = Get-TrimmedValue -Value $delegate
                if (-not [string]::IsNullOrWhiteSpace($raw)) {
                    $null = $currentDelegates.Add($raw.ToLowerInvariant())
                }
            }
        }

        if ($permissionAction -eq 'Add') {
            if ($currentDelegates.Contains($trusteeKey)) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetMailboxDelegation' -Status 'Skipped' -Message 'SendOnBehalf delegation already exists.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeIdentity", 'Add SendOnBehalf delegation')) {
                Invoke-WithRetry -OperationName "Add SendOnBehalf delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                    Set-Mailbox -Identity $mailbox.Identity -GrantSendOnBehalfTo @{ Add = $trusteeRecipient.Identity } -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetMailboxDelegation' -Status 'Added' -Message 'SendOnBehalf delegation added.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetMailboxDelegation' -Status 'WhatIf' -Message 'Delegation change skipped due to WhatIf.'))
            }
        }
        else {
            if (-not $currentDelegates.Contains($trusteeKey)) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetMailboxDelegation' -Status 'Skipped' -Message 'SendOnBehalf delegation does not exist.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeIdentity", 'Remove SendOnBehalf delegation')) {
                Invoke-WithRetry -OperationName "Remove SendOnBehalf delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                    Set-Mailbox -Identity $mailbox.Identity -GrantSendOnBehalfTo @{ Remove = $trusteeRecipient.Identity } -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetMailboxDelegation' -Status 'Removed' -Message 'SendOnBehalf delegation removed.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetMailboxDelegation' -Status 'WhatIf' -Message 'Delegation change skipped due to WhatIf.'))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity|$trusteeIdentity|$permissionTypeRaw|$permissionActionRaw) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|$permissionTypeRaw|$permissionActionRaw" -Action 'SetMailboxDelegation' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem mailbox delegation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
