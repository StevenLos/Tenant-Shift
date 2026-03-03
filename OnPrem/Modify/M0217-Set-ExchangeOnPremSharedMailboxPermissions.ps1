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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M0217-Set-ExchangeOnPremSharedMailboxPermissions_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'SharedMailboxIdentity',
    'TrusteeIdentity',
    'PermissionType',
    'PermissionAction',
    'AutoMapping'
)

Write-Status -Message 'Starting Exchange on-prem shared mailbox permission script.'
Ensure-ExchangeOnPremConnection

$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop
$supportsGrantSendOnBehalf = $setMailboxCommand.Parameters.ContainsKey('GrantSendOnBehalfTo')
$hasRecipientPermissionCmdlets = (Get-Command -Name Get-RecipientPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Add-RecipientPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Remove-RecipientPermission -ErrorAction SilentlyContinue)
$hasAdPermissionCmdlets = (Get-Command -Name Get-ADPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Add-ADPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Remove-ADPermission -ErrorAction SilentlyContinue)

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $sharedMailboxIdentity = Get-TrimmedValue -Value $row.SharedMailboxIdentity
    $trusteeIdentity = Get-TrimmedValue -Value $row.TrusteeIdentity
    $permissionTypeRaw = Get-TrimmedValue -Value $row.PermissionType
    $permissionActionRaw = Get-TrimmedValue -Value $row.PermissionAction

    try {
        if ([string]::IsNullOrWhiteSpace($sharedMailboxIdentity) -or [string]::IsNullOrWhiteSpace($trusteeIdentity)) {
            throw 'SharedMailboxIdentity and TrusteeIdentity are required.'
        }

        $permissionType = if ([string]::IsNullOrWhiteSpace($permissionTypeRaw)) { 'FullAccess' } else { $permissionTypeRaw }
        if ($permissionType -notin @('FullAccess', 'SendAs', 'SendOnBehalf')) {
            throw "PermissionType '$permissionType' is invalid. Use FullAccess, SendAs, or SendOnBehalf."
        }

        $permissionAction = if ([string]::IsNullOrWhiteSpace($permissionActionRaw)) { 'Add' } else { $permissionActionRaw }
        if ($permissionAction -notin @('Add', 'Remove')) {
            throw "PermissionAction '$permissionAction' is invalid. Use Add or Remove."
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup shared mailbox $sharedMailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $sharedMailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|$permissionType" -Action 'SetSharedMailboxPermission' -Status 'NotFound' -Message 'Shared mailbox not found.'))
            $rowNumber++
            continue
        }

        if ((Get-TrimmedValue -Value $mailbox.RecipientTypeDetails) -ne 'SharedMailbox') {
            throw "Recipient '$sharedMailboxIdentity' is '$($mailbox.RecipientTypeDetails)'. Expected SharedMailbox."
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

            $existingPermissions = @(Invoke-WithRetry -OperationName "Load FullAccess permissions for $sharedMailboxIdentity" -ScriptBlock {
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
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetSharedMailboxPermission' -Status 'Skipped' -Message 'FullAccess permission already exists.'))
                    $rowNumber++
                    continue
                }

                if ($PSCmdlet.ShouldProcess("$sharedMailboxIdentity -> $trusteeIdentity", 'Add FullAccess permission')) {
                    Invoke-WithRetry -OperationName "Add FullAccess permission $sharedMailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                        Add-MailboxPermission -Identity $mailbox.Identity -User $trusteeRecipient.Identity -AccessRights FullAccess -InheritanceType All -AutoMapping:$autoMapping -ErrorAction Stop | Out-Null
                    }

                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetSharedMailboxPermission' -Status 'Added' -Message 'FullAccess permission added.'))
                }
                else {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetSharedMailboxPermission' -Status 'WhatIf' -Message 'Permission change skipped due to WhatIf.'))
                }
            }
            else {
                if (-not $existing) {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetSharedMailboxPermission' -Status 'Skipped' -Message 'FullAccess permission does not exist.'))
                    $rowNumber++
                    continue
                }

                if ($PSCmdlet.ShouldProcess("$sharedMailboxIdentity -> $trusteeIdentity", 'Remove FullAccess permission')) {
                    Invoke-WithRetry -OperationName "Remove FullAccess permission $sharedMailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                        Remove-MailboxPermission -Identity $mailbox.Identity -User $trusteeRecipient.Identity -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                    }

                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetSharedMailboxPermission' -Status 'Removed' -Message 'FullAccess permission removed.'))
                }
                else {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetSharedMailboxPermission' -Status 'WhatIf' -Message 'Permission change skipped due to WhatIf.'))
                }
            }

            $rowNumber++
            continue
        }

        if ($permissionType -eq 'SendAs') {
            if (-not $hasRecipientPermissionCmdlets -and -not $hasAdPermissionCmdlets) {
                throw 'No supported cmdlet set for SendAs permissions was found (Get/Add/Remove-RecipientPermission or Get/Add/Remove-ADPermission).'
            }

            if ($PSCmdlet.ShouldProcess("$sharedMailboxIdentity -> $trusteeIdentity", "$permissionAction SendAs permission")) {
                if ($hasRecipientPermissionCmdlets) {
                    try {
                        if ($permissionAction -eq 'Add') {
                            Invoke-WithRetry -OperationName "Add SendAs permission $sharedMailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                                Add-RecipientPermission -Identity $mailbox.Identity -Trustee $trusteeRecipient.Identity -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                            }
                            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetSharedMailboxPermission' -Status 'Added' -Message 'SendAs permission add attempted via RecipientPermission cmdlets.'))
                        }
                        else {
                            Invoke-WithRetry -OperationName "Remove SendAs permission $sharedMailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                                Remove-RecipientPermission -Identity $mailbox.Identity -Trustee $trusteeRecipient.Identity -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                            }
                            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetSharedMailboxPermission' -Status 'Removed' -Message 'SendAs permission remove attempted via RecipientPermission cmdlets.'))
                        }
                    }
                    catch {
                        $messageLower = ([string]$_.Exception.Message).ToLowerInvariant()
                        if ($permissionAction -eq 'Add' -and $messageLower -match 'already|exists|duplicate') {
                            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetSharedMailboxPermission' -Status 'Skipped' -Message 'SendAs permission already exists.'))
                        }
                        elseif ($permissionAction -eq 'Remove' -and $messageLower -match 'cannot find|not found|doesn''t exist') {
                            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetSharedMailboxPermission' -Status 'Skipped' -Message 'SendAs permission does not exist.'))
                        }
                        else {
                            throw
                        }
                    }
                }
                else {
                    if ($permissionAction -eq 'Add') {
                        Invoke-WithRetry -OperationName "Add SendAs AD permission $sharedMailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                            Add-ADPermission -Identity $mailbox.Identity -User $trusteeRecipient.Identity -ExtendedRights 'Send As' -Confirm:$false -ErrorAction Stop | Out-Null
                        }
                        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetSharedMailboxPermission' -Status 'Added' -Message 'SendAs permission add attempted via ADPermission cmdlets.'))
                    }
                    else {
                        Invoke-WithRetry -OperationName "Remove SendAs AD permission $sharedMailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                            Remove-ADPermission -Identity $mailbox.Identity -User $trusteeRecipient.Identity -ExtendedRights 'Send As' -Confirm:$false -ErrorAction Stop | Out-Null
                        }
                        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetSharedMailboxPermission' -Status 'Removed' -Message 'SendAs permission remove attempted via ADPermission cmdlets.'))
                    }
                }
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetSharedMailboxPermission' -Status 'WhatIf' -Message 'Permission change skipped due to WhatIf.'))
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
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetSharedMailboxPermission' -Status 'Skipped' -Message 'SendOnBehalf delegate already exists.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$sharedMailboxIdentity -> $trusteeIdentity", 'Add SendOnBehalf delegate')) {
                Invoke-WithRetry -OperationName "Add SendOnBehalf delegate $sharedMailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                    Set-Mailbox -Identity $mailbox.Identity -GrantSendOnBehalfTo @{ Add = $trusteeRecipient.Identity } -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetSharedMailboxPermission' -Status 'Added' -Message 'SendOnBehalf delegate added.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetSharedMailboxPermission' -Status 'WhatIf' -Message 'Permission change skipped due to WhatIf.'))
            }
        }
        else {
            if (-not $currentDelegates.Contains($trusteeKey)) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetSharedMailboxPermission' -Status 'Skipped' -Message 'SendOnBehalf delegate does not exist.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$sharedMailboxIdentity -> $trusteeIdentity", 'Remove SendOnBehalf delegate')) {
                Invoke-WithRetry -OperationName "Remove SendOnBehalf delegate $sharedMailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                    Set-Mailbox -Identity $mailbox.Identity -GrantSendOnBehalfTo @{ Remove = $trusteeRecipient.Identity } -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetSharedMailboxPermission' -Status 'Removed' -Message 'SendOnBehalf delegate removed.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetSharedMailboxPermission' -Status 'WhatIf' -Message 'Permission change skipped due to WhatIf.'))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($sharedMailboxIdentity|$trusteeIdentity|$permissionTypeRaw|$permissionActionRaw) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$sharedMailboxIdentity|$trusteeIdentity|$permissionTypeRaw|$permissionActionRaw" -Action 'SetSharedMailboxPermission' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem shared mailbox permission script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
