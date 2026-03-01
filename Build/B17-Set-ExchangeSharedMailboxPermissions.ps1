#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath ("Results_B17-Set-ExchangeSharedMailboxPermissions_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'M365.Common.psm1') -Force -DisableNameChecking

$requiredHeaders = @(
    'SharedMailboxIdentity',
    'TrusteeUserPrincipalName',
    'FullAccess',
    'ReadOnly',
    'SendAs',
    'SendOnBehalf',
    'AutoMapping'
)

Write-Status -Message 'Starting shared mailbox permission assignment script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = ([string]$row.SharedMailboxIdentity).Trim()
    $trusteeUpn = ([string]$row.TrusteeUserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity) -or [string]::IsNullOrWhiteSpace($trusteeUpn)) {
            throw 'SharedMailboxIdentity and TrusteeUserPrincipalName are required.'
        }

        $grantFullAccess = ConvertTo-Bool -Value $row.FullAccess -Default $false
        $grantReadOnly = ConvertTo-Bool -Value $row.ReadOnly -Default $false
        $grantSendAs = ConvertTo-Bool -Value $row.SendAs -Default $false
        $grantSendOnBehalf = ConvertTo-Bool -Value $row.SendOnBehalf -Default $false
        $autoMapping = ConvertTo-Bool -Value $row.AutoMapping -Default $true

        if (-not ($grantFullAccess -or $grantReadOnly -or $grantSendAs -or $grantSendOnBehalf)) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeUpn" -Action 'SetMailboxPermission' -Status 'Skipped' -Message 'No permissions requested for this record.'))
            $rowNumber++
            continue
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup shared mailbox $mailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $mailboxIdentity -ErrorAction Stop
        }
        $trustee = Invoke-WithRetry -OperationName "Lookup trustee $trusteeUpn" -ScriptBlock {
            Get-Recipient -Identity $trusteeUpn -ErrorAction Stop
        }

        $messages = [System.Collections.Generic.List[string]]::new()
        $rowHadError = $false

        if ($grantFullAccess) {
            try {
                $existingFullAccess = Invoke-WithRetry -OperationName "Check FullAccess $mailboxIdentity -> $trusteeUpn" -ScriptBlock {
                    Get-MailboxPermission -Identity $mailbox.Identity -User $trustee.Identity -ErrorAction SilentlyContinue |
                        Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.Deny } |
                        Select-Object -First 1
                }

                if ($existingFullAccess) {
                    $messages.Add('FullAccess: already present (skipped).')
                }
                elseif ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeUpn", 'Grant FullAccess')) {
                    Invoke-WithRetry -OperationName "Grant FullAccess $mailboxIdentity -> $trusteeUpn" -ScriptBlock {
                        Add-MailboxPermission -Identity $mailbox.Identity -User $trustee.Identity -AccessRights FullAccess -InheritanceType All -AutoMapping:$autoMapping -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    $messages.Add('FullAccess: granted.')
                }
                else {
                    $messages.Add('FullAccess: skipped due to WhatIf.')
                }
            }
            catch {
                $rowHadError = $true
                $messages.Add("FullAccess: failed ($($_.Exception.Message)).")
            }
        }

        if ($grantReadOnly) {
            try {
                if ($grantFullAccess) {
                    $messages.Add('ReadOnly: skipped because FullAccess is also requested.')
                }
                else {
                    $existingReadOnly = Invoke-WithRetry -OperationName "Check ReadOnly $mailboxIdentity -> $trusteeUpn" -ScriptBlock {
                        Get-MailboxPermission -Identity $mailbox.Identity -User $trustee.Identity -ErrorAction SilentlyContinue |
                            Where-Object { $_.AccessRights -contains 'ReadPermission' -and -not $_.Deny } |
                            Select-Object -First 1
                    }

                    if ($existingReadOnly) {
                        $messages.Add('ReadOnly: already present (skipped).')
                    }
                    elseif ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeUpn", 'Grant ReadOnly')) {
                        Invoke-WithRetry -OperationName "Grant ReadOnly $mailboxIdentity -> $trusteeUpn" -ScriptBlock {
                            Add-MailboxPermission -Identity $mailbox.Identity -User $trustee.Identity -AccessRights ReadPermission -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                        }
                        $messages.Add('ReadOnly: granted.')
                    }
                    else {
                        $messages.Add('ReadOnly: skipped due to WhatIf.')
                    }
                }
            }
            catch {
                $rowHadError = $true
                $messages.Add("ReadOnly: failed ($($_.Exception.Message)).")
            }
        }

        if ($grantSendAs) {
            try {
                $existingSendAs = Invoke-WithRetry -OperationName "Check SendAs $mailboxIdentity -> $trusteeUpn" -ScriptBlock {
                    Get-RecipientPermission -Identity $mailbox.Identity -Trustee $trustee.Identity -ErrorAction SilentlyContinue |
                        Where-Object { $_.AccessRights -contains 'SendAs' -and -not $_.Deny } |
                        Select-Object -First 1
                }

                if ($existingSendAs) {
                    $messages.Add('SendAs: already present (skipped).')
                }
                elseif ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeUpn", 'Grant SendAs')) {
                    Invoke-WithRetry -OperationName "Grant SendAs $mailboxIdentity -> $trusteeUpn" -ScriptBlock {
                        Add-RecipientPermission -Identity $mailbox.Identity -Trustee $trustee.Identity -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    $messages.Add('SendAs: granted.')
                }
                else {
                    $messages.Add('SendAs: skipped due to WhatIf.')
                }
            }
            catch {
                $rowHadError = $true
                $messages.Add("SendAs: failed ($($_.Exception.Message)).")
            }
        }

        if ($grantSendOnBehalf) {
            try {
                $existingSOB = @($mailbox.GrantSendOnBehalfTo | Where-Object { $_.DistinguishedName -eq $trustee.DistinguishedName })

                if ($existingSOB.Count -gt 0) {
                    $messages.Add('SendOnBehalf: already present (skipped).')
                }
                elseif ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeUpn", 'Grant SendOnBehalf')) {
                    Invoke-WithRetry -OperationName "Grant SendOnBehalf $mailboxIdentity -> $trusteeUpn" -ScriptBlock {
                        Set-Mailbox -Identity $mailbox.Identity -GrantSendOnBehalfTo @{ Add = $trustee.DistinguishedName } -ErrorAction Stop
                    }
                    $messages.Add('SendOnBehalf: granted.')
                }
                else {
                    $messages.Add('SendOnBehalf: skipped due to WhatIf.')
                }
            }
            catch {
                $rowHadError = $true
                $messages.Add("SendOnBehalf: failed ($($_.Exception.Message)).")
            }
        }

        $status = if ($rowHadError) { 'CompletedWithErrors' } else { 'Completed' }
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeUpn" -Action 'SetMailboxPermission' -Status $status -Message ($messages -join ' ')))
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity|$trusteeUpn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeUpn" -Action 'SetMailboxPermission' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Shared mailbox permission assignment script completed.' -Level SUCCESS

