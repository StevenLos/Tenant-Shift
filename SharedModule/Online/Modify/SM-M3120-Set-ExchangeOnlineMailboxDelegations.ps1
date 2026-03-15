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
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3120-Set-ExchangeOnlineMailboxDelegations_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {


$requiredHeaders = @(
    'MailboxIdentity',
    'TrusteeUserPrincipalName',
    'FullAccess',
    'SendAs',
    'SendOnBehalf',
    'AutoMapping'
)

Write-Status -Message 'Starting Exchange Online mailbox delegation assignment script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = ([string]$row.MailboxIdentity).Trim()
    $trusteeUpn = ([string]$row.TrusteeUserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity) -or [string]::IsNullOrWhiteSpace($trusteeUpn)) {
            throw 'MailboxIdentity and TrusteeUserPrincipalName are required.'
        }

        $grantFullAccess = ConvertTo-Bool -Value $row.FullAccess -Default $false
        $grantSendAs = ConvertTo-Bool -Value $row.SendAs -Default $false
        $grantSendOnBehalf = ConvertTo-Bool -Value $row.SendOnBehalf -Default $false
        $autoMapping = ConvertTo-Bool -Value $row.AutoMapping -Default $true

        if (-not ($grantFullAccess -or $grantSendAs -or $grantSendOnBehalf)) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeUpn" -Action 'SetMailboxDelegation' -Status 'Skipped' -Message 'No delegations requested for this record.'))
            $rowNumber++
            continue
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentity" -ScriptBlock {
            Get-ExchangeOnlineMailbox -Identity $mailboxIdentity -ErrorAction Stop
        }
        $trustee = Invoke-WithRetry -OperationName "Lookup trustee $trusteeUpn" -ScriptBlock {
            Get-ExchangeOnlineRecipient -Identity $trusteeUpn -ErrorAction Stop
        }

        $messages = [System.Collections.Generic.List[string]]::new()
        $rowHadError = $false

        if ($grantFullAccess) {
            try {
                $existingFullAccess = Invoke-WithRetry -OperationName "Check FullAccess $mailboxIdentity -> $trusteeUpn" -ScriptBlock {
                    Get-ExchangeOnlineMailboxPermission -Identity $mailbox.Identity -User $trustee.Identity -ErrorAction SilentlyContinue |
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

        if ($grantSendAs) {
            try {
                $existingSendAs = Invoke-WithRetry -OperationName "Check SendAs $mailboxIdentity -> $trusteeUpn" -ScriptBlock {
                    Get-ExchangeOnlineRecipientPermission -Identity $mailbox.Identity -Trustee $trustee.Identity -ErrorAction SilentlyContinue |
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
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeUpn" -Action 'SetMailboxDelegation' -Status $status -Message ($messages -join ' ')))
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity|$trusteeUpn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeUpn" -Action 'SetMailboxDelegation' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mailbox delegation assignment script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}









