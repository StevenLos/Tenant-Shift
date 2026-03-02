<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-005957

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3120-Get-ExchangeOnlineMailboxDelegations_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Normalize-TrusteeKey {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    return $text.ToLowerInvariant()
}

$requiredHeaders = @(
    'MailboxIdentity'
)

Write-Status -Message 'Starting Exchange Online mailbox delegation inventory script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()
$recipientSummaryByKey = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

$resolveRecipientSummary = {
    param(
        [Parameter(Mandatory)]
        [string]$IdentityHint
    )

    $normalized = Normalize-TrusteeKey -Value $IdentityHint
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return [PSCustomObject]@{
            TrusteeIdentity           = ''
            TrusteePrimarySmtpAddress = ''
            TrusteeRecipientType      = ''
        }
    }

    if ($recipientSummaryByKey.ContainsKey($normalized)) {
        return $recipientSummaryByKey[$normalized]
    }

    $summary = $null
    try {
        $recipient = Invoke-WithRetry -OperationName "Lookup recipient $IdentityHint" -ScriptBlock {
            Get-Recipient -Identity $IdentityHint -ErrorAction Stop
        }

        $summary = [PSCustomObject]@{
            TrusteeIdentity           = ([string]$recipient.Identity).Trim()
            TrusteePrimarySmtpAddress = ([string]$recipient.PrimarySmtpAddress).Trim()
            TrusteeRecipientType      = ([string]$recipient.RecipientType).Trim()
        }
    }
    catch {
        $summary = [PSCustomObject]@{
            TrusteeIdentity           = $IdentityHint
            TrusteePrimarySmtpAddress = ''
            TrusteeRecipientType      = ''
        }
    }

    $recipientSummaryByKey[$normalized] = $summary
    return $summary
}

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = ([string]$row.MailboxIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
            throw 'MailboxIdentity is required. Use * to inventory delegations for all user/shared mailboxes.'
        }

        $mailboxes = @()
        if ($mailboxIdentity -eq '*') {
            $mailboxes = @(Invoke-WithRetry -OperationName 'Load all user/shared mailboxes for delegation inventory' -ScriptBlock {
                Get-Mailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited -ErrorAction Stop
            })
        }
        else {
            $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentity" -ScriptBlock {
                Get-Mailbox -Identity $mailboxIdentity -ErrorAction SilentlyContinue
            }
            if ($mailbox) {
                $mailboxes = @($mailbox)
            }
        }

        if ($mailboxes.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'GetExchangeMailboxDelegation' -Status 'NotFound' -Message 'No matching mailboxes were found.' -Data ([ordered]@{
                        MailboxIdentity             = $mailboxIdentity
                        MailboxRecipientTypeDetails = ''
                        TrusteeIdentity             = ''
                        TrusteePrimarySmtpAddress   = ''
                        TrusteeRecipientType        = ''
                        FullAccess                  = ''
                        ReadOnly                    = ''
                        SendAs                      = ''
                        SendOnBehalf                = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $mailboxIdentityResolved = ([string]$mailbox.Identity).Trim()
            $permissionMap = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

            $ensureEntry = {
                param(
                    [Parameter(Mandatory)]
                    [string]$TrusteeHint
                )

                $summary = & $resolveRecipientSummary -IdentityHint $TrusteeHint
                $key = Normalize-TrusteeKey -Value $summary.TrusteeIdentity
                if ([string]::IsNullOrWhiteSpace($key)) {
                    $key = Normalize-TrusteeKey -Value $TrusteeHint
                }
                if ([string]::IsNullOrWhiteSpace($key)) {
                    return $null
                }

                if ($permissionMap.ContainsKey($key)) {
                    return $permissionMap[$key]
                }

                $entry = [PSCustomObject]@{
                    TrusteeIdentity           = $summary.TrusteeIdentity
                    TrusteePrimarySmtpAddress = $summary.TrusteePrimarySmtpAddress
                    TrusteeRecipientType      = $summary.TrusteeRecipientType
                    FullAccess                = $false
                    ReadOnly                  = $false
                    SendAs                    = $false
                    SendOnBehalf              = $false
                }

                $permissionMap[$key] = $entry
                return $entry
            }

            $mailboxPermissions = @(Invoke-WithRetry -OperationName "Load mailbox permissions $mailboxIdentityResolved" -ScriptBlock {
                Get-MailboxPermission -Identity $mailboxIdentityResolved -ErrorAction Stop
            })

            foreach ($permission in $mailboxPermissions) {
                if ($permission.Deny -eq $true) { continue }
                if ($permission.IsInherited -eq $true) { continue }

                $trustee = ([string]$permission.User).Trim()
                if ([string]::IsNullOrWhiteSpace($trustee)) { continue }
                if ($trustee.Equals('NT AUTHORITY\\SELF', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                if ($trustee -match '^S-1-5-') { continue }

                $entry = & $ensureEntry -TrusteeHint $trustee
                if ($null -eq $entry) { continue }

                $accessRights = @($permission.AccessRights | ForEach-Object { ([string]$_).Trim() })
                if ($accessRights -contains 'FullAccess') {
                    $entry.FullAccess = $true
                }
                if ($accessRights -contains 'ReadPermission') {
                    $entry.ReadOnly = $true
                }
            }

            $recipientPermissions = @(Invoke-WithRetry -OperationName "Load recipient permissions $mailboxIdentityResolved" -ScriptBlock {
                Get-RecipientPermission -Identity $mailboxIdentityResolved -ErrorAction SilentlyContinue
            })

            foreach ($permission in $recipientPermissions) {
                if ($permission.Deny -eq $true) { continue }

                $accessRights = @($permission.AccessRights | ForEach-Object { ([string]$_).Trim() })
                if ($accessRights -notcontains 'SendAs') { continue }

                $trustee = ([string]$permission.Trustee).Trim()
                if ([string]::IsNullOrWhiteSpace($trustee)) { continue }

                $entry = & $ensureEntry -TrusteeHint $trustee
                if ($null -eq $entry) { continue }

                $entry.SendAs = $true
            }

            foreach ($delegate in @($mailbox.GrantSendOnBehalfTo)) {
                $delegateHint = ([string]$delegate.DistinguishedName).Trim()
                if ([string]::IsNullOrWhiteSpace($delegateHint)) {
                    $delegateHint = ([string]$delegate.Name).Trim()
                }
                if ([string]::IsNullOrWhiteSpace($delegateHint)) { continue }

                $entry = & $ensureEntry -TrusteeHint $delegateHint
                if ($null -eq $entry) { continue }

                $entry.SendOnBehalf = $true
            }

            if ($permissionMap.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityResolved -Action 'GetExchangeMailboxDelegation' -Status 'Completed' -Message 'No explicit delegated permissions found for mailbox.' -Data ([ordered]@{
                            MailboxIdentity             = $mailboxIdentityResolved
                            MailboxRecipientTypeDetails = ([string]$mailbox.RecipientTypeDetails).Trim()
                            TrusteeIdentity             = ''
                            TrusteePrimarySmtpAddress   = ''
                            TrusteeRecipientType        = ''
                            FullAccess                  = ''
                            ReadOnly                    = ''
                            SendAs                      = ''
                            SendOnBehalf                = ''
                        })))
                continue
            }

            foreach ($entry in @($permissionMap.Values | Sort-Object -Property TrusteeIdentity)) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$mailboxIdentityResolved|$($entry.TrusteeIdentity)" -Action 'GetExchangeMailboxDelegation' -Status 'Completed' -Message 'Mailbox delegation row exported.' -Data ([ordered]@{
                            MailboxIdentity             = $mailboxIdentityResolved
                            MailboxRecipientTypeDetails = ([string]$mailbox.RecipientTypeDetails).Trim()
                            TrusteeIdentity             = $entry.TrusteeIdentity
                            TrusteePrimarySmtpAddress   = $entry.TrusteePrimarySmtpAddress
                            TrusteeRecipientType        = $entry.TrusteeRecipientType
                            FullAccess                  = [string]$entry.FullAccess
                            ReadOnly                    = [string]$entry.ReadOnly
                            SendAs                      = [string]$entry.SendAs
                            SendOnBehalf                = [string]$entry.SendOnBehalf
                        })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'GetExchangeMailboxDelegation' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    MailboxIdentity             = $mailboxIdentity
                    MailboxRecipientTypeDetails = ''
                    TrusteeIdentity             = ''
                    TrusteePrimarySmtpAddress   = ''
                    TrusteeRecipientType        = ''
                    FullAccess                  = ''
                    ReadOnly                    = ''
                    SendAs                      = ''
                    SendOnBehalf                = ''
                })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mailbox delegation inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





