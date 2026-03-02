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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3121-Get-ExchangeOnlineMailboxFolderPermissions_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Convert-MultiValueToString {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [string]) {
        return ([string]$Value).Trim()
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        $items = [System.Collections.Generic.List[string]]::new()
        foreach ($item in $Value) {
            $text = ([string]$item).Trim()
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $items.Add($text)
            }
        }

        return (@($items | Sort-Object -Unique) -join ';')
    }

    return ([string]$Value).Trim()
}

$requiredHeaders = @(
    'MailboxIdentity'
)

Write-Status -Message 'Starting Exchange Online mailbox folder permission inventory script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = ([string]$row.MailboxIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
            throw 'MailboxIdentity is required. Use * to inventory folder permissions for all user/shared mailboxes.'
        }

        $mailboxes = @()
        if ($mailboxIdentity -eq '*') {
            $mailboxes = @(Invoke-WithRetry -OperationName 'Load all user/shared mailboxes for folder permission inventory' -ScriptBlock {
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
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'GetExchangeMailboxFolderPermission' -Status 'NotFound' -Message 'No matching mailboxes were found.' -Data ([ordered]@{
                        MailboxIdentity            = $mailboxIdentity
                        MailboxRecipientTypeDetails= ''
                        FolderIdentity             = ''
                        FolderPath                 = ''
                        TrusteeIdentity            = ''
                        AccessRights               = ''
                        SharingPermissionFlags     = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $mailboxIdentityResolved = ([string]$mailbox.Identity).Trim()
            $folderStats = @(Invoke-WithRetry -OperationName "Load mailbox folders $mailboxIdentityResolved" -ScriptBlock {
                Get-MailboxFolderStatistics -Identity $mailboxIdentityResolved -ErrorAction Stop
            })

            $rowAdded = $false
            foreach ($folderStat in $folderStats) {
                $folderIdentity = ([string]$folderStat.Identity).Trim()
                if ([string]::IsNullOrWhiteSpace($folderIdentity)) {
                    continue
                }

                $permissions = @(Invoke-WithRetry -OperationName "Load folder permissions $folderIdentity" -ScriptBlock {
                    Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction SilentlyContinue
                })

                foreach ($permission in $permissions) {
                    $userText = ([string]$permission.User).Trim()
                    if ([string]::IsNullOrWhiteSpace($userText)) { continue }
                    if ($userText.Equals('Default', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                    if ($userText.Equals('Anonymous', [System.StringComparison]::OrdinalIgnoreCase)) { continue }

                    $rowAdded = $true
                    $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$mailboxIdentityResolved|$folderIdentity|$userText" -Action 'GetExchangeMailboxFolderPermission' -Status 'Completed' -Message 'Mailbox folder permission row exported.' -Data ([ordered]@{
                                MailboxIdentity             = $mailboxIdentityResolved
                                MailboxRecipientTypeDetails = ([string]$mailbox.RecipientTypeDetails).Trim()
                                FolderIdentity              = $folderIdentity
                                FolderPath                  = ([string]$folderStat.FolderPath).Trim()
                                TrusteeIdentity             = $userText
                                AccessRights                = Convert-MultiValueToString -Value $permission.AccessRights
                                SharingPermissionFlags      = Convert-MultiValueToString -Value $permission.SharingPermissionFlags
                            })))
                }
            }

            if (-not $rowAdded) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityResolved -Action 'GetExchangeMailboxFolderPermission' -Status 'Completed' -Message 'No explicit folder permissions found.' -Data ([ordered]@{
                            MailboxIdentity             = $mailboxIdentityResolved
                            MailboxRecipientTypeDetails = ([string]$mailbox.RecipientTypeDetails).Trim()
                            FolderIdentity              = ''
                            FolderPath                  = ''
                            TrusteeIdentity             = ''
                            AccessRights                = ''
                            SharingPermissionFlags      = ''
                        })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'GetExchangeMailboxFolderPermission' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    MailboxIdentity             = $mailboxIdentity
                    MailboxRecipientTypeDetails = ''
                    FolderIdentity              = ''
                    FolderPath                  = ''
                    TrusteeIdentity             = ''
                    AccessRights                = ''
                    SharingPermissionFlags      = ''
                })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mailbox folder permission inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





