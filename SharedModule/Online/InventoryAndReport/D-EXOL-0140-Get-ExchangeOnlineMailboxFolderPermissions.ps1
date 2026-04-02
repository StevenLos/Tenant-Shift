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

.SYNOPSIS
    Gets ExchangeOnlineMailboxFolderPermissions and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineMailboxFolderPermissions from Microsoft 365 and writes the results to a CSV file.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all objects in scope (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all objects in scope rather than processing from an input CSV file. Uses the DiscoverAll parameter set.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-IR3121-Get-ExchangeOnlineMailboxFolderPermissions.ps1 -InputCsvPath .\3121.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3121-Get-ExchangeOnlineMailboxFolderPermissions.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator
    Limitations:      None known.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-EXOL-0140-Get-ExchangeOnlineMailboxFolderPermissions_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$requiredHeaders = @(
    'MailboxIdentity'
)

Write-Status -Message 'Starting Exchange Online mailbox folder permission inventory script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$scopeMode = 'Csv'
if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. CSV input is bypassed.' -Level WARN

    $discoverRow = [ordered]@{}
    foreach ($header in $requiredHeaders) {
        $discoverRow[$header] = '*'
    }

    $rows = @([PSCustomObject]$discoverRow)
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}
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
                Get-ExchangeOnlineMailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited -ErrorAction Stop
            })
        }
        else {
            $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentity" -ScriptBlock {
                Get-ExchangeOnlineMailbox -Identity $mailboxIdentity -ErrorAction SilentlyContinue
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
                Get-ExchangeOnlineMailboxFolderStatistics -Identity $mailboxIdentityResolved -ErrorAction Stop
            })

            $rowAdded = $false
            foreach ($folderStat in $folderStats) {
                $folderIdentity = ([string]$folderStat.Identity).Trim()
                if ([string]::IsNullOrWhiteSpace($folderIdentity)) {
                    continue
                }

                $permissions = @(Invoke-WithRetry -OperationName "Load folder permissions $folderIdentity" -ScriptBlock {
                    Get-ExchangeOnlineMailboxFolderPermission -Identity $folderIdentity -ErrorAction SilentlyContinue
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

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mailbox folder permission inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}










