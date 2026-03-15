<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-172000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR3127-Get-ExchangeOnlineMailboxSizes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-TrimmedValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    return ([string]$Value).Trim()
}

function Convert-ToByteCount {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Size
    )

    if ($null -eq $Size) {
        return [int64]0
    }

    $text = ([string]$Size).Trim()
    if ($text -match '\((?<bytes>[\d,]+)\sbytes\)') {
        return [int64](($matches['bytes']) -replace ',', '')
    }

    try {
        return [int64]$Size.Value.ToBytes()
    }
    catch {
        return [int64]0
    }
}

function Convert-BytesToMb {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int64]$Bytes
    )

    if ($Bytes -le 0) {
        return '0'
    }

    return [string][math]::Round(($Bytes / 1MB), 2)
}

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

Write-Status -Message 'Starting Exchange Online mailbox-size inventory script.'
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
    $mailboxIdentityInput = Get-TrimmedValue -Value $row.MailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentityInput)) {
            throw 'MailboxIdentity is required. Use * to inventory all user/shared/resource mailboxes.'
        }

        $mailboxes = @()
        if ($mailboxIdentityInput -eq '*') {
            $mailboxes = @(Invoke-WithRetry -OperationName 'Load all user/shared/resource mailboxes for size inventory' -ScriptBlock {
                Get-ExchangeOnlineMailbox -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox -ResultSize Unlimited -ErrorAction Stop
            })
        }
        else {
            $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentityInput" -ScriptBlock {
                Get-ExchangeOnlineMailbox -Identity $mailboxIdentityInput -ErrorAction SilentlyContinue
            }

            if ($mailbox) {
                $mailboxes = @($mailbox)
            }
        }

        if ($mailboxes.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityInput -Action 'GetExchangeMailboxSizes' -Status 'NotFound' -Message 'No matching mailboxes were found.' -Data ([ordered]@{
                        MailboxIdentity               = $mailboxIdentityInput
                        DisplayName                   = ''
                        PrimarySmtpAddress            = ''
                        RecipientTypeDetails          = ''
                        ArchiveStatus                 = ''
                        MainMailboxItemCount          = ''
                        MainMailboxTotalSizeBytes     = ''
                        MainMailboxTotalSizeMB        = ''
                        ArchiveMailboxItemCount       = ''
                        ArchiveMailboxTotalSizeBytes  = ''
                        ArchiveMailboxTotalSizeMB     = ''
                        TotalMailboxItemCount         = ''
                        TotalMailboxSizeBytes         = ''
                        TotalMailboxSizeMB            = ''
                        ProhibitSendQuota             = ''
                        ProhibitSendReceiveQuota      = ''
                        UseDatabaseQuotaDefaults      = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $mainStats = Invoke-WithRetry -OperationName "Get mailbox statistics for $($mailbox.Identity)" -ScriptBlock {
                Get-ExchangeOnlineMailboxStatistics -Identity $mailbox.Identity -ErrorAction Stop
            }

            $mainSizeBytes = Convert-ToByteCount -Size $mainStats.TotalItemSize
            $mainItemCount = [int64]$mainStats.ItemCount

            $archiveSizeBytes = [int64]0
            $archiveItemCount = [int64]0

            if (([string]$mailbox.ArchiveStatus).Trim() -eq 'Active') {
                $archiveStats = Invoke-WithRetry -OperationName "Get archive mailbox statistics for $($mailbox.Identity)" -ScriptBlock {
                    Get-ExchangeOnlineMailboxStatistics -Identity $mailbox.Identity -Archive -ErrorAction SilentlyContinue
                }

                if ($archiveStats) {
                    $archiveSizeBytes = Convert-ToByteCount -Size $archiveStats.TotalItemSize
                    $archiveItemCount = [int64]$archiveStats.ItemCount
                }
            }

            $totalSizeBytes = $mainSizeBytes + $archiveSizeBytes
            $totalItemCount = $mainItemCount + $archiveItemCount
            $mailboxIdentityResolved = ([string]$mailbox.Identity).Trim()

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityResolved -Action 'GetExchangeMailboxSizes' -Status 'Completed' -Message 'Mailbox size row exported.' -Data ([ordered]@{
                        MailboxIdentity               = $mailboxIdentityResolved
                        DisplayName                   = ([string]$mailbox.DisplayName).Trim()
                        PrimarySmtpAddress            = ([string]$mailbox.PrimarySmtpAddress).Trim()
                        RecipientTypeDetails          = ([string]$mailbox.RecipientTypeDetails).Trim()
                        ArchiveStatus                 = ([string]$mailbox.ArchiveStatus).Trim()
                        MainMailboxItemCount          = [string]$mainItemCount
                        MainMailboxTotalSizeBytes     = [string]$mainSizeBytes
                        MainMailboxTotalSizeMB        = Convert-BytesToMb -Bytes $mainSizeBytes
                        ArchiveMailboxItemCount       = [string]$archiveItemCount
                        ArchiveMailboxTotalSizeBytes  = [string]$archiveSizeBytes
                        ArchiveMailboxTotalSizeMB     = Convert-BytesToMb -Bytes $archiveSizeBytes
                        TotalMailboxItemCount         = [string]$totalItemCount
                        TotalMailboxSizeBytes         = [string]$totalSizeBytes
                        TotalMailboxSizeMB            = Convert-BytesToMb -Bytes $totalSizeBytes
                        ProhibitSendQuota             = Get-TrimmedValue -Value $mailbox.ProhibitSendQuota
                        ProhibitSendReceiveQuota      = Get-TrimmedValue -Value $mailbox.ProhibitSendReceiveQuota
                        UseDatabaseQuotaDefaults      = [string][bool]$mailbox.UseDatabaseQuotaDefaults
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentityInput) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityInput -Action 'GetExchangeMailboxSizes' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    MailboxIdentity               = $mailboxIdentityInput
                    DisplayName                   = ''
                    PrimarySmtpAddress            = ''
                    RecipientTypeDetails          = ''
                    ArchiveStatus                 = ''
                    MainMailboxItemCount          = ''
                    MainMailboxTotalSizeBytes     = ''
                    MainMailboxTotalSizeMB        = ''
                    ArchiveMailboxItemCount       = ''
                    ArchiveMailboxTotalSizeBytes  = ''
                    ArchiveMailboxTotalSizeMB     = ''
                    TotalMailboxItemCount         = ''
                    TotalMailboxSizeBytes         = ''
                    TotalMailboxSizeMB            = ''
                    ProhibitSendQuota             = ''
                    ProhibitSendReceiveQuota      = ''
                    UseDatabaseQuotaDefaults      = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mailbox-size inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
