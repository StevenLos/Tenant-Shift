<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-172500

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR3128-Get-ExchangeOnlineMailboxStatsPerMailbox_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Convert-ToIsoDateTimeString {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    try {
        return ([datetime]$Value).ToString('o')
    }
    catch {
        return ([string]$Value).Trim()
    }
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

Write-Status -Message 'Starting Exchange Online mailbox per-mailbox statistics inventory script.'
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
            $mailboxes = @(Invoke-WithRetry -OperationName 'Load all user/shared/resource mailboxes for mailbox stats inventory' -ScriptBlock {
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
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityInput -Action 'GetExchangeMailboxStatsPerMailbox' -Status 'NotFound' -Message 'No matching mailboxes were found.' -Data ([ordered]@{
                        MailboxIdentity                = $mailboxIdentityInput
                        DisplayName                    = ''
                        PrimarySmtpAddress             = ''
                        RecipientTypeDetails           = ''
                        ItemCount                      = ''
                        AssociatedItemCount            = ''
                        DeletedItemCount               = ''
                        TotalItemSizeBytes             = ''
                        TotalDeletedItemSizeBytes      = ''
                        StorageLimitStatus             = ''
                        LastLogonTime                  = ''
                        LastUserActionTime             = ''
                        IsQuarantined                  = ''
                        DisconnectDate                 = ''
                        ArchiveItemCount               = ''
                        ArchiveTotalItemSizeBytes      = ''
                        ArchiveStorageLimitStatus      = ''
                        ArchiveLastLogonTime           = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $mainStats = Invoke-WithRetry -OperationName "Get mailbox statistics for $($mailbox.Identity)" -ScriptBlock {
                Get-ExchangeOnlineMailboxStatistics -Identity $mailbox.Identity -ErrorAction Stop
            }

            $archiveItemCount = [int64]0
            $archiveTotalItemSizeBytes = [int64]0
            $archiveStorageLimitStatus = ''
            $archiveLastLogonTime = ''

            if (([string]$mailbox.ArchiveStatus).Trim() -eq 'Active') {
                $archiveStats = Invoke-WithRetry -OperationName "Get archive mailbox statistics for $($mailbox.Identity)" -ScriptBlock {
                    Get-ExchangeOnlineMailboxStatistics -Identity $mailbox.Identity -Archive -ErrorAction SilentlyContinue
                }

                if ($archiveStats) {
                    $archiveItemCount = [int64]$archiveStats.ItemCount
                    $archiveTotalItemSizeBytes = Convert-ToByteCount -Size $archiveStats.TotalItemSize
                    $archiveStorageLimitStatus = Get-TrimmedValue -Value $archiveStats.StorageLimitStatus
                    $archiveLastLogonTime = Convert-ToIsoDateTimeString -Value $archiveStats.LastLogonTime
                }
            }

            $mailboxIdentityResolved = ([string]$mailbox.Identity).Trim()
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityResolved -Action 'GetExchangeMailboxStatsPerMailbox' -Status 'Completed' -Message 'Mailbox statistics row exported.' -Data ([ordered]@{
                        MailboxIdentity                = $mailboxIdentityResolved
                        DisplayName                    = ([string]$mailbox.DisplayName).Trim()
                        PrimarySmtpAddress             = ([string]$mailbox.PrimarySmtpAddress).Trim()
                        RecipientTypeDetails           = ([string]$mailbox.RecipientTypeDetails).Trim()
                        ItemCount                      = [string][int64]$mainStats.ItemCount
                        AssociatedItemCount            = [string][int64]$mainStats.AssociatedItemCount
                        DeletedItemCount               = [string][int64]$mainStats.DeletedItemCount
                        TotalItemSizeBytes             = [string](Convert-ToByteCount -Size $mainStats.TotalItemSize)
                        TotalDeletedItemSizeBytes      = [string](Convert-ToByteCount -Size $mainStats.TotalDeletedItemSize)
                        StorageLimitStatus             = Get-TrimmedValue -Value $mainStats.StorageLimitStatus
                        LastLogonTime                  = Convert-ToIsoDateTimeString -Value $mainStats.LastLogonTime
                        LastUserActionTime             = Convert-ToIsoDateTimeString -Value $mainStats.LastUserActionTime
                        IsQuarantined                  = [string][bool]$mainStats.IsQuarantined
                        DisconnectDate                 = Convert-ToIsoDateTimeString -Value $mainStats.DisconnectDate
                        ArchiveItemCount               = [string]$archiveItemCount
                        ArchiveTotalItemSizeBytes      = [string]$archiveTotalItemSizeBytes
                        ArchiveStorageLimitStatus      = $archiveStorageLimitStatus
                        ArchiveLastLogonTime           = $archiveLastLogonTime
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentityInput) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityInput -Action 'GetExchangeMailboxStatsPerMailbox' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    MailboxIdentity                = $mailboxIdentityInput
                    DisplayName                    = ''
                    PrimarySmtpAddress             = ''
                    RecipientTypeDetails           = ''
                    ItemCount                      = ''
                    AssociatedItemCount            = ''
                    DeletedItemCount               = ''
                    TotalItemSizeBytes             = ''
                    TotalDeletedItemSizeBytes      = ''
                    StorageLimitStatus             = ''
                    LastLogonTime                  = ''
                    LastUserActionTime             = ''
                    IsQuarantined                  = ''
                    DisconnectDate                 = ''
                    ArchiveItemCount               = ''
                    ArchiveTotalItemSizeBytes      = ''
                    ArchiveStorageLimitStatus      = ''
                    ArchiveLastLogonTime           = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mailbox per-mailbox statistics inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
