<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-181500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineMailboxSizes and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineMailboxSizes from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-D3127-Get-ExchangeOnlineMailboxSizes.ps1 -InputCsvPath .\3127.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D3127-Get-ExchangeOnlineMailboxSizes.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-EXOL-0230-Get-ExchangeOnlineMailboxSizes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

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

function Get-ObjectPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($property) {
        return $property.Value
    }

    return $null
}

function Get-StringPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    return Get-TrimmedValue -Value (Get-ObjectPropertyValue -InputObject $InputObject -PropertyName $PropertyName)
}

function Get-MailboxStatisticsIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox
    )

    $primarySmtpAddress = Get-StringPropertyValue -InputObject $Mailbox -PropertyName 'PrimarySmtpAddress'
    if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
        return $primarySmtpAddress
    }

    return Get-StringPropertyValue -InputObject $Mailbox -PropertyName 'Identity'
}

$requiredHeaders = @(
    'MailboxIdentity'
)

$mailboxProperties = @(
    'ArchiveStatus',
    'ProhibitSendQuota',
    'ProhibitSendReceiveQuota',
    'UseDatabaseQuotaDefaults'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'MailboxIdentity',
    'DisplayName',
    'PrimarySmtpAddress',
    'RecipientTypeDetails',
    'ArchiveStatus',
    'MainMailboxItemCount',
    'MainMailboxTotalSizeBytes',
    'MainMailboxTotalSizeMB',
    'ArchiveMailboxItemCount',
    'ArchiveMailboxTotalSizeBytes',
    'ArchiveMailboxTotalSizeMB',
    'TotalMailboxItemCount',
    'TotalMailboxSizeBytes',
    'TotalMailboxSizeMB',
    'ProhibitSendQuota',
    'ProhibitSendReceiveQuota',
    'UseDatabaseQuotaDefaults'
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
                Get-ExchangeOnlineMailbox -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox -ResultSize Unlimited -Properties $mailboxProperties -ErrorAction Stop
            })
        }
        else {
            $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentityInput" -ScriptBlock {
                Get-ExchangeOnlineMailbox -Identity $mailboxIdentityInput -Properties $mailboxProperties -ErrorAction SilentlyContinue
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
            $mailboxIdentityResolved = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'Identity'
            $mailboxStatisticsIdentity = Get-MailboxStatisticsIdentity -Mailbox $mailbox
            if ([string]::IsNullOrWhiteSpace($mailboxStatisticsIdentity)) {
                throw 'Unable to resolve a unique mailbox identity for mailbox statistics lookup.'
            }

            $mainStats = Invoke-WithRetry -OperationName "Get mailbox statistics for $mailboxStatisticsIdentity" -ScriptBlock {
                Get-ExchangeOnlineMailboxStatistics -Identity $mailboxStatisticsIdentity -ErrorAction Stop
            }

            $mainSizeBytes = Convert-ToByteCount -Size $mainStats.TotalItemSize
            $mainItemCount = [int64]$mainStats.ItemCount

            $archiveSizeBytes = [int64]0
            $archiveItemCount = [int64]0

            $archiveStatus = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'ArchiveStatus'
            if ($archiveStatus -eq 'Active') {
                $archiveStats = Invoke-WithRetry -OperationName "Get archive mailbox statistics for $mailboxStatisticsIdentity" -ScriptBlock {
                    Get-ExchangeOnlineMailboxStatistics -Identity $mailboxStatisticsIdentity -Archive -ErrorAction SilentlyContinue
                }

                if ($archiveStats) {
                    $archiveSizeBytes = Convert-ToByteCount -Size $archiveStats.TotalItemSize
                    $archiveItemCount = [int64]$archiveStats.ItemCount
                }
            }

            $totalSizeBytes = $mainSizeBytes + $archiveSizeBytes
            $totalItemCount = $mainItemCount + $archiveItemCount
            if ([string]::IsNullOrWhiteSpace($mailboxIdentityResolved)) {
                $mailboxIdentityResolved = $mailboxStatisticsIdentity
            }

            $useDatabaseQuotaDefaultsValue = Get-ObjectPropertyValue -InputObject $mailbox -PropertyName 'UseDatabaseQuotaDefaults'
            $useDatabaseQuotaDefaults = if ($null -eq $useDatabaseQuotaDefaultsValue -or [string]::IsNullOrWhiteSpace(([string]$useDatabaseQuotaDefaultsValue).Trim())) {
                ''
            }
            else {
                [string][bool]$useDatabaseQuotaDefaultsValue
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxStatisticsIdentity -Action 'GetExchangeMailboxSizes' -Status 'Completed' -Message 'Mailbox size row exported.' -Data ([ordered]@{
                        MailboxIdentity               = $mailboxIdentityResolved
                        DisplayName                   = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'DisplayName'
                        PrimarySmtpAddress            = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'PrimarySmtpAddress'
                        RecipientTypeDetails          = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'RecipientTypeDetails'
                        ArchiveStatus                 = $archiveStatus
                        MainMailboxItemCount          = [string]$mainItemCount
                        MainMailboxTotalSizeBytes     = [string]$mainSizeBytes
                        MainMailboxTotalSizeMB        = Convert-BytesToMb -Bytes $mainSizeBytes
                        ArchiveMailboxItemCount       = [string]$archiveItemCount
                        ArchiveMailboxTotalSizeBytes  = [string]$archiveSizeBytes
                        ArchiveMailboxTotalSizeMB     = Convert-BytesToMb -Bytes $archiveSizeBytes
                        TotalMailboxItemCount         = [string]$totalItemCount
                        TotalMailboxSizeBytes         = [string]$totalSizeBytes
                        TotalMailboxSizeMB            = Convert-BytesToMb -Bytes $totalSizeBytes
                        ProhibitSendQuota             = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'ProhibitSendQuota'
                        ProhibitSendReceiveQuota      = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'ProhibitSendReceiveQuota'
                        UseDatabaseQuotaDefaults      = $useDatabaseQuotaDefaults
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

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mailbox-size inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
