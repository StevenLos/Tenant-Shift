<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-181000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineMailboxHighLevelStats and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineMailboxHighLevelStats from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-D3126-Get-ExchangeOnlineMailboxHighLevelStats.ps1 -InputCsvPath .\3126.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D3126-Get-ExchangeOnlineMailboxHighLevelStats.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-EXOL-0220-Get-ExchangeOnlineMailboxHighLevelStats_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-MedianInt64 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int64[]]$Values
    )

    if ($Values.Count -eq 0) {
        return [int64]0
    }

    $sorted = @($Values | Sort-Object)
    $mid = [int]([math]::Floor($sorted.Count / 2))

    if (($sorted.Count % 2) -eq 1) {
        return [int64]$sorted[$mid]
    }

    return [int64][math]::Round((([double]$sorted[$mid - 1] + [double]$sorted[$mid]) / 2), 0)
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
    'DisplayName',
    'PrimarySmtpAddress',
    'ArchiveStatus'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'MailboxIdentityInput',
    'MailboxCount',
    'MainMailboxTotalSizeBytes',
    'MainMailboxItemCount',
    'ArchiveMailboxTotalSizeBytes',
    'ArchiveMailboxItemCount',
    'TotalMailboxSizeBytes',
    'TotalMailboxItemCount',
    'AverageMailboxSizeBytes',
    'AverageMailboxItemCount',
    'MedianMailboxSizeBytes',
    'MedianMailboxItemCount',
    'Top5PercentMailboxCount',
    'Top5PercentMinSizeBytes',
    'Top5PercentMinItemCount',
    'LargestMailboxBySizeIdentity',
    'LargestMailboxBySizeBytes',
    'LargestMailboxByItemCountIdentity',
    'LargestMailboxByItemCount'
)

Write-Status -Message 'Starting Exchange Online mailbox high-level statistics inventory script.'
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
            throw 'MailboxIdentity is required. Use * to include all user/shared/resource mailboxes.'
        }

        $mailboxes = @()
        if ($mailboxIdentityInput -eq '*') {
            $mailboxes = @(Invoke-WithRetry -OperationName 'Load all user/shared/resource mailboxes' -ScriptBlock {
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
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityInput -Action 'GetExchangeMailboxHighLevelStats' -Status 'NotFound' -Message 'No matching mailboxes were found.' -Data ([ordered]@{
                        MailboxIdentityInput                = $mailboxIdentityInput
                        MailboxCount                        = ''
                        MainMailboxTotalSizeBytes           = ''
                        MainMailboxItemCount                = ''
                        ArchiveMailboxTotalSizeBytes        = ''
                        ArchiveMailboxItemCount             = ''
                        TotalMailboxSizeBytes               = ''
                        TotalMailboxItemCount               = ''
                        AverageMailboxSizeBytes             = ''
                        AverageMailboxItemCount             = ''
                        MedianMailboxSizeBytes              = ''
                        MedianMailboxItemCount              = ''
                        Top5PercentMailboxCount             = ''
                        Top5PercentMinSizeBytes             = ''
                        Top5PercentMinItemCount             = ''
                        LargestMailboxBySizeIdentity        = ''
                        LargestMailboxBySizeBytes           = ''
                        LargestMailboxByItemCountIdentity   = ''
                        LargestMailboxByItemCount           = ''
                    })))
            $rowNumber++
            continue
        }

        $mailboxStatsList = [System.Collections.Generic.List[object]]::new()

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

            if ((Get-StringPropertyValue -InputObject $mailbox -PropertyName 'ArchiveStatus') -eq 'Active') {
                $archiveStats = Invoke-WithRetry -OperationName "Get archive mailbox statistics for $mailboxStatisticsIdentity" -ScriptBlock {
                    Get-ExchangeOnlineMailboxStatistics -Identity $mailboxStatisticsIdentity -Archive -ErrorAction SilentlyContinue
                }

                if ($archiveStats) {
                    $archiveSizeBytes = Convert-ToByteCount -Size $archiveStats.TotalItemSize
                    $archiveItemCount = [int64]$archiveStats.ItemCount
                }
            }

            $mailboxStatsList.Add([PSCustomObject]@{
                    Identity       = $mailboxStatisticsIdentity
                    TotalSizeBytes = $mainSizeBytes + $archiveSizeBytes
                    TotalItemCount = $mainItemCount + $archiveItemCount
                    MainSizeBytes  = $mainSizeBytes
                    MainItemCount  = $mainItemCount
                    ArchiveSize    = $archiveSizeBytes
                    ArchiveItems   = $archiveItemCount
                })
        }

        $mailboxCount = [int64]$mailboxStatsList.Count
        $mainMailboxTotalSizeBytes = [int64](($mailboxStatsList | Measure-Object -Property MainSizeBytes -Sum).Sum)
        $mainMailboxItemCount = [int64](($mailboxStatsList | Measure-Object -Property MainItemCount -Sum).Sum)
        $archiveMailboxTotalSizeBytes = [int64](($mailboxStatsList | Measure-Object -Property ArchiveSize -Sum).Sum)
        $archiveMailboxItemCount = [int64](($mailboxStatsList | Measure-Object -Property ArchiveItems -Sum).Sum)
        $totalMailboxSizeBytes = [int64](($mailboxStatsList | Measure-Object -Property TotalSizeBytes -Sum).Sum)
        $totalMailboxItemCount = [int64](($mailboxStatsList | Measure-Object -Property TotalItemCount -Sum).Sum)

        $averageMailboxSizeBytes = [int64]0
        $averageMailboxItemCount = [int64]0
        if ($mailboxCount -gt 0) {
            $averageMailboxSizeBytes = [int64][math]::Round(($totalMailboxSizeBytes / [double]$mailboxCount), 0)
            $averageMailboxItemCount = [int64][math]::Round(($totalMailboxItemCount / [double]$mailboxCount), 0)
        }

        $sizeValues = @($mailboxStatsList | ForEach-Object { [int64]$_.TotalSizeBytes })
        $itemValues = @($mailboxStatsList | ForEach-Object { [int64]$_.TotalItemCount })
        $medianMailboxSizeBytes = Get-MedianInt64 -Values $sizeValues
        $medianMailboxItemCount = Get-MedianInt64 -Values $itemValues

        $top5PercentMailboxCount = [int64]0
        $top5PercentMinSizeBytes = [int64]0
        $top5PercentMinItemCount = [int64]0

        if ($mailboxCount -gt 0) {
            $top5PercentMailboxCount = [int64][math]::Ceiling($mailboxCount * 0.05)
            if ($top5PercentMailboxCount -lt 1) {
                $top5PercentMailboxCount = 1
            }

            $topBySize = @($mailboxStatsList | Sort-Object -Property TotalSizeBytes -Descending | Select-Object -First $top5PercentMailboxCount)
            $topByItems = @($mailboxStatsList | Sort-Object -Property TotalItemCount -Descending | Select-Object -First $top5PercentMailboxCount)

            if ($topBySize.Count -gt 0) {
                $top5PercentMinSizeBytes = [int64](($topBySize | Measure-Object -Property TotalSizeBytes -Minimum).Minimum)
            }
            if ($topByItems.Count -gt 0) {
                $top5PercentMinItemCount = [int64](($topByItems | Measure-Object -Property TotalItemCount -Minimum).Minimum)
            }
        }

        $largestMailboxBySize = @($mailboxStatsList | Sort-Object -Property TotalSizeBytes -Descending | Select-Object -First 1)
        $largestMailboxByItemCount = @($mailboxStatsList | Sort-Object -Property TotalItemCount -Descending | Select-Object -First 1)

        $largestMailboxBySizeIdentity = if ($largestMailboxBySize.Count -gt 0) { [string]$largestMailboxBySize[0].Identity } else { '' }
        $largestMailboxBySizeBytes = if ($largestMailboxBySize.Count -gt 0) { [int64]$largestMailboxBySize[0].TotalSizeBytes } else { [int64]0 }
        $largestMailboxByItemCountIdentity = if ($largestMailboxByItemCount.Count -gt 0) { [string]$largestMailboxByItemCount[0].Identity } else { '' }
        $largestMailboxByItemCount = if ($largestMailboxByItemCount.Count -gt 0) { [int64]$largestMailboxByItemCount[0].TotalItemCount } else { [int64]0 }

        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityInput -Action 'GetExchangeMailboxHighLevelStats' -Status 'Completed' -Message 'Mailbox high-level statistics exported.' -Data ([ordered]@{
                    MailboxIdentityInput                = $mailboxIdentityInput
                    MailboxCount                        = [string]$mailboxCount
                    MainMailboxTotalSizeBytes           = [string]$mainMailboxTotalSizeBytes
                    MainMailboxItemCount                = [string]$mainMailboxItemCount
                    ArchiveMailboxTotalSizeBytes        = [string]$archiveMailboxTotalSizeBytes
                    ArchiveMailboxItemCount             = [string]$archiveMailboxItemCount
                    TotalMailboxSizeBytes               = [string]$totalMailboxSizeBytes
                    TotalMailboxItemCount               = [string]$totalMailboxItemCount
                    AverageMailboxSizeBytes             = [string]$averageMailboxSizeBytes
                    AverageMailboxItemCount             = [string]$averageMailboxItemCount
                    MedianMailboxSizeBytes              = [string]$medianMailboxSizeBytes
                    MedianMailboxItemCount              = [string]$medianMailboxItemCount
                    Top5PercentMailboxCount             = [string]$top5PercentMailboxCount
                    Top5PercentMinSizeBytes             = [string]$top5PercentMinSizeBytes
                    Top5PercentMinItemCount             = [string]$top5PercentMinItemCount
                    LargestMailboxBySizeIdentity        = $largestMailboxBySizeIdentity
                    LargestMailboxBySizeBytes           = [string]$largestMailboxBySizeBytes
                    LargestMailboxByItemCountIdentity   = $largestMailboxByItemCountIdentity
                    LargestMailboxByItemCount           = [string]$largestMailboxByItemCount
                })))
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentityInput) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityInput -Action 'GetExchangeMailboxHighLevelStats' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    MailboxIdentityInput                = $mailboxIdentityInput
                    MailboxCount                        = ''
                    MainMailboxTotalSizeBytes           = ''
                    MainMailboxItemCount                = ''
                    ArchiveMailboxTotalSizeBytes        = ''
                    ArchiveMailboxItemCount             = ''
                    TotalMailboxSizeBytes               = ''
                    TotalMailboxItemCount               = ''
                    AverageMailboxSizeBytes             = ''
                    AverageMailboxItemCount             = ''
                    MedianMailboxSizeBytes              = ''
                    MedianMailboxItemCount              = ''
                    Top5PercentMailboxCount             = ''
                    Top5PercentMinSizeBytes             = ''
                    Top5PercentMinItemCount             = ''
                    LargestMailboxBySizeIdentity        = ''
                    LargestMailboxBySizeBytes           = ''
                    LargestMailboxByItemCountIdentity   = ''
                    LargestMailboxByItemCount           = ''
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
Write-Status -Message 'Exchange Online mailbox high-level statistics inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
