<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-182500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineMailboxStatsPerMailbox and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineMailboxStatsPerMailbox from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-IR3128-Get-ExchangeOnlineMailboxStatsPerMailbox.ps1 -InputCsvPath .\3128.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3128-Get-ExchangeOnlineMailboxStatsPerMailbox.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-EXOL-0240-Get-ExchangeOnlineMailboxStatsPerMailbox_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'MailboxIdentity',
    'DisplayName',
    'PrimarySmtpAddress',
    'RecipientTypeDetails',
    'ItemCount',
    'AssociatedItemCount',
    'DeletedItemCount',
    'TotalItemSizeBytes',
    'TotalDeletedItemSizeBytes',
    'StorageLimitStatus',
    'LastLogonTime',
    'LastUserActionTime',
    'IsQuarantined',
    'DisconnectDate',
    'ArchiveItemCount',
    'ArchiveTotalItemSizeBytes',
    'ArchiveStorageLimitStatus',
    'ArchiveLastLogonTime'
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
            $mailboxIdentityResolved = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'Identity'
            $mailboxStatisticsIdentity = Get-MailboxStatisticsIdentity -Mailbox $mailbox
            if ([string]::IsNullOrWhiteSpace($mailboxStatisticsIdentity)) {
                throw 'Unable to resolve a unique mailbox identity for mailbox statistics lookup.'
            }

            $mainStats = Invoke-WithRetry -OperationName "Get mailbox statistics for $mailboxStatisticsIdentity" -ScriptBlock {
                Get-ExchangeOnlineMailboxStatistics -Identity $mailboxStatisticsIdentity -ErrorAction Stop
            }

            $mainItemCount = [int64](Get-ObjectPropertyValue -InputObject $mainStats -PropertyName 'ItemCount')
            $mainAssociatedItemCount = [int64](Get-ObjectPropertyValue -InputObject $mainStats -PropertyName 'AssociatedItemCount')
            $mainDeletedItemCount = [int64](Get-ObjectPropertyValue -InputObject $mainStats -PropertyName 'DeletedItemCount')
            $mainTotalItemSizeBytes = Convert-ToByteCount -Size (Get-ObjectPropertyValue -InputObject $mainStats -PropertyName 'TotalItemSize')
            $mainTotalDeletedItemSizeBytes = Convert-ToByteCount -Size (Get-ObjectPropertyValue -InputObject $mainStats -PropertyName 'TotalDeletedItemSize')
            $mainStorageLimitStatus = Get-StringPropertyValue -InputObject $mainStats -PropertyName 'StorageLimitStatus'
            $mainLastLogonTime = Convert-ToIsoDateTimeString -Value (Get-ObjectPropertyValue -InputObject $mainStats -PropertyName 'LastLogonTime')
            $mainLastUserActionTime = Convert-ToIsoDateTimeString -Value (Get-ObjectPropertyValue -InputObject $mainStats -PropertyName 'LastUserActionTime')
            $mainDisconnectDate = Convert-ToIsoDateTimeString -Value (Get-ObjectPropertyValue -InputObject $mainStats -PropertyName 'DisconnectDate')

            $archiveItemCount = [int64]0
            $archiveTotalItemSizeBytes = [int64]0
            $archiveStorageLimitStatus = ''
            $archiveLastLogonTime = ''

            $archiveStatus = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'ArchiveStatus'
            if ($archiveStatus -eq 'Active') {
                $archiveStats = Invoke-WithRetry -OperationName "Get archive mailbox statistics for $mailboxStatisticsIdentity" -ScriptBlock {
                    Get-ExchangeOnlineMailboxStatistics -Identity $mailboxStatisticsIdentity -Archive -ErrorAction SilentlyContinue
                }

                if ($archiveStats) {
                    $archiveItemCount = [int64](Get-ObjectPropertyValue -InputObject $archiveStats -PropertyName 'ItemCount')
                    $archiveTotalItemSizeBytes = Convert-ToByteCount -Size (Get-ObjectPropertyValue -InputObject $archiveStats -PropertyName 'TotalItemSize')
                    $archiveStorageLimitStatus = Get-StringPropertyValue -InputObject $archiveStats -PropertyName 'StorageLimitStatus'
                    $archiveLastLogonTime = Convert-ToIsoDateTimeString -Value (Get-ObjectPropertyValue -InputObject $archiveStats -PropertyName 'LastLogonTime')
                }
            }

            if ([string]::IsNullOrWhiteSpace($mailboxIdentityResolved)) {
                $mailboxIdentityResolved = $mailboxStatisticsIdentity
            }

            $isQuarantinedValue = Get-ObjectPropertyValue -InputObject $mainStats -PropertyName 'IsQuarantined'
            $isQuarantined = if ($null -eq $isQuarantinedValue -or [string]::IsNullOrWhiteSpace(([string]$isQuarantinedValue).Trim())) {
                ''
            }
            else {
                [string][bool]$isQuarantinedValue
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxStatisticsIdentity -Action 'GetExchangeMailboxStatsPerMailbox' -Status 'Completed' -Message 'Mailbox statistics row exported.' -Data ([ordered]@{
                        MailboxIdentity                = $mailboxIdentityResolved
                        DisplayName                    = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'DisplayName'
                        PrimarySmtpAddress             = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'PrimarySmtpAddress'
                        RecipientTypeDetails           = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'RecipientTypeDetails'
                        ItemCount                      = [string]$mainItemCount
                        AssociatedItemCount            = [string]$mainAssociatedItemCount
                        DeletedItemCount               = [string]$mainDeletedItemCount
                        TotalItemSizeBytes             = [string]$mainTotalItemSizeBytes
                        TotalDeletedItemSizeBytes      = [string]$mainTotalDeletedItemSizeBytes
                        StorageLimitStatus             = $mainStorageLimitStatus
                        LastLogonTime                  = $mainLastLogonTime
                        LastUserActionTime             = $mainLastUserActionTime
                        IsQuarantined                  = $isQuarantined
                        DisconnectDate                 = $mainDisconnectDate
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

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mailbox per-mailbox statistics inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
