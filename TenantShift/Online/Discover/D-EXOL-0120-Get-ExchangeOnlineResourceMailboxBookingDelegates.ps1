<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-172500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineResourceMailboxBookingDelegates and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineResourceMailboxBookingDelegates from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-D3119-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1 -InputCsvPath .\3119.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D3119-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-EXOL-0120-Get-ExchangeOnlineResourceMailboxBookingDelegates_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'ResourceMailboxIdentity'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'ResourceMailboxIdentity',
    'ResourceType',
    'ResourceDelegates',
    'AutomateProcessing',
    'ForwardRequestsToDelegates',
    'AllBookInPolicy',
    'AllRequestInPolicy',
    'AllRequestOutOfPolicy',
    'BookInPolicy',
    'RequestInPolicy',
    'RequestOutOfPolicy',
    'BookingWindowInDays',
    'MaximumDurationInMinutes',
    'AllowRecurringMeetings',
    'EnforceSchedulingHorizon',
    'ScheduleOnlyDuringWorkHours'
)

Write-Status -Message 'Starting Exchange Online resource mailbox booking delegate inventory script.'
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
    $resourceMailboxIdentity = ([string]$row.ResourceMailboxIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($resourceMailboxIdentity)) {
            throw 'ResourceMailboxIdentity is required. Use * to inventory all room/equipment mailbox booking settings.'
        }

        $mailboxes = @()
        if ($resourceMailboxIdentity -eq '*') {
            $mailboxes = @(Invoke-WithRetry -OperationName 'Load all resource mailboxes for booking settings' -ScriptBlock {
                Get-ExchangeOnlineMailbox -RecipientTypeDetails RoomMailbox,EquipmentMailbox -ResultSize Unlimited -ErrorAction Stop
            })
        }
        else {
            $mailbox = Invoke-WithRetry -OperationName "Lookup resource mailbox $resourceMailboxIdentity" -ScriptBlock {
                Get-ExchangeOnlineMailbox -Identity $resourceMailboxIdentity -ErrorAction SilentlyContinue
            }
            if ($mailbox -and ([string]$mailbox.RecipientTypeDetails).Trim() -in @('RoomMailbox', 'EquipmentMailbox')) {
                $mailboxes = @($mailbox)
            }
        }

        if ($mailboxes.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'GetExchangeResourceMailboxBookingDelegates' -Status 'NotFound' -Message 'No matching resource mailboxes were found.' -Data ([ordered]@{
                        ResourceMailboxIdentity       = $resourceMailboxIdentity
                        ResourceType                  = ''
                        ResourceDelegates             = ''
                        AutomateProcessing            = ''
                        ForwardRequestsToDelegates    = ''
                        AllBookInPolicy               = ''
                        AllRequestInPolicy            = ''
                        AllRequestOutOfPolicy         = ''
                        BookInPolicy                  = ''
                        RequestInPolicy               = ''
                        RequestOutOfPolicy            = ''
                        BookingWindowInDays           = ''
                        MaximumDurationInMinutes      = ''
                        AllowRecurringMeetings        = ''
                        EnforceSchedulingHorizon      = ''
                        ScheduleOnlyDuringWorkHours   = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, Identity)) {
            $identity = ([string]$mailbox.Identity).Trim()
            $calendar = Invoke-WithRetry -OperationName "Load calendar processing $identity" -ScriptBlock {
                Get-CalendarProcessing -Identity $mailbox.Identity -ErrorAction Stop
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeResourceMailboxBookingDelegates' -Status 'Completed' -Message 'Resource mailbox booking settings exported.' -Data ([ordered]@{
                        ResourceMailboxIdentity       = $identity
                        ResourceType                  = ([string]$mailbox.RecipientTypeDetails).Trim()
                        ResourceDelegates             = Convert-MultiValueToString -Value $calendar.ResourceDelegates
                        AutomateProcessing            = ([string]$calendar.AutomateProcessing).Trim()
                        ForwardRequestsToDelegates    = [string]$calendar.ForwardRequestsToDelegates
                        AllBookInPolicy               = [string]$calendar.AllBookInPolicy
                        AllRequestInPolicy            = [string]$calendar.AllRequestInPolicy
                        AllRequestOutOfPolicy         = [string]$calendar.AllRequestOutOfPolicy
                        BookInPolicy                  = Convert-MultiValueToString -Value $calendar.BookInPolicy
                        RequestInPolicy               = Convert-MultiValueToString -Value $calendar.RequestInPolicy
                        RequestOutOfPolicy            = Convert-MultiValueToString -Value $calendar.RequestOutOfPolicy
                        BookingWindowInDays           = [string]$calendar.BookingWindowInDays
                        MaximumDurationInMinutes      = [string]$calendar.MaximumDurationInMinutes
                        AllowRecurringMeetings        = [string]$calendar.AllowRecurringMeetings
                        EnforceSchedulingHorizon      = [string]$calendar.EnforceSchedulingHorizon
                        ScheduleOnlyDuringWorkHours   = [string]$calendar.ScheduleOnlyDuringWorkHours
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($resourceMailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'GetExchangeResourceMailboxBookingDelegates' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    ResourceMailboxIdentity       = $resourceMailboxIdentity
                    ResourceType                  = ''
                    ResourceDelegates             = ''
                    AutomateProcessing            = ''
                    ForwardRequestsToDelegates    = ''
                    AllBookInPolicy               = ''
                    AllRequestInPolicy            = ''
                    AllRequestOutOfPolicy         = ''
                    BookInPolicy                  = ''
                    RequestInPolicy               = ''
                    RequestOutOfPolicy            = ''
                    BookingWindowInDays           = ''
                    MaximumDurationInMinutes      = ''
                    AllowRecurringMeetings        = ''
                    EnforceSchedulingHorizon      = ''
                    ScheduleOnlyDuringWorkHours   = ''
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
Write-Status -Message 'Exchange Online resource mailbox booking delegate inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}









