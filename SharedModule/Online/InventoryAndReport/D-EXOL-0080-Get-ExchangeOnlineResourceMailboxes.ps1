<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-172000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineResourceMailboxes and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineResourceMailboxes from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-IR3118-Get-ExchangeOnlineResourceMailboxes.ps1 -InputCsvPath .\3118.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3118-Get-ExchangeOnlineResourceMailboxes.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-EXOL-0080-Get-ExchangeOnlineResourceMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

    return ([string](Get-ObjectPropertyValue -InputObject $InputObject -PropertyName $PropertyName)).Trim()
}

$requiredHeaders = @(
    'ResourceMailboxIdentity'
)

$mailboxProperties = @(
    'RecipientTypeDetails',
    'DisplayName',
    'Alias',
    'PrimarySmtpAddress',
    'UserPrincipalName',
    'HiddenFromAddressListsEnabled',
    'ResourceCapacity'
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
    'DisplayName',
    'Alias',
    'PrimarySmtpAddress',
    'UserPrincipalName',
    'HiddenFromAddressListsEnabled',
    'ResourceCapacity',
    'Office',
    'Phone',
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

Write-Status -Message 'Starting Exchange Online resource mailbox inventory script.'
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
            throw 'ResourceMailboxIdentity is required. Use * to inventory all room/equipment mailboxes.'
        }

        $mailboxes = @()
        if ($resourceMailboxIdentity -eq '*') {
            $mailboxes = @(Invoke-WithRetry -OperationName 'Load all resource mailboxes' -ScriptBlock {
                Get-ExchangeOnlineMailbox -RecipientTypeDetails RoomMailbox,EquipmentMailbox -ResultSize Unlimited -Properties $mailboxProperties -ErrorAction Stop
            })
        }
        else {
            $mailbox = Invoke-WithRetry -OperationName "Lookup resource mailbox $resourceMailboxIdentity" -ScriptBlock {
                Get-ExchangeOnlineMailbox -Identity $resourceMailboxIdentity -Properties $mailboxProperties -ErrorAction SilentlyContinue
            }
            if ($mailbox -and (Get-StringPropertyValue -InputObject $mailbox -PropertyName 'RecipientTypeDetails') -in @('RoomMailbox', 'EquipmentMailbox')) {
                $mailboxes = @($mailbox)
            }
        }

        if ($mailboxes.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'GetExchangeResourceMailbox' -Status 'NotFound' -Message 'No matching resource mailboxes were found.' -Data ([ordered]@{
                        ResourceMailboxIdentity         = $resourceMailboxIdentity
                        ResourceType                    = ''
                        DisplayName                     = ''
                        Alias                           = ''
                        PrimarySmtpAddress              = ''
                        UserPrincipalName               = ''
                        HiddenFromAddressListsEnabled   = ''
                        ResourceCapacity                = ''
                        Office                          = ''
                        Phone                           = ''
                        AutomateProcessing              = ''
                        ForwardRequestsToDelegates      = ''
                        AllBookInPolicy                 = ''
                        AllRequestInPolicy              = ''
                        AllRequestOutOfPolicy           = ''
                        BookInPolicy                    = ''
                        RequestInPolicy                 = ''
                        RequestOutOfPolicy              = ''
                        BookingWindowInDays             = ''
                        MaximumDurationInMinutes        = ''
                        AllowRecurringMeetings          = ''
                        EnforceSchedulingHorizon        = ''
                        ScheduleOnlyDuringWorkHours     = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $identity = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'Identity'
            $calendar = $null
            try {
                $calendar = Invoke-WithRetry -OperationName "Load calendar processing $identity" -ScriptBlock {
                    Get-CalendarProcessing -Identity $identity -ErrorAction Stop
                }
            }
            catch {
                Write-Status -Message "Calendar processing lookup failed for '$identity': $($_.Exception.Message)" -Level WARN
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeResourceMailbox' -Status 'Completed' -Message 'Resource mailbox exported.' -Data ([ordered]@{
                        ResourceMailboxIdentity         = $identity
                        ResourceType                    = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'RecipientTypeDetails'
                        DisplayName                     = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'DisplayName'
                        Alias                           = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'Alias'
                        PrimarySmtpAddress              = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'PrimarySmtpAddress'
                        UserPrincipalName               = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'UserPrincipalName'
                        HiddenFromAddressListsEnabled   = [string](Get-ObjectPropertyValue -InputObject $mailbox -PropertyName 'HiddenFromAddressListsEnabled')
                        ResourceCapacity                = [string](Get-ObjectPropertyValue -InputObject $mailbox -PropertyName 'ResourceCapacity')
                        Office                          = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'Office'
                        Phone                           = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'Phone'
                        AutomateProcessing              = if ($calendar) { Get-StringPropertyValue -InputObject $calendar -PropertyName 'AutomateProcessing' } else { '' }
                        ForwardRequestsToDelegates      = if ($calendar) { [string](Get-ObjectPropertyValue -InputObject $calendar -PropertyName 'ForwardRequestsToDelegates') } else { '' }
                        AllBookInPolicy                 = if ($calendar) { [string](Get-ObjectPropertyValue -InputObject $calendar -PropertyName 'AllBookInPolicy') } else { '' }
                        AllRequestInPolicy              = if ($calendar) { [string](Get-ObjectPropertyValue -InputObject $calendar -PropertyName 'AllRequestInPolicy') } else { '' }
                        AllRequestOutOfPolicy           = if ($calendar) { [string](Get-ObjectPropertyValue -InputObject $calendar -PropertyName 'AllRequestOutOfPolicy') } else { '' }
                        BookInPolicy                    = if ($calendar) { Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $calendar -PropertyName 'BookInPolicy') } else { '' }
                        RequestInPolicy                 = if ($calendar) { Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $calendar -PropertyName 'RequestInPolicy') } else { '' }
                        RequestOutOfPolicy              = if ($calendar) { Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $calendar -PropertyName 'RequestOutOfPolicy') } else { '' }
                        BookingWindowInDays             = if ($calendar) { [string](Get-ObjectPropertyValue -InputObject $calendar -PropertyName 'BookingWindowInDays') } else { '' }
                        MaximumDurationInMinutes        = if ($calendar) { [string](Get-ObjectPropertyValue -InputObject $calendar -PropertyName 'MaximumDurationInMinutes') } else { '' }
                        AllowRecurringMeetings          = if ($calendar) { [string](Get-ObjectPropertyValue -InputObject $calendar -PropertyName 'AllowRecurringMeetings') } else { '' }
                        EnforceSchedulingHorizon        = if ($calendar) { [string](Get-ObjectPropertyValue -InputObject $calendar -PropertyName 'EnforceSchedulingHorizon') } else { '' }
                        ScheduleOnlyDuringWorkHours     = if ($calendar) { [string](Get-ObjectPropertyValue -InputObject $calendar -PropertyName 'ScheduleOnlyDuringWorkHours') } else { '' }
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($resourceMailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'GetExchangeResourceMailbox' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    ResourceMailboxIdentity         = $resourceMailboxIdentity
                    ResourceType                    = ''
                    DisplayName                     = ''
                    Alias                           = ''
                    PrimarySmtpAddress              = ''
                    UserPrincipalName               = ''
                    HiddenFromAddressListsEnabled   = ''
                    ResourceCapacity                = ''
                    Office                          = ''
                    Phone                           = ''
                    AutomateProcessing              = ''
                    ForwardRequestsToDelegates      = ''
                    AllBookInPolicy                 = ''
                    AllRequestInPolicy              = ''
                    AllRequestOutOfPolicy           = ''
                    BookInPolicy                    = ''
                    RequestInPolicy                 = ''
                    RequestOutOfPolicy              = ''
                    BookingWindowInDays             = ''
                    MaximumDurationInMinutes        = ''
                    AllowRecurringMeetings          = ''
                    EnforceSchedulingHorizon        = ''
                    ScheduleOnlyDuringWorkHours     = ''
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
Write-Status -Message 'Exchange Online resource mailbox inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







