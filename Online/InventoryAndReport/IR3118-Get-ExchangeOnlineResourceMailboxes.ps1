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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3118-Get-ExchangeOnlineResourceMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'ResourceMailboxIdentity'
)

Write-Status -Message 'Starting Exchange Online resource mailbox inventory script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
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
                Get-Mailbox -RecipientTypeDetails RoomMailbox,EquipmentMailbox -ResultSize Unlimited -ErrorAction Stop
            })
        }
        else {
            $mailbox = Invoke-WithRetry -OperationName "Lookup resource mailbox $resourceMailboxIdentity" -ScriptBlock {
                Get-Mailbox -Identity $resourceMailboxIdentity -ErrorAction SilentlyContinue
            }
            if ($mailbox -and ([string]$mailbox.RecipientTypeDetails).Trim() -in @('RoomMailbox', 'EquipmentMailbox')) {
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
            $identity = ([string]$mailbox.Identity).Trim()
            $calendar = $null
            try {
                $calendar = Invoke-WithRetry -OperationName "Load calendar processing $identity" -ScriptBlock {
                    Get-CalendarProcessing -Identity $mailbox.Identity -ErrorAction Stop
                }
            }
            catch {
                Write-Status -Message "Calendar processing lookup failed for '$identity': $($_.Exception.Message)" -Level WARN
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeResourceMailbox' -Status 'Completed' -Message 'Resource mailbox exported.' -Data ([ordered]@{
                        ResourceMailboxIdentity         = $identity
                        ResourceType                    = ([string]$mailbox.RecipientTypeDetails).Trim()
                        DisplayName                     = ([string]$mailbox.DisplayName).Trim()
                        Alias                           = ([string]$mailbox.Alias).Trim()
                        PrimarySmtpAddress              = ([string]$mailbox.PrimarySmtpAddress).Trim()
                        UserPrincipalName               = ([string]$mailbox.UserPrincipalName).Trim()
                        HiddenFromAddressListsEnabled   = [string]$mailbox.HiddenFromAddressListsEnabled
                        ResourceCapacity                = [string]$mailbox.ResourceCapacity
                        Office                          = ([string]$mailbox.Office).Trim()
                        Phone                           = ([string]$mailbox.Phone).Trim()
                        AutomateProcessing              = if ($calendar) { ([string]$calendar.AutomateProcessing).Trim() } else { '' }
                        ForwardRequestsToDelegates      = if ($calendar) { [string]$calendar.ForwardRequestsToDelegates } else { '' }
                        AllBookInPolicy                 = if ($calendar) { [string]$calendar.AllBookInPolicy } else { '' }
                        AllRequestInPolicy              = if ($calendar) { [string]$calendar.AllRequestInPolicy } else { '' }
                        AllRequestOutOfPolicy           = if ($calendar) { [string]$calendar.AllRequestOutOfPolicy } else { '' }
                        BookInPolicy                    = if ($calendar) { Convert-MultiValueToString -Value $calendar.BookInPolicy } else { '' }
                        RequestInPolicy                 = if ($calendar) { Convert-MultiValueToString -Value $calendar.RequestInPolicy } else { '' }
                        RequestOutOfPolicy              = if ($calendar) { Convert-MultiValueToString -Value $calendar.RequestOutOfPolicy } else { '' }
                        BookingWindowInDays             = if ($calendar) { [string]$calendar.BookingWindowInDays } else { '' }
                        MaximumDurationInMinutes        = if ($calendar) { [string]$calendar.MaximumDurationInMinutes } else { '' }
                        AllowRecurringMeetings          = if ($calendar) { [string]$calendar.AllowRecurringMeetings } else { '' }
                        EnforceSchedulingHorizon        = if ($calendar) { [string]$calendar.EnforceSchedulingHorizon } else { '' }
                        ScheduleOnlyDuringWorkHours     = if ($calendar) { [string]$calendar.ScheduleOnlyDuringWorkHours } else { '' }
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

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online resource mailbox inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





