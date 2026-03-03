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

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M3119-Set-ExchangeOnlineResourceMailboxBookingDelegates_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'ResourceMailboxIdentity',
    'DelegateUserPrincipalNames',
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

Write-Status -Message 'Starting Exchange Online resource mailbox booking delegate script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$setCalendarCommand = Get-Command -Name Set-CalendarProcessing -ErrorAction Stop
$supports = @{
    ResourceDelegates          = $setCalendarCommand.Parameters.ContainsKey('ResourceDelegates')
    AutomateProcessing         = $setCalendarCommand.Parameters.ContainsKey('AutomateProcessing')
    ForwardRequestsToDelegates = $setCalendarCommand.Parameters.ContainsKey('ForwardRequestsToDelegates')
    AllBookInPolicy            = $setCalendarCommand.Parameters.ContainsKey('AllBookInPolicy')
    AllRequestInPolicy         = $setCalendarCommand.Parameters.ContainsKey('AllRequestInPolicy')
    AllRequestOutOfPolicy      = $setCalendarCommand.Parameters.ContainsKey('AllRequestOutOfPolicy')
    BookInPolicy               = $setCalendarCommand.Parameters.ContainsKey('BookInPolicy')
    RequestInPolicy            = $setCalendarCommand.Parameters.ContainsKey('RequestInPolicy')
    RequestOutOfPolicy         = $setCalendarCommand.Parameters.ContainsKey('RequestOutOfPolicy')
    BookingWindowInDays        = $setCalendarCommand.Parameters.ContainsKey('BookingWindowInDays')
    MaximumDurationInMinutes   = $setCalendarCommand.Parameters.ContainsKey('MaximumDurationInMinutes')
    AllowRecurringMeetings     = $setCalendarCommand.Parameters.ContainsKey('AllowRecurringMeetings')
    EnforceSchedulingHorizon   = $setCalendarCommand.Parameters.ContainsKey('EnforceSchedulingHorizon')
    ScheduleOnlyDuringWorkHours= $setCalendarCommand.Parameters.ContainsKey('ScheduleOnlyDuringWorkHours')
}

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $resourceMailboxIdentity = ([string]$row.ResourceMailboxIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($resourceMailboxIdentity)) {
            throw 'ResourceMailboxIdentity is required.'
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup resource mailbox $resourceMailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $resourceMailboxIdentity -ErrorAction SilentlyContinue
        }
        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'SetResourceMailboxBookingDelegates' -Status 'NotFound' -Message 'Resource mailbox not found.'))
            $rowNumber++
            continue
        }

        $recipientTypeDetails = ([string]$mailbox.RecipientTypeDetails).Trim()
        if ($recipientTypeDetails -notin @('RoomMailbox', 'EquipmentMailbox')) {
            throw "Mailbox '$resourceMailboxIdentity' is '$recipientTypeDetails'. Only RoomMailbox and EquipmentMailbox are supported."
        }

        $setParams = @{
            Identity = $mailbox.Identity
        }
        $warnings = [System.Collections.Generic.List[string]]::new()

        $delegateRaw = ([string]$row.DelegateUserPrincipalNames).Trim()
        if (-not [string]::IsNullOrWhiteSpace($delegateRaw)) {
            if ($supports.ResourceDelegates) {
                if ($delegateRaw -eq '-') {
                    $setParams.ResourceDelegates = @()
                }
                else {
                    $delegateUpns = ConvertTo-Array -Value $delegateRaw
                    $delegateIdentities = [System.Collections.Generic.List[string]]::new()
                    foreach ($delegateUpn in $delegateUpns) {
                        $delegate = Invoke-WithRetry -OperationName "Resolve delegate $delegateUpn" -ScriptBlock {
                            Get-Recipient -Identity $delegateUpn -ErrorAction Stop
                        }
                        $delegateIdentities.Add(([string]$delegate.Identity).Trim())
                    }
                    $setParams.ResourceDelegates = $delegateIdentities.ToArray()
                }
            }
            else {
                $warnings.Add('DelegateUserPrincipalNames ignored (unsupported parameter).')
            }
        }

        $automateProcessing = ([string]$row.AutomateProcessing).Trim()
        if (-not [string]::IsNullOrWhiteSpace($automateProcessing) -and $supports.AutomateProcessing) {
            if ($automateProcessing -notin @('AutoAccept', 'AutoUpdate', 'None')) {
                throw "AutomateProcessing '$automateProcessing' is invalid. Use AutoAccept, AutoUpdate, or None."
            }
            $setParams.AutomateProcessing = $automateProcessing
        }

        $forwardRequestsToDelegatesRaw = ([string]$row.ForwardRequestsToDelegates).Trim()
        if (-not [string]::IsNullOrWhiteSpace($forwardRequestsToDelegatesRaw) -and $supports.ForwardRequestsToDelegates) {
            $setParams.ForwardRequestsToDelegates = ConvertTo-Bool -Value $forwardRequestsToDelegatesRaw
        }

        $allBookInPolicyRaw = ([string]$row.AllBookInPolicy).Trim()
        if (-not [string]::IsNullOrWhiteSpace($allBookInPolicyRaw) -and $supports.AllBookInPolicy) {
            $setParams.AllBookInPolicy = ConvertTo-Bool -Value $allBookInPolicyRaw
        }

        $allRequestInPolicyRaw = ([string]$row.AllRequestInPolicy).Trim()
        if (-not [string]::IsNullOrWhiteSpace($allRequestInPolicyRaw) -and $supports.AllRequestInPolicy) {
            $setParams.AllRequestInPolicy = ConvertTo-Bool -Value $allRequestInPolicyRaw
        }

        $allRequestOutOfPolicyRaw = ([string]$row.AllRequestOutOfPolicy).Trim()
        if (-not [string]::IsNullOrWhiteSpace($allRequestOutOfPolicyRaw) -and $supports.AllRequestOutOfPolicy) {
            $setParams.AllRequestOutOfPolicy = ConvertTo-Bool -Value $allRequestOutOfPolicyRaw
        }

        $bookInPolicy = ConvertTo-Array -Value ([string]$row.BookInPolicy)
        if ($bookInPolicy.Count -gt 0 -and $supports.BookInPolicy) {
            $setParams.BookInPolicy = $bookInPolicy
        }

        $requestInPolicy = ConvertTo-Array -Value ([string]$row.RequestInPolicy)
        if ($requestInPolicy.Count -gt 0 -and $supports.RequestInPolicy) {
            $setParams.RequestInPolicy = $requestInPolicy
        }

        $requestOutOfPolicy = ConvertTo-Array -Value ([string]$row.RequestOutOfPolicy)
        if ($requestOutOfPolicy.Count -gt 0 -and $supports.RequestOutOfPolicy) {
            $setParams.RequestOutOfPolicy = $requestOutOfPolicy
        }

        $bookingWindowInDaysRaw = ([string]$row.BookingWindowInDays).Trim()
        if (-not [string]::IsNullOrWhiteSpace($bookingWindowInDaysRaw) -and $supports.BookingWindowInDays) {
            $bookingWindowInDays = 0
            if (-not [int]::TryParse($bookingWindowInDaysRaw, [ref]$bookingWindowInDays)) {
                throw "BookingWindowInDays '$bookingWindowInDaysRaw' is not a valid integer."
            }
            $setParams.BookingWindowInDays = $bookingWindowInDays
        }

        $maximumDurationRaw = ([string]$row.MaximumDurationInMinutes).Trim()
        if (-not [string]::IsNullOrWhiteSpace($maximumDurationRaw) -and $supports.MaximumDurationInMinutes) {
            $maximumDuration = 0
            if (-not [int]::TryParse($maximumDurationRaw, [ref]$maximumDuration)) {
                throw "MaximumDurationInMinutes '$maximumDurationRaw' is not a valid integer."
            }
            $setParams.MaximumDurationInMinutes = $maximumDuration
        }

        $allowRecurringMeetingsRaw = ([string]$row.AllowRecurringMeetings).Trim()
        if (-not [string]::IsNullOrWhiteSpace($allowRecurringMeetingsRaw) -and $supports.AllowRecurringMeetings) {
            $setParams.AllowRecurringMeetings = ConvertTo-Bool -Value $allowRecurringMeetingsRaw
        }

        $enforceSchedulingHorizonRaw = ([string]$row.EnforceSchedulingHorizon).Trim()
        if (-not [string]::IsNullOrWhiteSpace($enforceSchedulingHorizonRaw) -and $supports.EnforceSchedulingHorizon) {
            $setParams.EnforceSchedulingHorizon = ConvertTo-Bool -Value $enforceSchedulingHorizonRaw
        }

        $scheduleOnlyDuringWorkHoursRaw = ([string]$row.ScheduleOnlyDuringWorkHours).Trim()
        if (-not [string]::IsNullOrWhiteSpace($scheduleOnlyDuringWorkHoursRaw) -and $supports.ScheduleOnlyDuringWorkHours) {
            $setParams.ScheduleOnlyDuringWorkHours = ConvertTo-Bool -Value $scheduleOnlyDuringWorkHoursRaw
        }

        if ($setParams.Count -eq 1) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'SetResourceMailboxBookingDelegates' -Status 'Skipped' -Message 'No updates specified.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($resourceMailboxIdentity, 'Set resource mailbox booking delegates and policy')) {
            Invoke-WithRetry -OperationName "Set calendar processing $resourceMailboxIdentity" -ScriptBlock {
                Set-CalendarProcessing @setParams -ErrorAction Stop
            }
            $message = 'Resource mailbox booking settings updated successfully.'
            if ($warnings.Count -gt 0) {
                $message = "$message $($warnings -join ' ')"
            }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'SetResourceMailboxBookingDelegates' -Status 'Completed' -Message $message))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'SetResourceMailboxBookingDelegates' -Status 'WhatIf' -Message 'Booking delegate update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($resourceMailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'SetResourceMailboxBookingDelegates' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online resource mailbox booking delegate script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}






