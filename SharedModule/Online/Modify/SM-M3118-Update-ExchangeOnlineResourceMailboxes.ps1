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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3118-Update-ExchangeOnlineResourceMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'DisplayName',
    'PrimarySmtpAddress',
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

Write-Status -Message 'Starting Exchange Online resource mailbox update script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop
$setCalendarCommand = Get-Command -Name Set-CalendarProcessing -ErrorAction Stop

$setMailboxSupportsResourceCapacity = $setMailboxCommand.Parameters.ContainsKey('ResourceCapacity')
$setMailboxSupportsOffice = $setMailboxCommand.Parameters.ContainsKey('Office')
$setMailboxSupportsPhone = $setMailboxCommand.Parameters.ContainsKey('Phone')

$setCalendarSupports = @{
    AutomateProcessing          = $setCalendarCommand.Parameters.ContainsKey('AutomateProcessing')
    ForwardRequestsToDelegates  = $setCalendarCommand.Parameters.ContainsKey('ForwardRequestsToDelegates')
    AllBookInPolicy             = $setCalendarCommand.Parameters.ContainsKey('AllBookInPolicy')
    AllRequestInPolicy          = $setCalendarCommand.Parameters.ContainsKey('AllRequestInPolicy')
    AllRequestOutOfPolicy       = $setCalendarCommand.Parameters.ContainsKey('AllRequestOutOfPolicy')
    BookInPolicy                = $setCalendarCommand.Parameters.ContainsKey('BookInPolicy')
    RequestInPolicy             = $setCalendarCommand.Parameters.ContainsKey('RequestInPolicy')
    RequestOutOfPolicy          = $setCalendarCommand.Parameters.ContainsKey('RequestOutOfPolicy')
    BookingWindowInDays         = $setCalendarCommand.Parameters.ContainsKey('BookingWindowInDays')
    MaximumDurationInMinutes    = $setCalendarCommand.Parameters.ContainsKey('MaximumDurationInMinutes')
    AllowRecurringMeetings      = $setCalendarCommand.Parameters.ContainsKey('AllowRecurringMeetings')
    EnforceSchedulingHorizon    = $setCalendarCommand.Parameters.ContainsKey('EnforceSchedulingHorizon')
    ScheduleOnlyDuringWorkHours = $setCalendarCommand.Parameters.ContainsKey('ScheduleOnlyDuringWorkHours')
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
            Get-ExchangeOnlineMailbox -Identity $resourceMailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'UpdateResourceMailbox' -Status 'NotFound' -Message 'Resource mailbox not found.'))
            $rowNumber++
            continue
        }

        $recipientTypeDetails = ([string]$mailbox.RecipientTypeDetails).Trim()
        if ($recipientTypeDetails -notin @('RoomMailbox', 'EquipmentMailbox')) {
            throw "Recipient '$resourceMailboxIdentity' is '$recipientTypeDetails'. Only RoomMailbox and EquipmentMailbox are supported."
        }

        $setMailboxParams = @{
            Identity = $mailbox.Identity
        }
        $warnings = [System.Collections.Generic.List[string]]::new()

        $displayName = ([string]$row.DisplayName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $setMailboxParams.DisplayName = $displayName
        }

        $primarySmtpAddress = ([string]$row.PrimarySmtpAddress).Trim()
        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $setMailboxParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $hiddenRaw = ([string]$row.HiddenFromAddressListsEnabled).Trim()
        if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
            $setMailboxParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
        }

        $resourceCapacityRaw = ([string]$row.ResourceCapacity).Trim()
        if (-not [string]::IsNullOrWhiteSpace($resourceCapacityRaw)) {
            if ($setMailboxSupportsResourceCapacity) {
                $resourceCapacity = 0
                if (-not [int]::TryParse($resourceCapacityRaw, [ref]$resourceCapacity)) {
                    throw "ResourceCapacity '$resourceCapacityRaw' is not a valid integer."
                }
                $setMailboxParams.ResourceCapacity = $resourceCapacity
            }
            else {
                $warnings.Add('ResourceCapacity ignored (unsupported parameter).')
            }
        }

        $office = ([string]$row.Office).Trim()
        if (-not [string]::IsNullOrWhiteSpace($office)) {
            if ($setMailboxSupportsOffice) {
                $setMailboxParams.Office = $office
            }
            else {
                $warnings.Add('Office ignored (unsupported parameter).')
            }
        }

        $phone = ([string]$row.Phone).Trim()
        if (-not [string]::IsNullOrWhiteSpace($phone)) {
            if ($setMailboxSupportsPhone) {
                $setMailboxParams.Phone = $phone
            }
            else {
                $warnings.Add('Phone ignored (unsupported parameter).')
            }
        }

        $setCalendarParams = @{
            Identity = $mailbox.Identity
        }

        $automateProcessing = ([string]$row.AutomateProcessing).Trim()
        if (-not [string]::IsNullOrWhiteSpace($automateProcessing) -and $setCalendarSupports.AutomateProcessing) {
            $setCalendarParams.AutomateProcessing = $automateProcessing
        }

        $forwardRequestsToDelegatesRaw = ([string]$row.ForwardRequestsToDelegates).Trim()
        if (-not [string]::IsNullOrWhiteSpace($forwardRequestsToDelegatesRaw) -and $setCalendarSupports.ForwardRequestsToDelegates) {
            $setCalendarParams.ForwardRequestsToDelegates = ConvertTo-Bool -Value $forwardRequestsToDelegatesRaw
        }

        $allBookInPolicyRaw = ([string]$row.AllBookInPolicy).Trim()
        if (-not [string]::IsNullOrWhiteSpace($allBookInPolicyRaw) -and $setCalendarSupports.AllBookInPolicy) {
            $setCalendarParams.AllBookInPolicy = ConvertTo-Bool -Value $allBookInPolicyRaw
        }

        $allRequestInPolicyRaw = ([string]$row.AllRequestInPolicy).Trim()
        if (-not [string]::IsNullOrWhiteSpace($allRequestInPolicyRaw) -and $setCalendarSupports.AllRequestInPolicy) {
            $setCalendarParams.AllRequestInPolicy = ConvertTo-Bool -Value $allRequestInPolicyRaw
        }

        $allRequestOutOfPolicyRaw = ([string]$row.AllRequestOutOfPolicy).Trim()
        if (-not [string]::IsNullOrWhiteSpace($allRequestOutOfPolicyRaw) -and $setCalendarSupports.AllRequestOutOfPolicy) {
            $setCalendarParams.AllRequestOutOfPolicy = ConvertTo-Bool -Value $allRequestOutOfPolicyRaw
        }

        $bookInPolicy = ConvertTo-Array -Value ([string]$row.BookInPolicy)
        if ($bookInPolicy.Count -gt 0 -and $setCalendarSupports.BookInPolicy) {
            $setCalendarParams.BookInPolicy = $bookInPolicy
        }

        $requestInPolicy = ConvertTo-Array -Value ([string]$row.RequestInPolicy)
        if ($requestInPolicy.Count -gt 0 -and $setCalendarSupports.RequestInPolicy) {
            $setCalendarParams.RequestInPolicy = $requestInPolicy
        }

        $requestOutOfPolicy = ConvertTo-Array -Value ([string]$row.RequestOutOfPolicy)
        if ($requestOutOfPolicy.Count -gt 0 -and $setCalendarSupports.RequestOutOfPolicy) {
            $setCalendarParams.RequestOutOfPolicy = $requestOutOfPolicy
        }

        $bookingWindowInDaysRaw = ([string]$row.BookingWindowInDays).Trim()
        if (-not [string]::IsNullOrWhiteSpace($bookingWindowInDaysRaw) -and $setCalendarSupports.BookingWindowInDays) {
            $bookingWindowInDays = 0
            if (-not [int]::TryParse($bookingWindowInDaysRaw, [ref]$bookingWindowInDays)) {
                throw "BookingWindowInDays '$bookingWindowInDaysRaw' is not a valid integer."
            }
            $setCalendarParams.BookingWindowInDays = $bookingWindowInDays
        }

        $maximumDurationRaw = ([string]$row.MaximumDurationInMinutes).Trim()
        if (-not [string]::IsNullOrWhiteSpace($maximumDurationRaw) -and $setCalendarSupports.MaximumDurationInMinutes) {
            $maximumDuration = 0
            if (-not [int]::TryParse($maximumDurationRaw, [ref]$maximumDuration)) {
                throw "MaximumDurationInMinutes '$maximumDurationRaw' is not a valid integer."
            }
            $setCalendarParams.MaximumDurationInMinutes = $maximumDuration
        }

        $allowRecurringMeetingsRaw = ([string]$row.AllowRecurringMeetings).Trim()
        if (-not [string]::IsNullOrWhiteSpace($allowRecurringMeetingsRaw) -and $setCalendarSupports.AllowRecurringMeetings) {
            $setCalendarParams.AllowRecurringMeetings = ConvertTo-Bool -Value $allowRecurringMeetingsRaw
        }

        $enforceSchedulingHorizonRaw = ([string]$row.EnforceSchedulingHorizon).Trim()
        if (-not [string]::IsNullOrWhiteSpace($enforceSchedulingHorizonRaw) -and $setCalendarSupports.EnforceSchedulingHorizon) {
            $setCalendarParams.EnforceSchedulingHorizon = ConvertTo-Bool -Value $enforceSchedulingHorizonRaw
        }

        $scheduleOnlyDuringWorkHoursRaw = ([string]$row.ScheduleOnlyDuringWorkHours).Trim()
        if (-not [string]::IsNullOrWhiteSpace($scheduleOnlyDuringWorkHoursRaw) -and $setCalendarSupports.ScheduleOnlyDuringWorkHours) {
            $setCalendarParams.ScheduleOnlyDuringWorkHours = ConvertTo-Bool -Value $scheduleOnlyDuringWorkHoursRaw
        }

        if ($setMailboxParams.Count -eq 1 -and $setCalendarParams.Count -eq 1) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'UpdateResourceMailbox' -Status 'Skipped' -Message 'No updates specified.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($resourceMailboxIdentity, 'Update Exchange Online resource mailbox settings')) {
            if ($setMailboxParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Update resource mailbox options $resourceMailboxIdentity" -ScriptBlock {
                    Set-Mailbox @setMailboxParams -ErrorAction Stop
                }
            }

            if ($setCalendarParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Update calendar processing $resourceMailboxIdentity" -ScriptBlock {
                    Set-CalendarProcessing @setCalendarParams -ErrorAction Stop
                }
            }

            $message = 'Resource mailbox updated successfully.'
            if ($warnings.Count -gt 0) {
                $message = "$message $($warnings -join ' ')"
            }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'UpdateResourceMailbox' -Status 'Updated' -Message $message))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'UpdateResourceMailbox' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($resourceMailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'UpdateResourceMailbox' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online resource mailbox update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





