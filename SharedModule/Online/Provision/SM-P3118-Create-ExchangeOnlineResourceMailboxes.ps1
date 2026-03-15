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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P3118-Create-ExchangeOnlineResourceMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'ResourceType',
    'Name',
    'Alias',
    'DisplayName',
    'UserPrincipalName',
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

Write-Status -Message 'Starting Exchange Online resource mailbox creation script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$newMailboxCommand = Get-Command -Name New-Mailbox -ErrorAction Stop
$newMailboxSupportsUserPrincipalName = $newMailboxCommand.Parameters.ContainsKey('UserPrincipalName')
$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop
$setCalendarCommand = Get-Command -Name Set-CalendarProcessing -ErrorAction Stop

$setMailboxSupportsResourceCapacity = $setMailboxCommand.Parameters.ContainsKey('ResourceCapacity')
$setMailboxSupportsOffice = $setMailboxCommand.Parameters.ContainsKey('Office')
$setMailboxSupportsPhone = $setMailboxCommand.Parameters.ContainsKey('Phone')

$setCalendarSupports = @{
    AutomateProcessing        = $setCalendarCommand.Parameters.ContainsKey('AutomateProcessing')
    ForwardRequestsToDelegates= $setCalendarCommand.Parameters.ContainsKey('ForwardRequestsToDelegates')
    AllBookInPolicy           = $setCalendarCommand.Parameters.ContainsKey('AllBookInPolicy')
    AllRequestInPolicy        = $setCalendarCommand.Parameters.ContainsKey('AllRequestInPolicy')
    AllRequestOutOfPolicy     = $setCalendarCommand.Parameters.ContainsKey('AllRequestOutOfPolicy')
    BookInPolicy              = $setCalendarCommand.Parameters.ContainsKey('BookInPolicy')
    RequestInPolicy           = $setCalendarCommand.Parameters.ContainsKey('RequestInPolicy')
    RequestOutOfPolicy        = $setCalendarCommand.Parameters.ContainsKey('RequestOutOfPolicy')
    BookingWindowInDays       = $setCalendarCommand.Parameters.ContainsKey('BookingWindowInDays')
    MaximumDurationInMinutes  = $setCalendarCommand.Parameters.ContainsKey('MaximumDurationInMinutes')
    AllowRecurringMeetings    = $setCalendarCommand.Parameters.ContainsKey('AllowRecurringMeetings')
    EnforceSchedulingHorizon  = $setCalendarCommand.Parameters.ContainsKey('EnforceSchedulingHorizon')
    ScheduleOnlyDuringWorkHours = $setCalendarCommand.Parameters.ContainsKey('ScheduleOnlyDuringWorkHours')
}

if (-not $newMailboxSupportsUserPrincipalName) {
    Write-Status -Message "New-Mailbox in this session does not support -UserPrincipalName. The 'UserPrincipalName' CSV value will be ignored." -Level WARN
}

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $name = ([string]$row.Name).Trim()
    $resourceTypeRaw = ([string]$row.ResourceType).Trim()
    $resourceType = $resourceTypeRaw.ToLowerInvariant()

    try {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Name is required.'
        }

        if ($resourceType -notin @('room', 'equipment')) {
            throw "ResourceType '$resourceTypeRaw' is invalid. Use Room or Equipment."
        }

        $alias = ([string]$row.Alias).Trim()
        $displayName = ([string]$row.DisplayName).Trim()
        $userPrincipalName = ([string]$row.UserPrincipalName).Trim()
        $primarySmtpAddress = ([string]$row.PrimarySmtpAddress).Trim()
        $office = ([string]$row.Office).Trim()
        $phone = ([string]$row.Phone).Trim()

        $lookupIdentity = if (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) {
            $userPrincipalName
        }
        elseif (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $primarySmtpAddress
        }
        elseif (-not [string]::IsNullOrWhiteSpace($alias)) {
            $alias
        }
        else {
            $name
        }

        $existingMailbox = Invoke-WithRetry -OperationName "Lookup resource mailbox $lookupIdentity" -ScriptBlock {
            Get-ExchangeOnlineMailbox -Identity $lookupIdentity -ErrorAction SilentlyContinue
        }

        if ($existingMailbox) {
            $recipientTypeDetails = ([string]$existingMailbox.RecipientTypeDetails).Trim()
            if ($resourceType -eq 'room' -and $recipientTypeDetails -eq 'RoomMailbox') {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateResourceMailbox' -Status 'Skipped' -Message 'Room mailbox already exists.'))
                $rowNumber++
                continue
            }

            if ($resourceType -eq 'equipment' -and $recipientTypeDetails -eq 'EquipmentMailbox') {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateResourceMailbox' -Status 'Skipped' -Message 'Equipment mailbox already exists.'))
                $rowNumber++
                continue
            }

            throw "Mailbox '$lookupIdentity' already exists with recipient type '$recipientTypeDetails', which does not match requested resource type '$resourceTypeRaw'."
        }

        $createParams = @{
            Name = $name
        }

        if ($resourceType -eq 'room') {
            $createParams.Room = $true
        }
        else {
            $createParams.Equipment = $true
        }

        if (-not [string]::IsNullOrWhiteSpace($alias)) {
            $createParams.Alias = $alias
        }

        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $createParams.DisplayName = $displayName
        }

        $upnIgnored = $false
        if (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) {
            if ($newMailboxSupportsUserPrincipalName) {
                $createParams.UserPrincipalName = $userPrincipalName
            }
            else {
                $upnIgnored = $true
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $createParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $resourceCapacity = 0
        $resourceCapacityRaw = ([string]$row.ResourceCapacity).Trim()
        $setResourceCapacity = $false
        if (-not [string]::IsNullOrWhiteSpace($resourceCapacityRaw)) {
            if (-not [int]::TryParse($resourceCapacityRaw, [ref]$resourceCapacity)) {
                throw "ResourceCapacity '$resourceCapacityRaw' is not a valid integer."
            }
            if ($resourceCapacity -lt 0) {
                throw 'ResourceCapacity must be zero or greater.'
            }
            $setResourceCapacity = $true
        }

        if ($PSCmdlet.ShouldProcess($lookupIdentity, 'Create Exchange Online resource mailbox')) {
            $createdMailbox = Invoke-WithRetry -OperationName "Create resource mailbox $lookupIdentity" -ScriptBlock {
                New-Mailbox @createParams -ErrorAction Stop
            }

            $setMailboxParams = @{
                Identity = $createdMailbox.Identity
            }
            $messages = [System.Collections.Generic.List[string]]::new()

            $hiddenRaw = ([string]$row.HiddenFromAddressListsEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
                $setMailboxParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
            }

            if ($setResourceCapacity) {
                if ($setMailboxSupportsResourceCapacity) {
                    $setMailboxParams.ResourceCapacity = $resourceCapacity
                }
                else {
                    $messages.Add('ResourceCapacity ignored (unsupported parameter).')
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($office)) {
                if ($setMailboxSupportsOffice) {
                    $setMailboxParams.Office = $office
                }
                else {
                    $messages.Add('Office ignored (unsupported parameter).')
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($phone)) {
                if ($setMailboxSupportsPhone) {
                    $setMailboxParams.Phone = $phone
                }
                else {
                    $messages.Add('Phone ignored (unsupported parameter).')
                }
            }

            if ($setMailboxParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Set resource mailbox options $lookupIdentity" -ScriptBlock {
                    Set-Mailbox @setMailboxParams -ErrorAction Stop
                }
            }

            $setCalendarParams = @{
                Identity = $createdMailbox.Identity
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

            if ($setCalendarParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Set calendar processing for $lookupIdentity" -ScriptBlock {
                    Set-CalendarProcessing @setCalendarParams -ErrorAction Stop
                }
            }

            $successMessage = 'Resource mailbox created successfully.'
            if ($upnIgnored) {
                $successMessage = "$successMessage UserPrincipalName was provided but ignored because this New-Mailbox session does not support -UserPrincipalName."
            }
            if ($messages.Count -gt 0) {
                $successMessage = "$successMessage $($messages -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateResourceMailbox' -Status 'Created' -Message $successMessage))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateResourceMailbox' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($name) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $name -Action 'CreateResourceMailbox' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online resource mailbox creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





