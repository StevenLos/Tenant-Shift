<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-220000

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M0218-Update-ExchangeOnPremResourceMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'ResourceMailboxIdentity',
    'DisplayName',
    'PrimarySmtpAddress',
    'Capacity',
    'HiddenFromAddressListsEnabled',
    'AutomateProcessing',
    'BookingWindowInDays',
    'MaximumDurationInMinutes',
    'AllowConflicts',
    'AllBookInPolicy',
    'AllRequestInPolicy',
    'AllRequestOutOfPolicy',
    'EnforceSchedulingHorizon'
)

Write-Status -Message 'Starting Exchange on-prem resource mailbox update script.'
Ensure-ExchangeOnPremConnection

$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop
$setCalendarCommand = Get-Command -Name Set-CalendarProcessing -ErrorAction SilentlyContinue

$supports = @{
    HiddenFromAddressListsEnabled = $setMailboxCommand.Parameters.ContainsKey('HiddenFromAddressListsEnabled')
    ResourceCapacity              = $setMailboxCommand.Parameters.ContainsKey('ResourceCapacity')
}

$calendarSupports = @{}
if ($setCalendarCommand) {
    foreach ($paramName in @('AutomateProcessing', 'BookingWindowInDays', 'MaximumDurationInMinutes', 'AllowConflicts', 'AllBookInPolicy', 'AllRequestInPolicy', 'AllRequestOutOfPolicy', 'EnforceSchedulingHorizon')) {
        $calendarSupports[$paramName] = $setCalendarCommand.Parameters.ContainsKey($paramName)
    }
}

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $resourceMailboxIdentity = Get-TrimmedValue -Value $row.ResourceMailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($resourceMailboxIdentity)) {
            throw 'ResourceMailboxIdentity is required.'
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup resource mailbox $resourceMailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $resourceMailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'UpdateResourceMailbox' -Status 'NotFound' -Message 'Resource mailbox not found.'))
            $rowNumber++
            continue
        }

        $recipientTypeDetails = Get-TrimmedValue -Value $mailbox.RecipientTypeDetails
        if ($recipientTypeDetails -notin @('RoomMailbox', 'EquipmentMailbox')) {
            throw "Recipient '$resourceMailboxIdentity' is '$recipientTypeDetails'. Expected RoomMailbox or EquipmentMailbox."
        }

        $setMailboxParams = @{ Identity = $mailbox.Identity }
        $warnings = [System.Collections.Generic.List[string]]::new()

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $setMailboxParams.DisplayName = $displayName
        }

        $primarySmtpAddress = Get-TrimmedValue -Value $row.PrimarySmtpAddress
        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $setMailboxParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $capacityRaw = Get-TrimmedValue -Value $row.Capacity
        if (-not [string]::IsNullOrWhiteSpace($capacityRaw)) {
            $parsedCapacity = 0
            if (-not [int]::TryParse($capacityRaw, [ref]$parsedCapacity) -or $parsedCapacity -lt 0) {
                throw "Capacity '$capacityRaw' is invalid. Use a non-negative integer."
            }

            if ($supports.ResourceCapacity) {
                $setMailboxParams.ResourceCapacity = $parsedCapacity
            }
            else {
                $warnings.Add('ResourceCapacity ignored (unsupported parameter).')
            }
        }

        $hiddenRaw = Get-TrimmedValue -Value $row.HiddenFromAddressListsEnabled
        if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
            if ($supports.HiddenFromAddressListsEnabled) {
                $setMailboxParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
            }
            else {
                $warnings.Add('HiddenFromAddressListsEnabled ignored (unsupported parameter).')
            }
        }

        $setCalendarParams = @{ Identity = $mailbox.Identity }
        if ($setCalendarCommand) {
            foreach ($calendarName in @('AutomateProcessing', 'BookingWindowInDays', 'MaximumDurationInMinutes', 'AllowConflicts', 'AllBookInPolicy', 'AllRequestInPolicy', 'AllRequestOutOfPolicy', 'EnforceSchedulingHorizon')) {
                $rawValue = Get-TrimmedValue -Value $row.$calendarName
                if ([string]::IsNullOrWhiteSpace($rawValue)) {
                    continue
                }

                if (-not $calendarSupports[$calendarName]) {
                    $warnings.Add("$calendarName ignored (unsupported parameter).")
                    continue
                }

                if ($calendarName -eq 'AutomateProcessing') {
                    $setCalendarParams[$calendarName] = $rawValue
                    continue
                }

                if ($calendarName -eq 'BookingWindowInDays' -or $calendarName -eq 'MaximumDurationInMinutes') {
                    $parsedNumber = 0
                    if (-not [int]::TryParse($rawValue, [ref]$parsedNumber) -or $parsedNumber -lt 0) {
                        throw "$calendarName '$rawValue' is invalid. Use a non-negative integer."
                    }

                    $setCalendarParams[$calendarName] = $parsedNumber
                    continue
                }

                $setCalendarParams[$calendarName] = ConvertTo-Bool -Value $rawValue
            }
        }
        else {
            foreach ($calendarName in @('AutomateProcessing', 'BookingWindowInDays', 'MaximumDurationInMinutes', 'AllowConflicts', 'AllBookInPolicy', 'AllRequestInPolicy', 'AllRequestOutOfPolicy', 'EnforceSchedulingHorizon')) {
                if (-not [string]::IsNullOrWhiteSpace((Get-TrimmedValue -Value $row.$calendarName))) {
                    $warnings.Add("$calendarName ignored (Set-CalendarProcessing unavailable).")
                }
            }
        }

        if ($setMailboxParams.Count -eq 1 -and $setCalendarParams.Count -eq 1) {
            $message = 'No updates specified.'
            if ($warnings.Count -gt 0) {
                $message = "$message $($warnings -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'UpdateResourceMailbox' -Status 'Skipped' -Message $message))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($resourceMailboxIdentity, 'Update Exchange on-prem resource mailbox')) {
            if ($setMailboxParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Update resource mailbox core properties $resourceMailboxIdentity" -ScriptBlock {
                    Set-Mailbox @setMailboxParams -ErrorAction Stop
                }
            }

            if ($setCalendarParams.Count -gt 1 -and $setCalendarCommand) {
                Invoke-WithRetry -OperationName "Update resource mailbox calendar processing $resourceMailboxIdentity" -ScriptBlock {
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
Write-Status -Message 'Exchange on-prem resource mailbox update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
