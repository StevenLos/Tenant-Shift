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

.SYNOPSIS
    Modifies ExchangeOnPremResourceMailboxBookingDelegates in Active Directory.

.DESCRIPTION
    Updates ExchangeOnPremResourceMailboxBookingDelegates in Active Directory based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M0219-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1 -InputCsvPath .\0219.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M0219-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1 -InputCsvPath .\0219.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Exchange Management Shell cmdlets (session)
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    Column                   Type      Required  Description
    -----------------------  ----      --------  -----------
    ResourceMailboxIdentity  String    Yes       <fill in description>
    DelegateIdentity         String    Yes       <fill in description>
    DelegateAction           String    Yes       <fill in description>
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M0219-Set-ExchangeOnPremResourceMailboxBookingDelegates_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-RecipientKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Recipient
    )

    $primary = Get-TrimmedValue -Value $Recipient.PrimarySmtpAddress
    if (-not [string]::IsNullOrWhiteSpace($primary)) {
        return $primary.ToLowerInvariant()
    }

    $identity = Get-TrimmedValue -Value $Recipient.Identity
    if (-not [string]::IsNullOrWhiteSpace($identity)) {
        return $identity.ToLowerInvariant()
    }

    return ''
}

$requiredHeaders = @(
    'ResourceMailboxIdentity',
    'DelegateIdentity',
    'DelegateAction'
)

Write-Status -Message 'Starting Exchange on-prem resource mailbox booking delegate script.'
Ensure-ExchangeOnPremConnection

$setCalendarCommand = Get-Command -Name Set-CalendarProcessing -ErrorAction Stop
if (-not $setCalendarCommand.Parameters.ContainsKey('ResourceDelegates')) {
    throw 'Set-CalendarProcessing -ResourceDelegates is not available in this session.'
}

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $resourceMailboxIdentity = Get-TrimmedValue -Value $row.ResourceMailboxIdentity
    $delegateIdentity = Get-TrimmedValue -Value $row.DelegateIdentity
    $delegateActionRaw = Get-TrimmedValue -Value $row.DelegateAction

    try {
        if ([string]::IsNullOrWhiteSpace($resourceMailboxIdentity) -or [string]::IsNullOrWhiteSpace($delegateIdentity)) {
            throw 'ResourceMailboxIdentity and DelegateIdentity are required.'
        }

        $delegateAction = if ([string]::IsNullOrWhiteSpace($delegateActionRaw)) { 'Add' } else { $delegateActionRaw }
        if ($delegateAction -notin @('Add', 'Remove')) {
            throw "DelegateAction '$delegateAction' is invalid. Use Add or Remove."
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup resource mailbox $resourceMailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $resourceMailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$resourceMailboxIdentity|$delegateIdentity" -Action 'SetResourceBookingDelegate' -Status 'NotFound' -Message 'Resource mailbox not found.'))
            $rowNumber++
            continue
        }

        if ((Get-TrimmedValue -Value $mailbox.RecipientTypeDetails) -notin @('RoomMailbox', 'EquipmentMailbox')) {
            throw "Recipient '$resourceMailboxIdentity' is '$($mailbox.RecipientTypeDetails)'. Expected RoomMailbox or EquipmentMailbox."
        }

        $delegateRecipient = Invoke-WithRetry -OperationName "Lookup delegate recipient $delegateIdentity" -ScriptBlock {
            Get-Recipient -Identity $delegateIdentity -ErrorAction SilentlyContinue
        }
        if (-not $delegateRecipient) {
            throw "Delegate recipient '$delegateIdentity' was not found."
        }

        $delegateKey = Get-RecipientKey -Recipient $delegateRecipient

        $calendar = Invoke-WithRetry -OperationName "Load calendar processing for $resourceMailboxIdentity" -ScriptBlock {
            Get-CalendarProcessing -Identity $mailbox.Identity -ErrorAction Stop
        }

        $currentDelegates = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($current in @($calendar.ResourceDelegates)) {
            $recipient = Invoke-WithRetry -OperationName "Resolve current booking delegate $current" -ScriptBlock {
                Get-Recipient -Identity $current -ErrorAction SilentlyContinue
            }

            if ($recipient) {
                $key = Get-RecipientKey -Recipient $recipient
                if (-not [string]::IsNullOrWhiteSpace($key)) {
                    $null = $currentDelegates.Add($key)
                }
            }
            else {
                $raw = Get-TrimmedValue -Value $current
                if (-not [string]::IsNullOrWhiteSpace($raw)) {
                    $null = $currentDelegates.Add($raw.ToLowerInvariant())
                }
            }
        }

        if ($delegateAction -eq 'Add') {
            if ($currentDelegates.Contains($delegateKey)) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$resourceMailboxIdentity|$delegateIdentity|Add" -Action 'SetResourceBookingDelegate' -Status 'Skipped' -Message 'Delegate already exists.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$resourceMailboxIdentity -> $delegateIdentity", 'Add booking delegate')) {
                Invoke-WithRetry -OperationName "Add booking delegate $resourceMailboxIdentity -> $delegateIdentity" -ScriptBlock {
                    Set-CalendarProcessing -Identity $mailbox.Identity -ResourceDelegates @{ Add = $delegateRecipient.Identity } -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$resourceMailboxIdentity|$delegateIdentity|Add" -Action 'SetResourceBookingDelegate' -Status 'Added' -Message 'Booking delegate added.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$resourceMailboxIdentity|$delegateIdentity|Add" -Action 'SetResourceBookingDelegate' -Status 'WhatIf' -Message 'Booking delegate update skipped due to WhatIf.'))
            }
        }
        else {
            if (-not $currentDelegates.Contains($delegateKey)) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$resourceMailboxIdentity|$delegateIdentity|Remove" -Action 'SetResourceBookingDelegate' -Status 'Skipped' -Message 'Delegate does not exist.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$resourceMailboxIdentity -> $delegateIdentity", 'Remove booking delegate')) {
                Invoke-WithRetry -OperationName "Remove booking delegate $resourceMailboxIdentity -> $delegateIdentity" -ScriptBlock {
                    Set-CalendarProcessing -Identity $mailbox.Identity -ResourceDelegates @{ Remove = $delegateRecipient.Identity } -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$resourceMailboxIdentity|$delegateIdentity|Remove" -Action 'SetResourceBookingDelegate' -Status 'Removed' -Message 'Booking delegate removed.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$resourceMailboxIdentity|$delegateIdentity|Remove" -Action 'SetResourceBookingDelegate' -Status 'WhatIf' -Message 'Booking delegate update skipped due to WhatIf.'))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($resourceMailboxIdentity|$delegateIdentity|$delegateActionRaw) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$resourceMailboxIdentity|$delegateIdentity|$delegateActionRaw" -Action 'SetResourceBookingDelegate' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem resource mailbox booking delegate script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
