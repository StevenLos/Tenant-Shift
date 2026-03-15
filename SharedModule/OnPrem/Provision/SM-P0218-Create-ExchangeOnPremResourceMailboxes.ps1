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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P0218-Create-ExchangeOnPremResourceMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-OptionalColumnValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [psobject]$Row,

        [Parameter(Mandatory)]
        [string]$ColumnName
    )

    $property = $Row.PSObject.Properties[$ColumnName]
    if ($null -eq $property) {
        return ''
    }

    return Get-TrimmedValue -Value $property.Value
}

$requiredHeaders = @(
    'Name',
    'Alias',
    'DisplayName',
    'PrimarySmtpAddress',
    'ResourceType',
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

Write-Status -Message 'Starting Exchange on-prem resource mailbox creation script.'
Ensure-ExchangeOnPremConnection

$newMailboxCommand = Get-Command -Name New-Mailbox -ErrorAction Stop
$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop
$setCalendarCommand = Get-Command -Name Set-CalendarProcessing -ErrorAction SilentlyContinue

$supports = @{
    OrganizationalUnit            = $newMailboxCommand.Parameters.ContainsKey('OrganizationalUnit')
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
    $name = Get-TrimmedValue -Value $row.Name

    try {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Name is required.'
        }

        $resourceTypeRaw = Get-TrimmedValue -Value $row.ResourceType
        if ([string]::IsNullOrWhiteSpace($resourceTypeRaw)) {
            throw 'ResourceType is required. Use Room or Equipment.'
        }

        $resourceType = $resourceTypeRaw.ToLowerInvariant()
        if ($resourceType -notin @('room', 'equipment')) {
            throw "ResourceType '$resourceTypeRaw' is invalid. Use Room or Equipment."
        }

        $alias = Get-TrimmedValue -Value $row.Alias
        if ([string]::IsNullOrWhiteSpace($alias)) {
            $alias = ($name -replace '[^a-zA-Z0-9]', '')
            if ([string]::IsNullOrWhiteSpace($alias)) {
                throw 'Alias is empty after sanitization. Provide an Alias value in the CSV.'
            }
        }

        $primarySmtpAddress = Get-TrimmedValue -Value $row.PrimarySmtpAddress
        $identityToCheck = if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) { $primarySmtpAddress } else { $name }

        $existingMailbox = Invoke-WithRetry -OperationName "Lookup resource mailbox $identityToCheck" -ScriptBlock {
            Get-Mailbox -Identity $identityToCheck -ErrorAction SilentlyContinue
        }

        if ($existingMailbox) {
            $recipientTypeDetails = Get-TrimmedValue -Value $existingMailbox.RecipientTypeDetails
            if ($recipientTypeDetails -eq 'RoomMailbox' -or $recipientTypeDetails -eq 'EquipmentMailbox') {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateResourceMailbox' -Status 'Skipped' -Message 'Resource mailbox already exists.'))
                $rowNumber++
                continue
            }

            throw "Recipient '$identityToCheck' already exists with type '$recipientTypeDetails'."
        }

        $createParams = @{
            Name  = $name
            Alias = $alias
        }

        if ($resourceType -eq 'room') {
            $createParams.Room = $true
        }
        else {
            $createParams.Equipment = $true
        }

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $createParams.DisplayName = $displayName
        }

        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $createParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $organizationalUnit = Get-OptionalColumnValue -Row $row -ColumnName 'OrganizationalUnit'
        if ($supports.OrganizationalUnit -and -not [string]::IsNullOrWhiteSpace($organizationalUnit)) {
            $createParams.OrganizationalUnit = $organizationalUnit
        }

        if ($PSCmdlet.ShouldProcess($identityToCheck, 'Create Exchange on-prem resource mailbox')) {
            $createdMailbox = Invoke-WithRetry -OperationName "Create resource mailbox $identityToCheck" -ScriptBlock {
                New-Mailbox @createParams -ErrorAction Stop
            }

            $warnings = [System.Collections.Generic.List[string]]::new()

            $setMailboxParams = @{ Identity = $createdMailbox.Identity }
            $hiddenRaw = Get-TrimmedValue -Value $row.HiddenFromAddressListsEnabled
            if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
                if ($supports.HiddenFromAddressListsEnabled) {
                    $setMailboxParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
                }
                else {
                    $warnings.Add('HiddenFromAddressListsEnabled ignored (unsupported parameter).')
                }
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

            if ($setMailboxParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Set resource mailbox options $identityToCheck" -ScriptBlock {
                    Set-Mailbox @setMailboxParams -ErrorAction Stop
                }
            }

            if ($setCalendarCommand) {
                $setCalendarParams = @{ Identity = $createdMailbox.Identity }

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

                if ($setCalendarParams.Count -gt 1) {
                    Invoke-WithRetry -OperationName "Set resource calendar processing $identityToCheck" -ScriptBlock {
                        Set-CalendarProcessing @setCalendarParams -ErrorAction Stop
                    }
                }
            }
            else {
                $warnings.Add('Set-CalendarProcessing not available; booking settings were not applied.')
            }

            $message = 'Resource mailbox created successfully.'
            if ($warnings.Count -gt 0) {
                $message = "$message $($warnings -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateResourceMailbox' -Status 'Created' -Message $message))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateResourceMailbox' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($name) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $name -Action 'CreateResourceMailbox' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem resource mailbox creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
