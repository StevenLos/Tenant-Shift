<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-235500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)
#>
#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$SearchBase,

    [Parameter(ParameterSetName = 'FromCsv')]
    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$Server,

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [ValidateRange(0, 10000000)]
    [int]$MaxObjects = 0,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR0218-Get-ExchangeOnPremResourceMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Add-SupportedParameter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$ParameterHashtable,

        [Parameter(Mandatory)]
        [string]$CommandName,

        [Parameter(Mandatory)]
        [string]$ParameterName,

        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return
    }

    $command = Get-Command -Name $CommandName -ErrorAction Stop
    if ($command.Parameters.ContainsKey($ParameterName)) {
        $ParameterHashtable[$ParameterName] = $text
    }
}

function Resolve-ResourceMailboxesByScope {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,

        [AllowEmptyString()]
        [string]$SearchBase,

        [AllowEmptyString()]
        [string]$Server
    )

    if ($Identity -eq '*') {
        $params = @{
            RecipientTypeDetails = @('RoomMailbox', 'EquipmentMailbox')
            ResultSize           = 'Unlimited'
            ErrorAction          = 'Stop'
        }

        Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-Mailbox' -ParameterName 'OrganizationalUnit' -Value $SearchBase
        Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-Mailbox' -ParameterName 'DomainController' -Value $Server

        return @(Get-Mailbox @params)
    }

    $params = @{
        Identity    = $Identity
        ErrorAction = 'SilentlyContinue'
    }

    Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-Mailbox' -ParameterName 'DomainController' -Value $Server

    $mailbox = Get-Mailbox @params
    if ($mailbox -and (Get-TrimmedValue -Value $mailbox.RecipientTypeDetails) -in @('RoomMailbox', 'EquipmentMailbox')) {
        return @($mailbox)
    }

    return @()
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

Write-Status -Message 'Starting Exchange on-prem resource mailbox inventory script.'
Ensure-ExchangeOnPremConnection

$scopeMode = 'Csv'
$resolvedServer = Get-TrimmedValue -Value $Server
$resolvedSearchBase = ''
$runWasTruncated = $false

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    $resolvedSearchBase = Get-TrimmedValue -Value $SearchBase
    Write-Status -Message "DiscoverAll enabled for Exchange on-prem resource mailboxes. SearchBase='$resolvedSearchBase'." -Level WARN
    $rows = @([PSCustomObject]@{
            ResourceMailboxIdentity = '*'
        })
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $resourceMailboxIdentity = Get-TrimmedValue -Value $row.ResourceMailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($resourceMailboxIdentity)) {
            throw 'ResourceMailboxIdentity is required. Use * to inventory all room/equipment mailboxes.'
        }

        $effectiveSearchBase = if ($scopeMode -eq 'DiscoverAll') { $resolvedSearchBase } else { '' }
        $mailboxes = @(Invoke-WithRetry -OperationName "Load resource mailboxes for $resourceMailboxIdentity" -ScriptBlock {
            Resolve-ResourceMailboxesByScope -Identity $resourceMailboxIdentity -SearchBase $effectiveSearchBase -Server $resolvedServer
        })

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $mailboxes.Count -gt $MaxObjects) {
            $mailboxes = @($mailboxes | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($mailboxes.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'GetExchangeResourceMailbox' -Status 'NotFound' -Message 'No matching resource mailboxes were found.' -Data ([ordered]@{
                        ResourceMailboxIdentity       = $resourceMailboxIdentity
                        ResourceType                  = ''
                        DisplayName                   = ''
                        Alias                         = ''
                        PrimarySmtpAddress            = ''
                        UserPrincipalName             = ''
                        HiddenFromAddressListsEnabled = ''
                        ResourceCapacity              = ''
                        Office                        = ''
                        Phone                         = ''
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

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $identity = Get-TrimmedValue -Value $mailbox.Identity
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
                        ResourceMailboxIdentity       = $identity
                        ResourceType                  = Get-TrimmedValue -Value $mailbox.RecipientTypeDetails
                        DisplayName                   = Get-TrimmedValue -Value $mailbox.DisplayName
                        Alias                         = Get-TrimmedValue -Value $mailbox.Alias
                        PrimarySmtpAddress            = Get-TrimmedValue -Value $mailbox.PrimarySmtpAddress
                        UserPrincipalName             = Get-TrimmedValue -Value $mailbox.UserPrincipalName
                        HiddenFromAddressListsEnabled = [string]$mailbox.HiddenFromAddressListsEnabled
                        ResourceCapacity              = [string]$mailbox.ResourceCapacity
                        Office                        = Get-TrimmedValue -Value $mailbox.Office
                        Phone                         = Get-TrimmedValue -Value $mailbox.Phone
                        AutomateProcessing            = if ($calendar) { Get-TrimmedValue -Value $calendar.AutomateProcessing } else { '' }
                        ForwardRequestsToDelegates    = if ($calendar) { [string]$calendar.ForwardRequestsToDelegates } else { '' }
                        AllBookInPolicy               = if ($calendar) { [string]$calendar.AllBookInPolicy } else { '' }
                        AllRequestInPolicy            = if ($calendar) { [string]$calendar.AllRequestInPolicy } else { '' }
                        AllRequestOutOfPolicy         = if ($calendar) { [string]$calendar.AllRequestOutOfPolicy } else { '' }
                        BookInPolicy                  = if ($calendar) { Convert-MultiValueToString -Value $calendar.BookInPolicy } else { '' }
                        RequestInPolicy               = if ($calendar) { Convert-MultiValueToString -Value $calendar.RequestInPolicy } else { '' }
                        RequestOutOfPolicy            = if ($calendar) { Convert-MultiValueToString -Value $calendar.RequestOutOfPolicy } else { '' }
                        BookingWindowInDays           = if ($calendar) { [string]$calendar.BookingWindowInDays } else { '' }
                        MaximumDurationInMinutes      = if ($calendar) { [string]$calendar.MaximumDurationInMinutes } else { '' }
                        AllowRecurringMeetings        = if ($calendar) { [string]$calendar.AllowRecurringMeetings } else { '' }
                        EnforceSchedulingHorizon      = if ($calendar) { [string]$calendar.EnforceSchedulingHorizon } else { '' }
                        ScheduleOnlyDuringWorkHours   = if ($calendar) { [string]$calendar.ScheduleOnlyDuringWorkHours } else { '' }
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($resourceMailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'GetExchangeResourceMailbox' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    ResourceMailboxIdentity       = $resourceMailboxIdentity
                    ResourceType                  = ''
                    DisplayName                   = ''
                    Alias                         = ''
                    PrimarySmtpAddress            = ''
                    UserPrincipalName             = ''
                    HiddenFromAddressListsEnabled = ''
                    ResourceCapacity              = ''
                    Office                        = ''
                    Phone                         = ''
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
    Add-Member -InputObject $result -NotePropertyName 'ScopeSearchBase' -NotePropertyValue $resolvedSearchBase -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeServer' -NotePropertyValue $resolvedServer -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeMaxObjects' -NotePropertyValue ($(if ($scopeMode -eq 'DiscoverAll') { [string]$MaxObjects } else { '' })) -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeWasTruncated' -NotePropertyValue ([string]$runWasTruncated) -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem resource mailbox inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
