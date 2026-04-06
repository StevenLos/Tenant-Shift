<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-164500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineDistributionLists and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineDistributionLists from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-D3114-Get-ExchangeOnlineDistributionLists.ps1 -InputCsvPath .\3114.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D3114-Get-ExchangeOnlineDistributionLists.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-EXOL-0040-Get-ExchangeOnlineDistributionLists_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'DistributionGroupIdentity'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'DistributionGroupIdentity',
    'DisplayName',
    'Name',
    'Alias',
    'PrimarySmtpAddress',
    'ManagedBy',
    'Notes',
    'MemberJoinRestriction',
    'MemberDepartRestriction',
    'ModerationEnabled',
    'ModeratedBy',
    'RequireSenderAuthenticationEnabled',
    'HiddenFromAddressListsEnabled',
    'AcceptMessagesOnlyFrom',
    'AcceptMessagesOnlyFromDLMembers',
    'RejectMessagesFrom',
    'RejectMessagesFromDLMembers',
    'BypassModerationFromSendersOrMembers',
    'SendModerationNotifications'
)

Write-Status -Message 'Starting Exchange Online distribution list inventory script.'
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
    $distributionGroupIdentity = ([string]$row.DistributionGroupIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($distributionGroupIdentity)) {
            throw 'DistributionGroupIdentity is required. Use * to inventory all distribution lists.'
        }

        $groups = @()
        if ($distributionGroupIdentity -eq '*') {
            $groups = @(Invoke-WithRetry -OperationName 'Load all distribution lists' -ScriptBlock {
                Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop
            })
        }
        else {
            $group = Invoke-WithRetry -OperationName "Lookup distribution list $distributionGroupIdentity" -ScriptBlock {
                Get-DistributionGroup -Identity $distributionGroupIdentity -ErrorAction SilentlyContinue
            }

            if ($group) {
                $groups = @($group)
            }
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'GetExchangeDistributionList' -Status 'NotFound' -Message 'No matching distribution lists were found.' -Data ([ordered]@{
                        DistributionGroupIdentity                 = $distributionGroupIdentity
                        Name                                      = ''
                        Alias                                     = ''
                        DisplayName                               = ''
                        PrimarySmtpAddress                        = ''
                        ManagedBy                                 = ''
                        Notes                                     = ''
                        MemberJoinRestriction                     = ''
                        MemberDepartRestriction                   = ''
                        ModerationEnabled                         = ''
                        ModeratedBy                               = ''
                        RequireSenderAuthenticationEnabled        = ''
                        HiddenFromAddressListsEnabled             = ''
                        AcceptMessagesOnlyFrom                    = ''
                        AcceptMessagesOnlyFromDLMembers           = ''
                        RejectMessagesFrom                        = ''
                        RejectMessagesFromDLMembers               = ''
                        BypassModerationFromSendersOrMembers      = ''
                        SendModerationNotifications               = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Identity)) {
            $identity = Get-StringPropertyValue -InputObject $group -PropertyName 'Identity'
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeDistributionList' -Status 'Completed' -Message 'Distribution list exported.' -Data ([ordered]@{
                        DistributionGroupIdentity                 = $identity
                        Name                                      = Get-StringPropertyValue -InputObject $group -PropertyName 'Name'
                        Alias                                     = Get-StringPropertyValue -InputObject $group -PropertyName 'Alias'
                        DisplayName                               = Get-StringPropertyValue -InputObject $group -PropertyName 'DisplayName'
                        PrimarySmtpAddress                        = Get-StringPropertyValue -InputObject $group -PropertyName 'PrimarySmtpAddress'
                        ManagedBy                                 = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'ManagedBy')
                        Notes                                     = Get-StringPropertyValue -InputObject $group -PropertyName 'Notes'
                        MemberJoinRestriction                     = Get-StringPropertyValue -InputObject $group -PropertyName 'MemberJoinRestriction'
                        MemberDepartRestriction                   = Get-StringPropertyValue -InputObject $group -PropertyName 'MemberDepartRestriction'
                        ModerationEnabled                         = [string](Get-ObjectPropertyValue -InputObject $group -PropertyName 'ModerationEnabled')
                        ModeratedBy                               = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'ModeratedBy')
                        RequireSenderAuthenticationEnabled        = [string](Get-ObjectPropertyValue -InputObject $group -PropertyName 'RequireSenderAuthenticationEnabled')
                        HiddenFromAddressListsEnabled             = [string](Get-ObjectPropertyValue -InputObject $group -PropertyName 'HiddenFromAddressListsEnabled')
                        AcceptMessagesOnlyFrom                    = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'AcceptMessagesOnlyFrom')
                        AcceptMessagesOnlyFromDLMembers           = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'AcceptMessagesOnlyFromDLMembers')
                        RejectMessagesFrom                        = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'RejectMessagesFrom')
                        RejectMessagesFromDLMembers               = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'RejectMessagesFromDLMembers')
                        BypassModerationFromSendersOrMembers      = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'BypassModerationFromSendersOrMembers')
                        SendModerationNotifications               = Get-StringPropertyValue -InputObject $group -PropertyName 'SendModerationNotifications'
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($distributionGroupIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'GetExchangeDistributionList' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    DistributionGroupIdentity                 = $distributionGroupIdentity
                    Name                                      = ''
                    Alias                                     = ''
                    DisplayName                               = ''
                    PrimarySmtpAddress                        = ''
                    ManagedBy                                 = ''
                    Notes                                     = ''
                    MemberJoinRestriction                     = ''
                    MemberDepartRestriction                   = ''
                    ModerationEnabled                         = ''
                    ModeratedBy                               = ''
                    RequireSenderAuthenticationEnabled        = ''
                    HiddenFromAddressListsEnabled             = ''
                    AcceptMessagesOnlyFrom                    = ''
                    AcceptMessagesOnlyFromDLMembers           = ''
                    RejectMessagesFrom                        = ''
                    RejectMessagesFromDLMembers               = ''
                    BypassModerationFromSendersOrMembers      = ''
                    SendModerationNotifications               = ''
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
Write-Status -Message 'Exchange Online distribution list inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}









