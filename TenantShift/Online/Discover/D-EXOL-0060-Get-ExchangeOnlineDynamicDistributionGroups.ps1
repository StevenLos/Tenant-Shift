<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-175000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineDynamicDistributionGroups and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineDynamicDistributionGroups from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-D3123-Get-ExchangeOnlineDynamicDistributionGroups.ps1 -InputCsvPath .\3123.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D3123-Get-ExchangeOnlineDynamicDistributionGroups.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-EXOL-0060-Get-ExchangeOnlineDynamicDistributionGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$requiredHeaders = @(
    'DynamicDistributionGroupIdentity'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'DynamicDistributionGroupIdentity',
    'DisplayName',
    'Name',
    'Alias',
    'PrimarySmtpAddress',
    'ManagedBy',
    'RecipientFilter',
    'IncludedRecipients',
    'ConditionalCompany',
    'ConditionalDepartment',
    'ConditionalCustomAttribute1',
    'ConditionalCustomAttribute2',
    'ConditionalStateOrProvince',
    'RequireSenderAuthenticationEnabled',
    'HiddenFromAddressListsEnabled',
    'ModerationEnabled',
    'ModeratedBy',
    'SendModerationNotifications'
)

Write-Status -Message 'Starting Exchange Online dynamic distribution group inventory script.'
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
    $dynamicDistributionGroupIdentity = ([string]$row.DynamicDistributionGroupIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($dynamicDistributionGroupIdentity)) {
            throw 'DynamicDistributionGroupIdentity is required. Use * to inventory all dynamic distribution groups.'
        }

        $groups = @()
        if ($dynamicDistributionGroupIdentity -eq '*') {
            $groups = @(Invoke-WithRetry -OperationName 'Load all dynamic distribution groups' -ScriptBlock {
                Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction Stop
            })
        }
        else {
            $group = Invoke-WithRetry -OperationName "Lookup dynamic distribution group $dynamicDistributionGroupIdentity" -ScriptBlock {
                Get-DynamicDistributionGroup -Identity $dynamicDistributionGroupIdentity -ErrorAction SilentlyContinue
            }
            if ($group) {
                $groups = @($group)
            }
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $dynamicDistributionGroupIdentity -Action 'GetExchangeDynamicDistributionGroup' -Status 'NotFound' -Message 'No matching dynamic distribution groups were found.' -Data ([ordered]@{
                        DynamicDistributionGroupIdentity        = $dynamicDistributionGroupIdentity
                        Name                                    = ''
                        Alias                                   = ''
                        DisplayName                             = ''
                        PrimarySmtpAddress                      = ''
                        ManagedBy                               = ''
                        RecipientFilter                         = ''
                        IncludedRecipients                      = ''
                        ConditionalCompany                      = ''
                        ConditionalDepartment                   = ''
                        ConditionalCustomAttribute1             = ''
                        ConditionalCustomAttribute2             = ''
                        ConditionalStateOrProvince              = ''
                        RequireSenderAuthenticationEnabled      = ''
                        HiddenFromAddressListsEnabled           = ''
                        ModerationEnabled                       = ''
                        ModeratedBy                             = ''
                        SendModerationNotifications             = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Identity)) {
            $identity = ([string]$group.Identity).Trim()
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeDynamicDistributionGroup' -Status 'Completed' -Message 'Dynamic distribution group exported.' -Data ([ordered]@{
                        DynamicDistributionGroupIdentity        = $identity
                        Name                                    = ([string]$group.Name).Trim()
                        Alias                                   = ([string]$group.Alias).Trim()
                        DisplayName                             = ([string]$group.DisplayName).Trim()
                        PrimarySmtpAddress                      = ([string]$group.PrimarySmtpAddress).Trim()
                        ManagedBy                               = Convert-MultiValueToString -Value $group.ManagedBy
                        RecipientFilter                         = ([string]$group.RecipientFilter).Trim()
                        IncludedRecipients                      = ([string]$group.IncludedRecipients).Trim()
                        ConditionalCompany                      = Convert-MultiValueToString -Value $group.ConditionalCompany
                        ConditionalDepartment                   = Convert-MultiValueToString -Value $group.ConditionalDepartment
                        ConditionalCustomAttribute1             = Convert-MultiValueToString -Value $group.ConditionalCustomAttribute1
                        ConditionalCustomAttribute2             = Convert-MultiValueToString -Value $group.ConditionalCustomAttribute2
                        ConditionalStateOrProvince              = Convert-MultiValueToString -Value $group.ConditionalStateOrProvince
                        RequireSenderAuthenticationEnabled      = [string]$group.RequireSenderAuthenticationEnabled
                        HiddenFromAddressListsEnabled           = [string]$group.HiddenFromAddressListsEnabled
                        ModerationEnabled                       = [string]$group.ModerationEnabled
                        ModeratedBy                             = Convert-MultiValueToString -Value $group.ModeratedBy
                        SendModerationNotifications             = ([string]$group.SendModerationNotifications).Trim()
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($dynamicDistributionGroupIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $dynamicDistributionGroupIdentity -Action 'GetExchangeDynamicDistributionGroup' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    DynamicDistributionGroupIdentity        = $dynamicDistributionGroupIdentity
                    Name                                    = ''
                    Alias                                   = ''
                    DisplayName                             = ''
                    PrimarySmtpAddress                      = ''
                    ManagedBy                               = ''
                    RecipientFilter                         = ''
                    IncludedRecipients                      = ''
                    ConditionalCompany                      = ''
                    ConditionalDepartment                   = ''
                    ConditionalCustomAttribute1             = ''
                    ConditionalCustomAttribute2             = ''
                    ConditionalStateOrProvince              = ''
                    RequireSenderAuthenticationEnabled      = ''
                    HiddenFromAddressListsEnabled           = ''
                    ModerationEnabled                       = ''
                    ModeratedBy                             = ''
                    SendModerationNotifications             = ''
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
Write-Status -Message 'Exchange Online dynamic distribution group inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}









