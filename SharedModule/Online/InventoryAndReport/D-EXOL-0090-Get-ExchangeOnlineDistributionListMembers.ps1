<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-165000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineDistributionListMembers and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineDistributionListMembers from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-IR3115-Get-ExchangeOnlineDistributionListMembers.ps1 -InputCsvPath .\3115.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3115-Get-ExchangeOnlineDistributionListMembers.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-EXOL-0090-Get-ExchangeOnlineDistributionListMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'DistributionGroupDisplayName',
    'MemberDisplayName',
    'MemberIdentity',
    'MemberPrimarySmtpAddress',
    'MemberRecipientType',
    'MemberExternalEmailAddress'
)

Write-Status -Message 'Starting Exchange Online distribution list member inventory script.'
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
            throw 'DistributionGroupIdentity is required. Use * to inventory members for all distribution lists.'
        }

        $groups = @()
        if ($distributionGroupIdentity -eq '*') {
            $groups = @(Invoke-WithRetry -OperationName 'Load all distribution lists for membership export' -ScriptBlock {
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
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'GetExchangeDistributionListMember' -Status 'NotFound' -Message 'No matching distribution lists were found.' -Data ([ordered]@{
                        DistributionGroupIdentity         = $distributionGroupIdentity
                        DistributionGroupDisplayName      = ''
                        MemberIdentity                    = ''
                        MemberDisplayName                 = ''
                        MemberPrimarySmtpAddress          = ''
                        MemberRecipientType               = ''
                        MemberExternalEmailAddress        = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Identity)) {
            $groupIdentity = ([string]$group.Identity).Trim()
            $groupDisplayName = ([string]$group.DisplayName).Trim()

            $members = @(Invoke-WithRetry -OperationName "Load members for distribution list $groupIdentity" -ScriptBlock {
                Get-DistributionGroupMember -Identity $groupIdentity -ResultSize Unlimited -ErrorAction Stop
            })

            if ($members.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupIdentity -Action 'GetExchangeDistributionListMember' -Status 'Completed' -Message 'Distribution list has no members.' -Data ([ordered]@{
                            DistributionGroupIdentity         = $groupIdentity
                            DistributionGroupDisplayName      = $groupDisplayName
                            MemberIdentity                    = ''
                            MemberDisplayName                 = ''
                            MemberPrimarySmtpAddress          = ''
                            MemberRecipientType               = ''
                            MemberExternalEmailAddress        = ''
                        })))
                continue
            }

            foreach ($member in @($members | Sort-Object -Property DisplayName, Identity)) {
                $memberIdentity = ([string]$member.Identity).Trim()
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$groupIdentity|$memberIdentity" -Action 'GetExchangeDistributionListMember' -Status 'Completed' -Message 'Distribution list member exported.' -Data ([ordered]@{
                            DistributionGroupIdentity         = $groupIdentity
                            DistributionGroupDisplayName      = $groupDisplayName
                            MemberIdentity                    = $memberIdentity
                            MemberDisplayName                 = ([string]$member.DisplayName).Trim()
                            MemberPrimarySmtpAddress          = ([string]$member.PrimarySmtpAddress).Trim()
                            MemberRecipientType               = ([string]$member.RecipientType).Trim()
                            MemberExternalEmailAddress        = ([string]$member.ExternalEmailAddress).Trim()
                        })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($distributionGroupIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'GetExchangeDistributionListMember' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    DistributionGroupIdentity         = $distributionGroupIdentity
                    DistributionGroupDisplayName      = ''
                    MemberIdentity                    = ''
                    MemberDisplayName                 = ''
                    MemberPrimarySmtpAddress          = ''
                    MemberRecipientType               = ''
                    MemberExternalEmailAddress        = ''
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
Write-Status -Message 'Exchange Online distribution list member inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}












