<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-180000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineRecipientTypeCounts and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineRecipientTypeCounts from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-IR3125-Get-ExchangeOnlineRecipientTypeCounts.ps1 -InputCsvPath .\3125.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3125-Get-ExchangeOnlineRecipientTypeCounts.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator
    Limitations:      None known.

    CSV Fields:
    Column                Type      Required  Description
    --------------------  ----      --------  -----------
    RecipientIdentity     String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-EXOL-0210-Get-ExchangeOnlineRecipientTypeCounts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'RecipientIdentity'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'RecipientIdentityInput',
    'RecipientTypeDetails',
    'RecipientCount',
    'ScopeRecipientCountTotal',
    'DistinctPrimarySmtpAddressCount'
)

Write-Status -Message 'Starting Exchange Online recipient type-count inventory script.'
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
    $recipientIdentityRaw = Get-TrimmedValue -Value $row.RecipientIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($recipientIdentityRaw)) {
            throw 'RecipientIdentity is required. Use * to count all recipients.'
        }

        $recipients = [System.Collections.Generic.List[object]]::new()

        if ($recipientIdentityRaw -eq '*') {
            $allRecipients = @(Invoke-WithRetry -OperationName 'Load all recipients for type count inventory' -ScriptBlock {
                Get-ExchangeOnlineRecipient -ResultSize Unlimited -ErrorAction Stop
            })

            foreach ($recipient in $allRecipients) {
                $recipients.Add($recipient)
            }
        }
        else {
            $identities = ConvertTo-Array -Value $recipientIdentityRaw
            if ($identities.Count -eq 0) {
                throw 'RecipientIdentity did not contain any usable identities.'
            }

            foreach ($identity in $identities) {
                $resolved = Invoke-WithRetry -OperationName "Lookup recipient $identity" -ScriptBlock {
                    Get-ExchangeOnlineRecipient -Identity $identity -ErrorAction SilentlyContinue
                }

                if ($resolved) {
                    $recipients.Add($resolved)
                }
                else {
                    $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeRecipientTypeCounts' -Status 'NotFound' -Message 'Recipient was not found.' -Data ([ordered]@{
                                RecipientIdentityInput            = $recipientIdentityRaw
                                RecipientTypeDetails              = ''
                                RecipientCount                    = ''
                                ScopeRecipientCountTotal          = ''
                                DistinctPrimarySmtpAddressCount   = ''
                            })))
                }
            }
        }

        if ($recipients.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $recipientIdentityRaw -Action 'GetExchangeRecipientTypeCounts' -Status 'NotFound' -Message 'No matching recipients were found.' -Data ([ordered]@{
                        RecipientIdentityInput            = $recipientIdentityRaw
                        RecipientTypeDetails              = ''
                        RecipientCount                    = ''
                        ScopeRecipientCountTotal          = ''
                        DistinctPrimarySmtpAddressCount   = ''
                    })))
            $rowNumber++
            continue
        }

        $grouped = @($recipients | Group-Object -Property RecipientTypeDetails | Sort-Object -Property Name)
        foreach ($group in $grouped) {
            $primarySmtpDistinct = @(
                $group.Group |
                    ForEach-Object { ([string]$_.PrimarySmtpAddress).Trim().ToLowerInvariant() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Sort-Object -Unique
            ).Count

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$recipientIdentityRaw|$($group.Name)" -Action 'GetExchangeRecipientTypeCounts' -Status 'Completed' -Message 'Recipient type count exported.' -Data ([ordered]@{
                        RecipientIdentityInput            = $recipientIdentityRaw
                        RecipientTypeDetails              = ([string]$group.Name).Trim()
                        RecipientCount                    = [string]$group.Count
                        ScopeRecipientCountTotal          = [string]$recipients.Count
                        DistinctPrimarySmtpAddressCount   = [string]$primarySmtpDistinct
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($recipientIdentityRaw) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $recipientIdentityRaw -Action 'GetExchangeRecipientTypeCounts' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    RecipientIdentityInput            = $recipientIdentityRaw
                    RecipientTypeDetails              = ''
                    RecipientCount                    = ''
                    ScopeRecipientCountTotal          = ''
                    DistinctPrimarySmtpAddressCount   = ''
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
Write-Status -Message 'Exchange Online recipient type-count inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
