<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-164000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineMailContacts and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineMailContacts from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-IR3113-Get-ExchangeOnlineMailContacts.ps1 -InputCsvPath .\3113.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3113-Get-ExchangeOnlineMailContacts.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-EXOL-0030-Get-ExchangeOnlineMailContacts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'MailContactIdentity'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'MailContactIdentity',
    'DisplayName',
    'Name',
    'Alias',
    'ExternalEmailAddress',
    'PrimarySmtpAddress',
    'FirstName',
    'LastName',
    'HiddenFromAddressListsEnabled',
    'WhenCreatedUTC'
)

Write-Status -Message 'Starting Exchange Online mail contact inventory script.'
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
    $mailContactIdentity = ([string]$row.MailContactIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($mailContactIdentity)) {
            throw 'MailContactIdentity is required. Use * to inventory all mail contacts.'
        }

        $contacts = @()
        if ($mailContactIdentity -eq '*') {
            $contacts = @(Invoke-WithRetry -OperationName 'Load all mail contacts' -ScriptBlock {
                Get-MailContact -ResultSize Unlimited -ErrorAction Stop
            })
        }
        else {
            $contact = Invoke-WithRetry -OperationName "Lookup mail contact $mailContactIdentity" -ScriptBlock {
                Get-MailContact -Identity $mailContactIdentity -ErrorAction SilentlyContinue
            }
            if ($contact) {
                $contacts = @($contact)
            }
        }

        if ($contacts.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'GetExchangeMailContact' -Status 'NotFound' -Message 'No matching mail contacts were found.' -Data ([ordered]@{
                        MailContactIdentity            = $mailContactIdentity
                        Name                           = ''
                        Alias                          = ''
                        DisplayName                    = ''
                        ExternalEmailAddress           = ''
                        PrimarySmtpAddress             = ''
                        FirstName                      = ''
                        LastName                       = ''
                        HiddenFromAddressListsEnabled  = ''
                        WhenCreatedUTC                 = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($contact in @($contacts | Sort-Object -Property DisplayName, Identity)) {
            $identity = ([string]$contact.Identity).Trim()
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeMailContact' -Status 'Completed' -Message 'Mail contact exported.' -Data ([ordered]@{
                        MailContactIdentity            = $identity
                        Name                           = ([string]$contact.Name).Trim()
                        Alias                          = ([string]$contact.Alias).Trim()
                        DisplayName                    = ([string]$contact.DisplayName).Trim()
                        ExternalEmailAddress           = ([string]$contact.ExternalEmailAddress).Trim()
                        PrimarySmtpAddress             = ([string]$contact.PrimarySmtpAddress).Trim()
                        FirstName                      = if ($null -ne $contact.PSObject.Properties['FirstName']) { ([string]$contact.FirstName).Trim() } else { '' }
                        LastName                       = if ($null -ne $contact.PSObject.Properties['LastName']) { ([string]$contact.LastName).Trim() } else { '' }
                        HiddenFromAddressListsEnabled  = [string]$contact.HiddenFromAddressListsEnabled
                        WhenCreatedUTC                 = [string]$contact.WhenCreatedUTC
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailContactIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'GetExchangeMailContact' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    MailContactIdentity            = $mailContactIdentity
                    Name                           = ''
                    Alias                          = ''
                    DisplayName                    = ''
                    ExternalEmailAddress           = ''
                    PrimarySmtpAddress             = ''
                    FirstName                      = ''
                    LastName                       = ''
                    HiddenFromAddressListsEnabled  = ''
                    WhenCreatedUTC                 = ''
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
Write-Status -Message 'Exchange Online mail contact inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}









