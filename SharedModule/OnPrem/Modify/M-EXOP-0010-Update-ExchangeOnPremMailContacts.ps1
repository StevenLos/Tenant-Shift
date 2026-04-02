<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-000100

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)

.SYNOPSIS
    Modifies ExchangeOnPremMailContacts in Active Directory.

.DESCRIPTION
    Updates ExchangeOnPremMailContacts in Active Directory based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M0213-Update-ExchangeOnPremMailContacts.ps1 -InputCsvPath .\0213.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M0213-Update-ExchangeOnPremMailContacts.ps1 -InputCsvPath .\0213.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Exchange Management Shell cmdlets (session)
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    Column                         Type      Required  Description
    -----------------------------  ----      --------  -----------
    MailContactIdentity            String    Yes       <fill in description>
    Name                           String    Yes       <fill in description>
    DisplayName                    String    Yes       <fill in description>
    FirstName                      String    Yes       <fill in description>
    LastName                       String    Yes       <fill in description>
    ExternalEmailAddress           String    Yes       <fill in description>
    PrimarySmtpAddress             String    Yes       <fill in description>
    HiddenFromAddressListsEnabled  String    Yes       <fill in description>
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M0213-Update-ExchangeOnPremMailContacts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'MailContactIdentity',
    'Name',
    'DisplayName',
    'FirstName',
    'LastName',
    'ExternalEmailAddress',
    'PrimarySmtpAddress',
    'HiddenFromAddressListsEnabled'
)

Write-Status -Message 'Starting Exchange on-prem mail contact update script.'
Ensure-ExchangeOnPremConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailContactIdentity = Get-TrimmedValue -Value $row.MailContactIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($mailContactIdentity)) {
            throw 'MailContactIdentity is required.'
        }

        $contact = Invoke-WithRetry -OperationName "Lookup mail contact $mailContactIdentity" -ScriptBlock {
            Get-MailContact -Identity $mailContactIdentity -ErrorAction SilentlyContinue
        }
        if (-not $contact) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'UpdateMailContact' -Status 'NotFound' -Message 'Mail contact not found.'))
            $rowNumber++
            continue
        }

        $setParams = @{
            Identity = $contact.Identity
        }

        $name = Get-TrimmedValue -Value $row.Name
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $setParams.Name = $name
        }

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $setParams.DisplayName = $displayName
        }

        $firstName = Get-TrimmedValue -Value $row.FirstName
        if (-not [string]::IsNullOrWhiteSpace($firstName)) {
            $setParams.FirstName = $firstName
        }

        $lastName = Get-TrimmedValue -Value $row.LastName
        if (-not [string]::IsNullOrWhiteSpace($lastName)) {
            $setParams.LastName = $lastName
        }

        $externalEmailAddress = Get-TrimmedValue -Value $row.ExternalEmailAddress
        if (-not [string]::IsNullOrWhiteSpace($externalEmailAddress)) {
            $setParams.ExternalEmailAddress = $externalEmailAddress
        }

        $primarySmtpAddress = Get-TrimmedValue -Value $row.PrimarySmtpAddress
        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $setParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $hiddenRaw = Get-TrimmedValue -Value $row.HiddenFromAddressListsEnabled
        if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
            $setParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
        }

        if ($setParams.Count -eq 1) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'UpdateMailContact' -Status 'Skipped' -Message 'No updates specified.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($mailContactIdentity, 'Update Exchange on-prem mail contact')) {
            Invoke-WithRetry -OperationName "Update mail contact $mailContactIdentity" -ScriptBlock {
                Set-MailContact @setParams -ErrorAction Stop
            }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'UpdateMailContact' -Status 'Updated' -Message 'Mail contact updated successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'UpdateMailContact' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailContactIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'UpdateMailContact' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem mail contact update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
