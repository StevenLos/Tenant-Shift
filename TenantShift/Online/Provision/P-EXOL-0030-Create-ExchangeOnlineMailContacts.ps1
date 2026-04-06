<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-004416

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Provisions ExchangeOnlineMailContacts in Microsoft 365.

.DESCRIPTION
    Creates ExchangeOnlineMailContacts in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P3113-Create-ExchangeOnlineMailContacts.ps1 -InputCsvPath .\3113.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P3113-Create-ExchangeOnlineMailContacts.ps1 -InputCsvPath .\3113.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator
    Limitations:      None known.

    CSV Fields:
    Column                         Type      Required  Description
    -----------------------------  ----      --------  -----------
    Name                           String    Yes       <fill in description>
    ExternalEmailAddress           String    Yes       <fill in description>
    Alias                          String    Yes       <fill in description>
    DisplayName                    String    Yes       <fill in description>
    FirstName                      String    Yes       <fill in description>
    LastName                       String    Yes       <fill in description>
    PrimarySmtpAddress             String    Yes       <fill in description>
    HiddenFromAddressListsEnabled  String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P3113-Create-ExchangeOnlineMailContacts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {


function ConvertTo-NormalizedSmtpAddress {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    $trimmed = $Value.Trim()
    if ($trimmed.StartsWith('SMTP:', [System.StringComparison]::OrdinalIgnoreCase)) {
        $trimmed = $trimmed.Substring(5)
    }

    return $trimmed.ToLowerInvariant()
}

$requiredHeaders = @(
    'Name',
    'ExternalEmailAddress',
    'Alias',
    'DisplayName',
    'FirstName',
    'LastName',
    'PrimarySmtpAddress',
    'HiddenFromAddressListsEnabled'
)

Write-Status -Message 'Starting Exchange Online mail contact creation script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

Write-Status -Message 'Loading existing mail contacts for idempotent checks.'
$existingContacts = @(Invoke-WithRetry -OperationName 'Load existing mail contacts' -ScriptBlock {
    Get-MailContact -ResultSize Unlimited -ErrorAction Stop
})

$contactsByExternalEmail = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$contactsByPrimarySmtp = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$contactsByAlias = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($contact in $existingContacts) {
    $externalKey = ConvertTo-NormalizedSmtpAddress -Value ([string]$contact.ExternalEmailAddress)
    if (-not [string]::IsNullOrWhiteSpace($externalKey) -and -not $contactsByExternalEmail.ContainsKey($externalKey)) {
        $contactsByExternalEmail[$externalKey] = $contact
    }

    $primaryKey = ConvertTo-NormalizedSmtpAddress -Value ([string]$contact.PrimarySmtpAddress)
    if (-not [string]::IsNullOrWhiteSpace($primaryKey) -and -not $contactsByPrimarySmtp.ContainsKey($primaryKey)) {
        $contactsByPrimarySmtp[$primaryKey] = $contact
    }

    $aliasKey = ([string]$contact.Alias).Trim()
    if (-not [string]::IsNullOrWhiteSpace($aliasKey) -and -not $contactsByAlias.ContainsKey($aliasKey)) {
        $contactsByAlias[$aliasKey] = $contact
    }
}

$rowNumber = 1
foreach ($row in $rows) {
    $externalEmailAddress = ([string]$row.ExternalEmailAddress).Trim()
    $name = ([string]$row.Name).Trim()
    $alias = ([string]$row.Alias).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($externalEmailAddress)) {
            throw 'ExternalEmailAddress is required.'
        }

        if ([string]::IsNullOrWhiteSpace($name)) {
            $name = $externalEmailAddress.Split('@')[0]
        }

        if ([string]::IsNullOrWhiteSpace($alias)) {
            $alias = (($externalEmailAddress.Split('@')[0]) -replace '[^a-zA-Z0-9._-]', '')
        }

        if ([string]::IsNullOrWhiteSpace($alias)) {
            throw 'Alias is empty after sanitization. Provide an Alias value in the CSV.'
        }

        $externalKey = ConvertTo-NormalizedSmtpAddress -Value $externalEmailAddress
        $primarySmtpAddress = ([string]$row.PrimarySmtpAddress).Trim()
        $primarySmtpKey = ConvertTo-NormalizedSmtpAddress -Value $primarySmtpAddress

        $existingContact = $null
        if ($contactsByExternalEmail.ContainsKey($externalKey)) {
            $existingContact = $contactsByExternalEmail[$externalKey]
        }
        elseif (-not [string]::IsNullOrWhiteSpace($primarySmtpKey) -and $contactsByPrimarySmtp.ContainsKey($primarySmtpKey)) {
            $existingContact = $contactsByPrimarySmtp[$primarySmtpKey]
        }
        elseif ($contactsByAlias.ContainsKey($alias)) {
            $existingContact = $contactsByAlias[$alias]
        }

        if ($existingContact) {
            $existingIdentity = ([string]$existingContact.Identity).Trim()
            if ([string]::IsNullOrWhiteSpace($existingIdentity)) {
                $existingIdentity = ([string]$existingContact.Name).Trim()
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $externalEmailAddress -Action 'CreateMailContact' -Status 'Skipped' -Message "Mail contact already exists (identity '$existingIdentity')."))
            $rowNumber++
            continue
        }

        $params = @{
            Name                 = $name
            ExternalEmailAddress = $externalEmailAddress
            Alias                = $alias
        }

        $displayName = ([string]$row.DisplayName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $params.DisplayName = $displayName
        }

        $firstName = ([string]$row.FirstName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($firstName)) {
            $params.FirstName = $firstName
        }

        $lastName = ([string]$row.LastName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($lastName)) {
            $params.LastName = $lastName
        }

        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $params.PrimarySmtpAddress = $primarySmtpAddress
        }

        $hiddenRaw = ([string]$row.HiddenFromAddressListsEnabled).Trim()
        $setHidden = -not [string]::IsNullOrWhiteSpace($hiddenRaw)

        if ($PSCmdlet.ShouldProcess($externalEmailAddress, 'Create Exchange Online mail contact')) {
            $createdContact = Invoke-WithRetry -OperationName "Create mail contact $externalEmailAddress" -ScriptBlock {
                New-MailContact @params -ErrorAction Stop
            }

            if ($setHidden) {
                $hiddenValue = ConvertTo-Bool -Value $hiddenRaw
                Invoke-WithRetry -OperationName "Set hidden from GAL for $externalEmailAddress" -ScriptBlock {
                    Set-MailContact -Identity $createdContact.Identity -HiddenFromAddressListsEnabled $hiddenValue -ErrorAction Stop
                }
            }

            $createdExternalKey = ConvertTo-NormalizedSmtpAddress -Value ([string]$createdContact.ExternalEmailAddress)
            if (-not [string]::IsNullOrWhiteSpace($createdExternalKey) -and -not $contactsByExternalEmail.ContainsKey($createdExternalKey)) {
                $contactsByExternalEmail[$createdExternalKey] = $createdContact
            }

            $createdPrimaryKey = ConvertTo-NormalizedSmtpAddress -Value ([string]$createdContact.PrimarySmtpAddress)
            if (-not [string]::IsNullOrWhiteSpace($createdPrimaryKey) -and -not $contactsByPrimarySmtp.ContainsKey($createdPrimaryKey)) {
                $contactsByPrimarySmtp[$createdPrimaryKey] = $createdContact
            }

            $createdAlias = ([string]$createdContact.Alias).Trim()
            if (-not [string]::IsNullOrWhiteSpace($createdAlias) -and -not $contactsByAlias.ContainsKey($createdAlias)) {
                $contactsByAlias[$createdAlias] = $createdContact
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $externalEmailAddress -Action 'CreateMailContact' -Status 'Created' -Message 'Mail contact created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $externalEmailAddress -Action 'CreateMailContact' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($externalEmailAddress) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $externalEmailAddress -Action 'CreateMailContact' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mail contact creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}








