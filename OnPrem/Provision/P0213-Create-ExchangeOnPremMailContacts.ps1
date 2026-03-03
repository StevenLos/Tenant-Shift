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
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_P0213-Create-ExchangeOnPremMailContacts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
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
    'ExternalEmailAddress',
    'Alias',
    'DisplayName',
    'FirstName',
    'LastName',
    'PrimarySmtpAddress',
    'HiddenFromAddressListsEnabled'
)

Write-Status -Message 'Starting Exchange on-prem mail contact creation script.'
Ensure-ExchangeOnPremConnection

$newMailContactCommand = Get-Command -Name New-MailContact -ErrorAction Stop
$supportsOrganizationalUnit = $newMailContactCommand.Parameters.ContainsKey('OrganizationalUnit')

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
    $externalEmailAddress = Get-TrimmedValue -Value $row.ExternalEmailAddress
    $name = Get-TrimmedValue -Value $row.Name
    $alias = Get-TrimmedValue -Value $row.Alias

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
        $primarySmtpAddress = Get-TrimmedValue -Value $row.PrimarySmtpAddress
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
            $existingIdentity = Get-TrimmedValue -Value $existingContact.Identity
            if ([string]::IsNullOrWhiteSpace($existingIdentity)) {
                $existingIdentity = Get-TrimmedValue -Value $existingContact.Name
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

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $params.DisplayName = $displayName
        }

        $firstName = Get-TrimmedValue -Value $row.FirstName
        if (-not [string]::IsNullOrWhiteSpace($firstName)) {
            $params.FirstName = $firstName
        }

        $lastName = Get-TrimmedValue -Value $row.LastName
        if (-not [string]::IsNullOrWhiteSpace($lastName)) {
            $params.LastName = $lastName
        }

        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $params.PrimarySmtpAddress = $primarySmtpAddress
        }

        $organizationalUnit = Get-OptionalColumnValue -Row $row -ColumnName 'OrganizationalUnit'
        if ($supportsOrganizationalUnit -and -not [string]::IsNullOrWhiteSpace($organizationalUnit)) {
            $params.OrganizationalUnit = $organizationalUnit
        }

        $hiddenRaw = Get-TrimmedValue -Value $row.HiddenFromAddressListsEnabled
        $setHidden = -not [string]::IsNullOrWhiteSpace($hiddenRaw)

        if ($PSCmdlet.ShouldProcess($externalEmailAddress, 'Create Exchange on-prem mail contact')) {
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

            $createdAlias = Get-TrimmedValue -Value $createdContact.Alias
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
Write-Status -Message 'Exchange on-prem mail contact creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
