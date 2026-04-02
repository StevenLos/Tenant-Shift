<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260302-191500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
ActiveDirectory

.MODULEVERSIONPOLICY
Best-effort PSGallery check (offline-safe)

.SYNOPSIS
    Provisions ActiveDirectoryOrganizationalUnits in Active Directory.

.DESCRIPTION
    Creates ActiveDirectoryOrganizationalUnits in Active Directory based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P0009-Create-ActiveDirectoryOrganizationalUnits.ps1 -InputCsvPath .\0009.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P0009-Create-ActiveDirectoryOrganizationalUnits.ps1 -InputCsvPath .\0009.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: ActiveDirectory
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    Column                Type      Required  Description
    --------------------  ----      --------  -----------
    Action                String    Yes       <fill in description>
    Notes                 String    Yes       <fill in description>
    Name                  String    Yes       <fill in description>
    ParentPath            String    Yes       <fill in description>
    Description           String    Yes       <fill in description>
    DisplayName           String    Yes       <fill in description>
    ManagedBy             String    Yes       <fill in description>
    StreetAddress         String    Yes       <fill in description>
    City                  String    Yes       <fill in description>
    StateOrProvince       String    Yes       <fill in description>
    PostalCode            String    Yes       <fill in description>
    CountryCode           String    Yes       <fill in description>
    CountryName           String    Yes       <fill in description>
    CountryNumericCode    String    Yes       <fill in description>
    SeeAlso               String    Yes       <fill in description>
    ProtectionEnabled     String    Yes       <fill in description>
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P0009-Create-ActiveDirectoryOrganizationalUnits_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Add-IfValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Hashtable,

        [Parameter(Mandatory)]
        [string]$Key,

        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        $Hashtable[$Key] = $text
    }
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'Name',
    'ParentPath',
    'Description',
    'DisplayName',
    'ManagedBy',
    'StreetAddress',
    'City',
    'StateOrProvince',
    'PostalCode',
    'CountryCode',
    'CountryName',
    'CountryNumericCode',
    'SeeAlso',
    'ProtectionEnabled'
)

Write-Status -Message 'Starting Active Directory organizational unit creation script.'
Ensure-ActiveDirectoryConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $name = Get-TrimmedValue -Value $row.Name
    $parentPath = Get-TrimmedValue -Value $row.ParentPath
    $primaryKey = "$name,$parentPath"

    try {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Name is required.'
        }

        if ([string]::IsNullOrWhiteSpace($parentPath)) {
            throw 'ParentPath (target parent OU distinguished name) is required.'
        }

        $escapedName = Escape-AdFilterValue -Value $name
        $existingOu = Invoke-WithRetry -OperationName "Lookup OU $name in path $parentPath" -ScriptBlock {
            Get-ADOrganizationalUnit -SearchBase $parentPath -SearchScope OneLevel -Filter "Name -eq '$escapedName'" -Properties DistinguishedName -ErrorAction SilentlyContinue |
                Select-Object -First 1
        }

        if ($existingOu) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryOrganizationalUnit' -Status 'Skipped' -Message "OU already exists as '$($existingOu.DistinguishedName)'."))
            $rowNumber++
            continue
        }

        $protectionEnabled = ConvertTo-Bool -Value $row.ProtectionEnabled -Default $true

        $newOuParams = @{
            Name                            = $name
            Path                            = $parentPath
            ProtectedFromAccidentalDeletion = $protectionEnabled
            ErrorAction                     = 'Stop'
        }

        Add-IfValue -Hashtable $newOuParams -Key 'Description' -Value $row.Description
        Add-IfValue -Hashtable $newOuParams -Key 'DisplayName' -Value $row.DisplayName
        Add-IfValue -Hashtable $newOuParams -Key 'ManagedBy' -Value $row.ManagedBy
        Add-IfValue -Hashtable $newOuParams -Key 'StreetAddress' -Value $row.StreetAddress
        Add-IfValue -Hashtable $newOuParams -Key 'City' -Value $row.City
        Add-IfValue -Hashtable $newOuParams -Key 'State' -Value $row.StateOrProvince
        Add-IfValue -Hashtable $newOuParams -Key 'PostalCode' -Value $row.PostalCode
        Add-IfValue -Hashtable $newOuParams -Key 'Country' -Value $row.CountryCode

        if ($PSCmdlet.ShouldProcess($primaryKey, 'Create Active Directory organizational unit')) {
            $createdOu = Invoke-WithRetry -OperationName "Create AD organizational unit $primaryKey" -ScriptBlock {
                New-ADOrganizationalUnit @newOuParams -PassThru
            }

            $replaceAttributes = @{}
            Add-IfValue -Hashtable $replaceAttributes -Key 'co' -Value $row.CountryName

            $seeAlsoValues = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.SeeAlso)
            if ($seeAlsoValues.Count -gt 0) {
                $replaceAttributes['seeAlso'] = $seeAlsoValues
            }

            $countryNumericCode = Get-TrimmedValue -Value $row.CountryNumericCode
            if (-not [string]::IsNullOrWhiteSpace($countryNumericCode)) {
                try {
                    $replaceAttributes['countryCode'] = [int]$countryNumericCode
                }
                catch {
                    throw "CountryNumericCode '$countryNumericCode' must be an integer value."
                }
            }

            if ($replaceAttributes.Count -gt 0) {
                Invoke-WithRetry -OperationName "Set AD organizational unit attributes $primaryKey" -ScriptBlock {
                    Set-ADOrganizationalUnit -Identity $createdOu.DistinguishedName -Replace $replaceAttributes -ErrorAction Stop
                }
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryOrganizationalUnit' -Status 'Created' -Message 'Organizational unit created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryOrganizationalUnit' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryOrganizationalUnit' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory organizational unit creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
