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
    Provisions ActiveDirectoryContacts in Active Directory.

.DESCRIPTION
    Creates ActiveDirectoryContacts in Active Directory based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P0002-Create-ActiveDirectoryContacts.ps1 -InputCsvPath .\0002.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P0002-Create-ActiveDirectoryContacts.ps1 -InputCsvPath .\0002.input.csv -WhatIf

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
    DisplayName           String    Yes       <fill in description>
    GivenName             String    Yes       <fill in description>
    Initials              String    Yes       <fill in description>
    Surname               String    Yes       <fill in description>
    Description           String    Yes       <fill in description>
    Company               String    Yes       <fill in description>
    Department            String    Yes       <fill in description>
    Division              String    Yes       <fill in description>
    Title                 String    Yes       <fill in description>
    Office                String    Yes       <fill in description>
    Manager               String    Yes       <fill in description>
    OfficePhone           String    Yes       <fill in description>
    MobilePhone           String    Yes       <fill in description>
    HomePhone             String    Yes       <fill in description>
    IpPhone               String    Yes       <fill in description>
    Fax                   String    Yes       <fill in description>
    Pager                 String    Yes       <fill in description>
    Mail                  String    Yes       <fill in description>
    MailNickname          String    Yes       <fill in description>
    ProxyAddresses        String    Yes       <fill in description>
    TargetAddress         String    Yes       <fill in description>
    HideFromAddressLists  String    Yes       <fill in description>
    StreetAddress         String    Yes       <fill in description>
    PostOfficeBox         String    Yes       <fill in description>
    City                  String    Yes       <fill in description>
    StateOrProvince       String    Yes       <fill in description>
    PostalCode            String    Yes       <fill in description>
    CountryCode           String    Yes       <fill in description>
    CountryName           String    Yes       <fill in description>
    CountryNumericCode    String    Yes       <fill in description>
    WebPage               String    Yes       <fill in description>
    ExtensionAttribute1   String    Yes       <fill in description>
    ExtensionAttribute2   String    Yes       <fill in description>
    ExtensionAttribute3   String    Yes       <fill in description>
    ExtensionAttribute4   String    Yes       <fill in description>
    ExtensionAttribute5   String    Yes       <fill in description>
    ExtensionAttribute6   String    Yes       <fill in description>
    ExtensionAttribute7   String    Yes       <fill in description>
    ExtensionAttribute8   String    Yes       <fill in description>
    ExtensionAttribute9   String    Yes       <fill in description>
    ExtensionAttribute10  String    Yes       <fill in description>
    ExtensionAttribute11  String    Yes       <fill in description>
    ExtensionAttribute12  String    Yes       <fill in description>
    ExtensionAttribute13  String    Yes       <fill in description>
    ExtensionAttribute14  String    Yes       <fill in description>
    ExtensionAttribute15  String    Yes       <fill in description>
    Path                  String    Yes       <fill in description>
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P0002-Create-ActiveDirectoryContacts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'DisplayName',
    'GivenName',
    'Initials',
    'Surname',
    'Description',
    'Company',
    'Department',
    'Division',
    'Title',
    'Office',
    'Manager',
    'OfficePhone',
    'MobilePhone',
    'HomePhone',
    'IpPhone',
    'Fax',
    'Pager',
    'Mail',
    'MailNickname',
    'ProxyAddresses',
    'TargetAddress',
    'HideFromAddressLists',
    'StreetAddress',
    'PostOfficeBox',
    'City',
    'StateOrProvince',
    'PostalCode',
    'CountryCode',
    'CountryName',
    'CountryNumericCode',
    'WebPage',
    'ExtensionAttribute1',
    'ExtensionAttribute2',
    'ExtensionAttribute3',
    'ExtensionAttribute4',
    'ExtensionAttribute5',
    'ExtensionAttribute6',
    'ExtensionAttribute7',
    'ExtensionAttribute8',
    'ExtensionAttribute9',
    'ExtensionAttribute10',
    'ExtensionAttribute11',
    'ExtensionAttribute12',
    'ExtensionAttribute13',
    'ExtensionAttribute14',
    'ExtensionAttribute15',
    'Path'
)

Write-Status -Message 'Starting Active Directory contact creation script.'
Ensure-ActiveDirectoryConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $name = Get-TrimmedValue -Value $row.Name
    $mail = Get-TrimmedValue -Value $row.Mail
    $path = Get-TrimmedValue -Value $row.Path
    $primaryKey = if (-not [string]::IsNullOrWhiteSpace($mail)) { $mail } else { $name }

    try {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Name is required.'
        }

        if ([string]::IsNullOrWhiteSpace($path)) {
            throw 'Path (target OU distinguished name) is required.'
        }

        $escapedName = Escape-AdFilterValue -Value $name
        $existingContact = Invoke-WithRetry -OperationName "Lookup contact $name in path $path" -ScriptBlock {
            Get-ADObject -SearchBase $path -SearchScope OneLevel -Filter "ObjectClass -eq 'contact' -and Name -eq '$escapedName'" -Properties DistinguishedName -ErrorAction SilentlyContinue |
                Select-Object -First 1
        }

        if ($existingContact) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryContact' -Status 'Skipped' -Message "Contact already exists as '$($existingContact.DistinguishedName)'."))
            $rowNumber++
            continue
        }

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            $displayName = $name
        }

        $otherAttributes = @{}
        Add-IfValue -Hashtable $otherAttributes -Key 'displayName' -Value $displayName
        Add-IfValue -Hashtable $otherAttributes -Key 'givenName' -Value $row.GivenName
        Add-IfValue -Hashtable $otherAttributes -Key 'initials' -Value $row.Initials
        Add-IfValue -Hashtable $otherAttributes -Key 'sn' -Value $row.Surname
        Add-IfValue -Hashtable $otherAttributes -Key 'description' -Value $row.Description
        Add-IfValue -Hashtable $otherAttributes -Key 'company' -Value $row.Company
        Add-IfValue -Hashtable $otherAttributes -Key 'department' -Value $row.Department
        Add-IfValue -Hashtable $otherAttributes -Key 'division' -Value $row.Division
        Add-IfValue -Hashtable $otherAttributes -Key 'title' -Value $row.Title
        Add-IfValue -Hashtable $otherAttributes -Key 'physicalDeliveryOfficeName' -Value $row.Office
        Add-IfValue -Hashtable $otherAttributes -Key 'manager' -Value $row.Manager
        Add-IfValue -Hashtable $otherAttributes -Key 'telephoneNumber' -Value $row.OfficePhone
        Add-IfValue -Hashtable $otherAttributes -Key 'mobile' -Value $row.MobilePhone
        Add-IfValue -Hashtable $otherAttributes -Key 'homePhone' -Value $row.HomePhone
        Add-IfValue -Hashtable $otherAttributes -Key 'ipPhone' -Value $row.IpPhone
        Add-IfValue -Hashtable $otherAttributes -Key 'facsimileTelephoneNumber' -Value $row.Fax
        Add-IfValue -Hashtable $otherAttributes -Key 'pager' -Value $row.Pager
        Add-IfValue -Hashtable $otherAttributes -Key 'mail' -Value $row.Mail
        Add-IfValue -Hashtable $otherAttributes -Key 'mailNickname' -Value $row.MailNickname
        Add-IfValue -Hashtable $otherAttributes -Key 'targetAddress' -Value $row.TargetAddress
        Add-IfValue -Hashtable $otherAttributes -Key 'streetAddress' -Value $row.StreetAddress
        Add-IfValue -Hashtable $otherAttributes -Key 'postOfficeBox' -Value $row.PostOfficeBox
        Add-IfValue -Hashtable $otherAttributes -Key 'l' -Value $row.City
        Add-IfValue -Hashtable $otherAttributes -Key 'st' -Value $row.StateOrProvince
        Add-IfValue -Hashtable $otherAttributes -Key 'postalCode' -Value $row.PostalCode
        Add-IfValue -Hashtable $otherAttributes -Key 'c' -Value $row.CountryCode
        Add-IfValue -Hashtable $otherAttributes -Key 'co' -Value $row.CountryName
        Add-IfValue -Hashtable $otherAttributes -Key 'wWWHomePage' -Value $row.WebPage

        $proxyAddresses = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ProxyAddresses)
        if ($proxyAddresses.Count -gt 0) {
            $otherAttributes['proxyAddresses'] = [string[]]$proxyAddresses
        }

        $hideFromAddressLists = Get-NullableBool -Value $row.HideFromAddressLists
        if ($null -ne $hideFromAddressLists) {
            $otherAttributes['msExchHideFromAddressLists'] = $hideFromAddressLists
        }

        $countryNumericCode = Get-TrimmedValue -Value $row.CountryNumericCode
        if (-not [string]::IsNullOrWhiteSpace($countryNumericCode)) {
            try {
                $otherAttributes['countryCode'] = [int]$countryNumericCode
            }
            catch {
                throw "CountryNumericCode '$countryNumericCode' must be an integer value."
            }
        }

        for ($i = 1; $i -le 15; $i++) {
            $columnName = "ExtensionAttribute$i"
            $attributeName = "extensionAttribute$i"
            Add-IfValue -Hashtable $otherAttributes -Key $attributeName -Value $row.$columnName
        }

        $newContactParams = @{
            Name       = $name
            Type       = 'contact'
            Path       = $path
            ErrorAction = 'Stop'
        }

        if ($otherAttributes.Count -gt 0) {
            $newContactParams['OtherAttributes'] = $otherAttributes
        }

        if ($PSCmdlet.ShouldProcess($primaryKey, 'Create Active Directory contact')) {
            Invoke-WithRetry -OperationName "Create AD contact $primaryKey" -ScriptBlock {
                New-ADObject @newContactParams
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryContact' -Status 'Created' -Message 'Contact created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryContact' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryContact' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory contact creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
