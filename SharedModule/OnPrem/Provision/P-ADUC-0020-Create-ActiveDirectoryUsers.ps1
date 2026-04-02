<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260302-190500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
ActiveDirectory

.MODULEVERSIONPOLICY
Best-effort PSGallery check (offline-safe)

.SYNOPSIS
    Provisions ActiveDirectoryUsers in Active Directory.

.DESCRIPTION
    Creates ActiveDirectoryUsers in Active Directory based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P0001-Create-ActiveDirectoryUsers.ps1 -InputCsvPath .\0001.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P0001-Create-ActiveDirectoryUsers.ps1 -InputCsvPath .\0001.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: ActiveDirectory
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    Column                  Type      Required  Description
    ----------------------  ----      --------  -----------
    Action                  String    Yes       <fill in description>
    Notes                   String    Yes       <fill in description>
    SamAccountName          String    Yes       <fill in description>
    UserPrincipalName       String    Yes       <fill in description>
    GivenName               String    Yes       <fill in description>
    Initials                String    Yes       <fill in description>
    Surname                 String    Yes       <fill in description>
    DisplayName             String    Yes       <fill in description>
    Name                    String    Yes       <fill in description>
    Description             String    Yes       <fill in description>
    EmployeeID              String    Yes       <fill in description>
    EmployeeNumber          String    Yes       <fill in description>
    EmployeeType            String    Yes       <fill in description>
    Enabled                 String    Yes       <fill in description>
    AccountPassword         String    Yes       <fill in description>
    ChangePasswordAtLogon   String    Yes       <fill in description>
    PasswordNeverExpires    String    Yes       <fill in description>
    CannotChangePassword    String    Yes       <fill in description>
    SmartcardLogonRequired  String    Yes       <fill in description>
    AccountExpirationDate   String    Yes       <fill in description>
    UserWorkstations        String    Yes       <fill in description>
    ScriptPath              String    Yes       <fill in description>
    ProfilePath             String    Yes       <fill in description>
    HomeDirectory           String    Yes       <fill in description>
    HomeDrive               String    Yes       <fill in description>
    Title                   String    Yes       <fill in description>
    Department              String    Yes       <fill in description>
    Company                 String    Yes       <fill in description>
    Division                String    Yes       <fill in description>
    Office                  String    Yes       <fill in description>
    Manager                 String    Yes       <fill in description>
    OfficePhone             String    Yes       <fill in description>
    MobilePhone             String    Yes       <fill in description>
    Mail                    String    Yes       <fill in description>
    MailNickname            String    Yes       <fill in description>
    ProxyAddresses          String    Yes       <fill in description>
    TargetAddress           String    Yes       <fill in description>
    HideFromAddressLists    String    Yes       <fill in description>
    StreetAddress           String    Yes       <fill in description>
    PostOfficeBox           String    Yes       <fill in description>
    City                    String    Yes       <fill in description>
    StateOrProvince         String    Yes       <fill in description>
    PostalCode              String    Yes       <fill in description>
    CountryCode             String    Yes       <fill in description>
    CountryName             String    Yes       <fill in description>
    CountryNumericCode      String    Yes       <fill in description>
    HomePhone               String    Yes       <fill in description>
    IpPhone                 String    Yes       <fill in description>
    Fax                     String    Yes       <fill in description>
    Pager                   String    Yes       <fill in description>
    WebPage                 String    Yes       <fill in description>
    ExtensionAttribute1     String    Yes       <fill in description>
    ExtensionAttribute2     String    Yes       <fill in description>
    ExtensionAttribute3     String    Yes       <fill in description>
    ExtensionAttribute4     String    Yes       <fill in description>
    ExtensionAttribute5     String    Yes       <fill in description>
    ExtensionAttribute6     String    Yes       <fill in description>
    ExtensionAttribute7     String    Yes       <fill in description>
    ExtensionAttribute8     String    Yes       <fill in description>
    ExtensionAttribute9     String    Yes       <fill in description>
    ExtensionAttribute10    String    Yes       <fill in description>
    ExtensionAttribute11    String    Yes       <fill in description>
    ExtensionAttribute12    String    Yes       <fill in description>
    ExtensionAttribute13    String    Yes       <fill in description>
    ExtensionAttribute14    String    Yes       <fill in description>
    ExtensionAttribute15    String    Yes       <fill in description>
    Path                    String    Yes       <fill in description>
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P0001-Create-ActiveDirectoryUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Add-IfNullableBool {
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

    $boolValue = Get-NullableBool -Value $Value
    if ($null -ne $boolValue) {
        $Hashtable[$Key] = $boolValue
    }
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'SamAccountName',
    'UserPrincipalName',
    'GivenName',
    'Initials',
    'Surname',
    'DisplayName',
    'Name',
    'Description',
    'EmployeeID',
    'EmployeeNumber',
    'EmployeeType',
    'Enabled',
    'AccountPassword',
    'ChangePasswordAtLogon',
    'PasswordNeverExpires',
    'CannotChangePassword',
    'SmartcardLogonRequired',
    'AccountExpirationDate',
    'UserWorkstations',
    'ScriptPath',
    'ProfilePath',
    'HomeDirectory',
    'HomeDrive',
    'Title',
    'Department',
    'Company',
    'Division',
    'Office',
    'Manager',
    'OfficePhone',
    'MobilePhone',
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
    'HomePhone',
    'IpPhone',
    'Fax',
    'Pager',
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

Write-Status -Message 'Starting Active Directory user creation script.'
Ensure-ActiveDirectoryConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $samAccountName = Get-TrimmedValue -Value $row.SamAccountName
    $userPrincipalName = Get-TrimmedValue -Value $row.UserPrincipalName
    $primaryKey = if (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) { $userPrincipalName } else { $samAccountName }

    try {
        if ([string]::IsNullOrWhiteSpace($samAccountName)) {
            throw 'SamAccountName is required.'
        }

        if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
            throw 'UserPrincipalName is required.'
        }

        $givenName = Get-TrimmedValue -Value $row.GivenName
        $surname = Get-TrimmedValue -Value $row.Surname
        if ([string]::IsNullOrWhiteSpace($givenName) -or [string]::IsNullOrWhiteSpace($surname)) {
            throw 'GivenName and Surname are required.'
        }

        $path = Get-TrimmedValue -Value $row.Path
        if ([string]::IsNullOrWhiteSpace($path)) {
            throw 'Path (target OU distinguished name) is required.'
        }

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            $displayName = "$givenName $surname".Trim()
        }

        $name = Get-TrimmedValue -Value $row.Name
        if ([string]::IsNullOrWhiteSpace($name)) {
            $name = $displayName
        }

        $enabled = ConvertTo-Bool -Value $row.Enabled -Default $true
        $passwordText = Get-TrimmedValue -Value $row.AccountPassword
        if ($enabled -and [string]::IsNullOrWhiteSpace($passwordText)) {
            throw 'AccountPassword is required when Enabled is TRUE.'
        }

        $existingUser = $null
        $escapedSam = Escape-AdFilterValue -Value $samAccountName
        $existingBySam = Invoke-WithRetry -OperationName "Lookup user by SamAccountName $samAccountName" -ScriptBlock {
            Get-ADUser -Filter "SamAccountName -eq '$escapedSam'" -ErrorAction SilentlyContinue | Select-Object -First 1
        }

        if ($existingBySam) {
            $existingUser = $existingBySam
        }
        else {
            $escapedUpn = Escape-AdFilterValue -Value $userPrincipalName
            $existingByUpn = Invoke-WithRetry -OperationName "Lookup user by UPN $userPrincipalName" -ScriptBlock {
                Get-ADUser -Filter "UserPrincipalName -eq '$escapedUpn'" -ErrorAction SilentlyContinue | Select-Object -First 1
            }

            if ($existingByUpn) {
                $existingUser = $existingByUpn
            }
        }

        if ($existingUser) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryUser' -Status 'Skipped' -Message "User already exists as '$($existingUser.DistinguishedName)'."))
            $rowNumber++
            continue
        }

        $newUserParams = @{
            Name              = $name
            SamAccountName    = $samAccountName
            UserPrincipalName = $userPrincipalName
            GivenName         = $givenName
            Surname           = $surname
            DisplayName       = $displayName
            Enabled           = $enabled
            Path              = $path
        }

        Add-IfValue -Hashtable $newUserParams -Key 'Initials' -Value $row.Initials
        Add-IfValue -Hashtable $newUserParams -Key 'Description' -Value $row.Description
        Add-IfValue -Hashtable $newUserParams -Key 'EmployeeID' -Value $row.EmployeeID
        Add-IfValue -Hashtable $newUserParams -Key 'EmployeeNumber' -Value $row.EmployeeNumber
        Add-IfNullableBool -Hashtable $newUserParams -Key 'ChangePasswordAtLogon' -Value $row.ChangePasswordAtLogon
        Add-IfNullableBool -Hashtable $newUserParams -Key 'PasswordNeverExpires' -Value $row.PasswordNeverExpires
        Add-IfNullableBool -Hashtable $newUserParams -Key 'CannotChangePassword' -Value $row.CannotChangePassword
        Add-IfNullableBool -Hashtable $newUserParams -Key 'SmartcardLogonRequired' -Value $row.SmartcardLogonRequired

        $accountExpirationDate = ConvertTo-NullableDateTime -Value $row.AccountExpirationDate
        if ($null -ne $accountExpirationDate) {
            $newUserParams['AccountExpirationDate'] = $accountExpirationDate
        }

        Add-IfValue -Hashtable $newUserParams -Key 'LogonWorkstations' -Value $row.UserWorkstations
        Add-IfValue -Hashtable $newUserParams -Key 'ScriptPath' -Value $row.ScriptPath
        Add-IfValue -Hashtable $newUserParams -Key 'ProfilePath' -Value $row.ProfilePath
        Add-IfValue -Hashtable $newUserParams -Key 'HomeDirectory' -Value $row.HomeDirectory
        Add-IfValue -Hashtable $newUserParams -Key 'HomeDrive' -Value $row.HomeDrive
        Add-IfValue -Hashtable $newUserParams -Key 'Title' -Value $row.Title
        Add-IfValue -Hashtable $newUserParams -Key 'Department' -Value $row.Department
        Add-IfValue -Hashtable $newUserParams -Key 'Company' -Value $row.Company
        Add-IfValue -Hashtable $newUserParams -Key 'Division' -Value $row.Division
        Add-IfValue -Hashtable $newUserParams -Key 'Office' -Value $row.Office
        Add-IfValue -Hashtable $newUserParams -Key 'Manager' -Value $row.Manager
        Add-IfValue -Hashtable $newUserParams -Key 'OfficePhone' -Value $row.OfficePhone
        Add-IfValue -Hashtable $newUserParams -Key 'MobilePhone' -Value $row.MobilePhone
        Add-IfValue -Hashtable $newUserParams -Key 'HomePhone' -Value $row.HomePhone
        Add-IfValue -Hashtable $newUserParams -Key 'EmailAddress' -Value $row.Mail
        Add-IfValue -Hashtable $newUserParams -Key 'StreetAddress' -Value $row.StreetAddress
        Add-IfValue -Hashtable $newUserParams -Key 'POBox' -Value $row.PostOfficeBox
        Add-IfValue -Hashtable $newUserParams -Key 'City' -Value $row.City
        Add-IfValue -Hashtable $newUserParams -Key 'State' -Value $row.StateOrProvince
        Add-IfValue -Hashtable $newUserParams -Key 'PostalCode' -Value $row.PostalCode
        Add-IfValue -Hashtable $newUserParams -Key 'Fax' -Value $row.Fax
        Add-IfValue -Hashtable $newUserParams -Key 'Pager' -Value $row.Pager
        Add-IfValue -Hashtable $newUserParams -Key 'HomePage' -Value $row.WebPage

        if (-not [string]::IsNullOrWhiteSpace($passwordText)) {
            $newUserParams['AccountPassword'] = ConvertTo-SecureString -String $passwordText -AsPlainText -Force
        }

        $otherAttributes = @{}

        Add-IfValue -Hashtable $otherAttributes -Key 'employeeType' -Value $row.EmployeeType
        Add-IfValue -Hashtable $otherAttributes -Key 'mailNickname' -Value $row.MailNickname
        Add-IfValue -Hashtable $otherAttributes -Key 'targetAddress' -Value $row.TargetAddress
        Add-IfValue -Hashtable $otherAttributes -Key 'ipPhone' -Value $row.IpPhone

        $proxyAddresses = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ProxyAddresses)
        if ($proxyAddresses.Count -gt 0) {
            $otherAttributes['proxyAddresses'] = [string[]]$proxyAddresses
        }

        $hideFromAddressLists = Get-NullableBool -Value $row.HideFromAddressLists
        if ($null -ne $hideFromAddressLists) {
            $otherAttributes['msExchHideFromAddressLists'] = $hideFromAddressLists
        }

        Add-IfValue -Hashtable $otherAttributes -Key 'c' -Value $row.CountryCode
        Add-IfValue -Hashtable $otherAttributes -Key 'co' -Value $row.CountryName
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

        if ($otherAttributes.Count -gt 0) {
            $newUserParams['OtherAttributes'] = $otherAttributes
        }

        if ($PSCmdlet.ShouldProcess($primaryKey, 'Create Active Directory user')) {
            Invoke-WithRetry -OperationName "Create AD user $primaryKey" -ScriptBlock {
                New-ADUser @newUserParams -ErrorAction Stop
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryUser' -Status 'Created' -Message 'User created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryUser' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryUser' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory user creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
