<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-120000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Provisions EntraUsers in Microsoft 365.

.DESCRIPTION
    Creates EntraUsers in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P3001-Create-EntraUsers.ps1 -InputCsvPath .\3001.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P3001-Create-EntraUsers.ps1 -InputCsvPath .\3001.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Users
    Required roles:   Global Administrator or appropriate workload-specific role
    Limitations:      None known.

    CSV Fields:
    Column                                Type      Required  Description
    ------------------------------------  ----      --------  -----------
    Action                                String    Yes       <fill in description>
    Notes                                 String    Yes       <fill in description>
    UserPrincipalName                     String    Yes       <fill in description>
    DisplayName                           String    Yes       <fill in description>
    GivenName                             String    Yes       <fill in description>
    Surname                               String    Yes       <fill in description>
    MailNickname                          String    Yes       <fill in description>
    UserType                              String    Yes       <fill in description>
    Password                              String    Yes       <fill in description>
    ForceChangePasswordNextSignIn         String    Yes       <fill in description>
    ForceChangePasswordNextSignInWithMfa  String    Yes       <fill in description>
    AccountEnabled                        String    Yes       <fill in description>
    UsageLocation                         String    Yes       <fill in description>
    PreferredLanguage                     String    Yes       <fill in description>
    Department                            String    Yes       <fill in description>
    JobTitle                              String    Yes       <fill in description>
    CompanyName                           String    Yes       <fill in description>
    OfficeLocation                        String    Yes       <fill in description>
    EmployeeId                            String    Yes       <fill in description>
    EmployeeType                          String    Yes       <fill in description>
    EmployeeHireDate                      String    Yes       <fill in description>
    MobilePhone                           String    Yes       <fill in description>
    BusinessPhones                        String    Yes       <fill in description>
    FaxNumber                             String    Yes       <fill in description>
    OtherMails                            String    Yes       <fill in description>
    StreetAddress                         String    Yes       <fill in description>
    City                                  String    Yes       <fill in description>
    State                                 String    Yes       <fill in description>
    PostalCode                            String    Yes       <fill in description>
    Country                               String    Yes       <fill in description>
    PasswordPolicies                      String    Yes       <fill in description>
    ExtensionAttribute1                   String    Yes       <fill in description>
    ExtensionAttribute2                   String    Yes       <fill in description>
    ExtensionAttribute3                   String    Yes       <fill in description>
    ExtensionAttribute4                   String    Yes       <fill in description>
    ExtensionAttribute5                   String    Yes       <fill in description>
    ExtensionAttribute6                   String    Yes       <fill in description>
    ExtensionAttribute7                   String    Yes       <fill in description>
    ExtensionAttribute8                   String    Yes       <fill in description>
    ExtensionAttribute9                   String    Yes       <fill in description>
    ExtensionAttribute10                  String    Yes       <fill in description>
    ExtensionAttribute11                  String    Yes       <fill in description>
    ExtensionAttribute12                  String    Yes       <fill in description>
    ExtensionAttribute13                  String    Yes       <fill in description>
    ExtensionAttribute14                  String    Yes       <fill in description>
    ExtensionAttribute15                  String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P3001-Create-EntraUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Add-BodyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Body,

        [Parameter(Mandatory)]
        [string]$PropertyName,

        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        $Body[$PropertyName] = $text
    }
}

function Add-BodyArray {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Body,

        [Parameter(Mandatory)]
        [string]$PropertyName,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    $items = ConvertTo-Array -Value $Value
    if ($items.Count -gt 0) {
        $Body[$PropertyName] = $items
    }
}

function Convert-ToIsoDateTimeOffsetString {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value,

        [Parameter(Mandatory)]
        [string]$FieldName
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    try {
        $parsed = [datetimeoffset]::Parse($text)
        return $parsed.ToString('o')
    }
    catch {
        throw "$FieldName value '$text' is invalid. Use an ISA-8601 compatible date/time value."
    }
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'UserPrincipalName',
    'DisplayName',
    'GivenName',
    'Surname',
    'MailNickname',
    'UserType',
    'Password',
    'ForceChangePasswordNextSignIn',
    'ForceChangePasswordNextSignInWithMfa',
    'AccountEnabled',
    'UsageLocation',
    'PreferredLanguage',
    'Department',
    'JobTitle',
    'CompanyName',
    'OfficeLocation',
    'EmployeeId',
    'EmployeeType',
    'EmployeeHireDate',
    'MobilePhone',
    'BusinessPhones',
    'FaxNumber',
    'OtherMails',
    'StreetAddress',
    'City',
    'State',
    'PostalCode',
    'Country',
    'PasswordPolicies',
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
    'ExtensionAttribute15'
)

Write-Status -Message 'Starting Entra ID user creation script (expanded field model).'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('User.ReadWrite.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $upn = Get-TrimmedValue -Value $row.UserPrincipalName

    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw 'UserPrincipalName is required.'
        }

        $escapedUpn = Escape-ODataString -Value $upn
        $existingUser = Invoke-WithRetry -OperationName "Lookup user $upn" -ScriptBlock {
            Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -ErrorAction Stop | Select-Object -First 1
        }

        if ($existingUser) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'CreateUser' -Status 'Skipped' -Message 'User already exists.'))
            $rowNumber++
            continue
        }

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            $displayName = ("{0} {1}" -f (Get-TrimmedValue -Value $row.GivenName), (Get-TrimmedValue -Value $row.Surname)).Trim()
        }

        if ([string]::IsNullOrWhiteSpace($displayName)) {
            throw 'DisplayName is required (or provide GivenName/Surname to derive one).'
        }

        $mailNickname = Get-TrimmedValue -Value $row.MailNickname
        if ([string]::IsNullOrWhiteSpace($mailNickname)) {
            $mailNickname = $upn.Split('@')[0]
        }

        if ([string]::IsNullOrWhiteSpace($mailNickname)) {
            throw 'MailNickname could not be derived. Provide MailNickname explicitly.'
        }

        $userType = Get-TrimmedValue -Value $row.UserType
        if ([string]::IsNullOrWhiteSpace($userType)) {
            $userType = 'Member'
        }

        if ($userType -notin @('Member', 'Guest')) {
            throw "UserType '$userType' is invalid. Use Member or Guest."
        }

        $password = [string]$row.Password
        if ([string]::IsNullOrWhiteSpace($password)) {
            throw 'Password is required for user creation.'
        }

        $body = @{
            accountEnabled    = ConvertTo-Bool -Value $row.AccountEnabled -Default $true
            displayName       = $displayName
            mailNickname      = $mailNickname
            userPrincipalName = $upn
            userType          = $userType
            passwordProfile   = @{
                password                             = $password
                forceChangePasswordNextSignIn       = ConvertTo-Bool -Value $row.ForceChangePasswordNextSignIn -Default $true
                forceChangePasswordNextSignInWithMfa = ConvertTo-Bool -Value $row.ForceChangePasswordNextSignInWithMfa -Default $false
            }
        }

        Add-BodyValue -Body $body -PropertyName 'givenName' -Value $row.GivenName
        Add-BodyValue -Body $body -PropertyName 'surname' -Value $row.Surname
        Add-BodyValue -Body $body -PropertyName 'usageLocation' -Value $row.UsageLocation
        Add-BodyValue -Body $body -PropertyName 'preferredLanguage' -Value $row.PreferredLanguage
        Add-BodyValue -Body $body -PropertyName 'department' -Value $row.Department
        Add-BodyValue -Body $body -PropertyName 'jobTitle' -Value $row.JobTitle
        Add-BodyValue -Body $body -PropertyName 'companyName' -Value $row.CompanyName
        Add-BodyValue -Body $body -PropertyName 'officeLocation' -Value $row.OfficeLocation
        Add-BodyValue -Body $body -PropertyName 'employeeId' -Value $row.EmployeeId
        Add-BodyValue -Body $body -PropertyName 'employeeType' -Value $row.EmployeeType
        Add-BodyValue -Body $body -PropertyName 'mobilePhone' -Value $row.MobilePhone
        Add-BodyValue -Body $body -PropertyName 'faxNumber' -Value $row.FaxNumber
        Add-BodyValue -Body $body -PropertyName 'streetAddress' -Value $row.StreetAddress
        Add-BodyValue -Body $body -PropertyName 'city' -Value $row.City
        Add-BodyValue -Body $body -PropertyName 'state' -Value $row.State
        Add-BodyValue -Body $body -PropertyName 'postalCode' -Value $row.PostalCode
        Add-BodyValue -Body $body -PropertyName 'country' -Value $row.Country
        Add-BodyValue -Body $body -PropertyName 'passwordPolicies' -Value $row.PasswordPolicies

        $employeeHireDate = Convert-ToIsoDateTimeOffsetString -Value $row.EmployeeHireDate -FieldName 'EmployeeHireDate'
        if (-not [string]::IsNullOrWhiteSpace($employeeHireDate)) {
            $body['employeeHireDate'] = $employeeHireDate
        }

        Add-BodyArray -Body $body -PropertyName 'businessPhones' -Value ([string]$row.BusinessPhones)
        Add-BodyArray -Body $body -PropertyName 'otherMails' -Value ([string]$row.OtherMails)

        $extensionAttributes = @{}
        for ($i = 1; $i -le 15; $i++) {
            $columnName = "ExtensionAttribute$i"
            $attributeName = "extensionAttribute$i"
            $value = Get-TrimmedValue -Value $row.$columnName
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $extensionAttributes[$attributeName] = $value
            }
        }

        if ($extensionAttributes.Count -gt 0) {
            $body['onPremisesExtensionAttributes'] = $extensionAttributes
        }

        if ($PSCmdlet.ShouldProcess($upn, 'Create Entra ID user')) {
            Invoke-WithRetry -OperationName "Create user $upn" -ScriptBlock {
                New-MgUser -BodyParameter $body -ErrorAction Stop | Out-Null
            }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'CreateUser' -Status 'Created' -Message 'User created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'CreateUser' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'CreateUser' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID user creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
