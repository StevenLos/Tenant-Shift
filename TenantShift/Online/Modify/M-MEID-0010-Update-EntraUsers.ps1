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
    Modifies EntraUsers in Microsoft 365.

.DESCRIPTION
    Updates EntraUsers in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M3001-Update-EntraUsers.ps1 -InputCsvPath .\3001.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M3001-Update-EntraUsers.ps1 -InputCsvPath .\3001.input.csv -WhatIf

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
    ClearAttributes                       String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3001-Update-EntraUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

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
    'ExtensionAttribute15',
    'ClearAttributes'
)

$clearFieldMap = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$clearFieldMap['GivenName'] = 'givenName'
$clearFieldMap['Surname'] = 'surname'
$clearFieldMap['UsageLocation'] = 'usageLocation'
$clearFieldMap['PreferredLanguage'] = 'preferredLanguage'
$clearFieldMap['Department'] = 'department'
$clearFieldMap['JobTitle'] = 'jobTitle'
$clearFieldMap['CompanyName'] = 'companyName'
$clearFieldMap['OfficeLocation'] = 'officeLocation'
$clearFieldMap['EmployeeId'] = 'employeeId'
$clearFieldMap['EmployeeType'] = 'employeeType'
$clearFieldMap['EmployeeHireDate'] = 'employeeHireDate'
$clearFieldMap['MobilePhone'] = 'mobilePhone'
$clearFieldMap['BusinessPhones'] = 'businessPhones'
$clearFieldMap['FaxNumber'] = 'faxNumber'
$clearFieldMap['OtherMails'] = 'otherMails'
$clearFieldMap['StreetAddress'] = 'streetAddress'
$clearFieldMap['City'] = 'city'
$clearFieldMap['State'] = 'state'
$clearFieldMap['PostalCode'] = 'postalCode'
$clearFieldMap['Country'] = 'country'
$clearFieldMap['PasswordPolicies'] = 'passwordPolicies'

Write-Status -Message 'Starting Entra ID user update script (expanded field model).'
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
        $users = @(Invoke-WithRetry -OperationName "Lookup user $upn" -ScriptBlock {
            Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -Property 'id,userPrincipalName,displayName,accountEnabled' -ErrorAction Stop
        })

        if ($users.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'UpdateUser' -Status 'NotFound' -Message 'User not found.'))
            $rowNumber++
            continue
        }

        if ($users.Count -gt 1) {
            throw "Multiple users were returned for UPN '$upn'. Resolve duplicate directory objects before retrying."
        }

        $user = $users[0]
        $body = @{}
        $setFields = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        $scalarMappings = @(
            @{ Column = 'DisplayName'; Property = 'displayName' },
            @{ Column = 'GivenName'; Property = 'givenName' },
            @{ Column = 'Surname'; Property = 'surname' },
            @{ Column = 'MailNickname'; Property = 'mailNickname' },
            @{ Column = 'UserType'; Property = 'userType' },
            @{ Column = 'UsageLocation'; Property = 'usageLocation' },
            @{ Column = 'PreferredLanguage'; Property = 'preferredLanguage' },
            @{ Column = 'Department'; Property = 'department' },
            @{ Column = 'JobTitle'; Property = 'jobTitle' },
            @{ Column = 'CompanyName'; Property = 'companyName' },
            @{ Column = 'OfficeLocation'; Property = 'officeLocation' },
            @{ Column = 'EmployeeId'; Property = 'employeeId' },
            @{ Column = 'EmployeeType'; Property = 'employeeType' },
            @{ Column = 'MobilePhone'; Property = 'mobilePhone' },
            @{ Column = 'FaxNumber'; Property = 'faxNumber' },
            @{ Column = 'StreetAddress'; Property = 'streetAddress' },
            @{ Column = 'City'; Property = 'city' },
            @{ Column = 'State'; Property = 'state' },
            @{ Column = 'PostalCode'; Property = 'postalCode' },
            @{ Column = 'Country'; Property = 'country' },
            @{ Column = 'PasswordPolicies'; Property = 'passwordPolicies' }
        )

        foreach ($mapping in $scalarMappings) {
            $value = Get-TrimmedValue -Value $row.($mapping.Column)
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                if ($mapping.Property -eq 'userType' -and $value -notin @('Member', 'Guest')) {
                    throw "UserType '$value' is invalid. Use Member or Guest."
                }

                $body[$mapping.Property] = $value
                $null = $setFields.Add($mapping.Column)
            }
        }

        $accountEnabledRaw = Get-TrimmedValue -Value $row.AccountEnabled
        if (-not [string]::IsNullOrWhiteSpace($accountEnabledRaw)) {
            $body['accountEnabled'] = ConvertTo-Bool -Value $accountEnabledRaw
            $null = $setFields.Add('AccountEnabled')
        }

        $businessPhonesRaw = Get-TrimmedValue -Value $row.BusinessPhones
        if (-not [string]::IsNullOrWhiteSpace($businessPhonesRaw)) {
            $body['businessPhones'] = ConvertTo-Array -Value $businessPhonesRaw
            $null = $setFields.Add('BusinessPhones')
        }

        $otherMailsRaw = Get-TrimmedValue -Value $row.OtherMails
        if (-not [string]::IsNullOrWhiteSpace($otherMailsRaw)) {
            $body['otherMails'] = ConvertTo-Array -Value $otherMailsRaw
            $null = $setFields.Add('OtherMails')
        }

        $employeeHireDate = Convert-ToIsoDateTimeOffsetString -Value $row.EmployeeHireDate -FieldName 'EmployeeHireDate'
        if (-not [string]::IsNullOrWhiteSpace($employeeHireDate)) {
            $body['employeeHireDate'] = $employeeHireDate
            $null = $setFields.Add('EmployeeHireDate')
        }

        $password = [string]$row.Password
        if (-not [string]::IsNullOrWhiteSpace($password)) {
            $passwordProfile = @{ password = $password }

            $forceChange = Get-TrimmedValue -Value $row.ForceChangePasswordNextSignIn
            if (-not [string]::IsNullOrWhiteSpace($forceChange)) {
                $passwordProfile['forceChangePasswordNextSignIn'] = ConvertTo-Bool -Value $forceChange
            }

            $forceChangeWithMfa = Get-TrimmedValue -Value $row.ForceChangePasswordNextSignInWithMfa
            if (-not [string]::IsNullOrWhiteSpace($forceChangeWithMfa)) {
                $passwordProfile['forceChangePasswordNextSignInWithMfa'] = ConvertTo-Bool -Value $forceChangeWithMfa
            }

            $body['passwordProfile'] = $passwordProfile
            $null = $setFields.Add('Password')
        }

        $extensionAttributes = @{}
        $extensionTouched = $false
        for ($i = 1; $i -le 15; $i++) {
            $columnName = "ExtensionAttribute$i"
            $attributeName = "extensionAttribute$i"
            $value = Get-TrimmedValue -Value $row.$columnName
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $extensionAttributes[$attributeName] = $value
                $extensionTouched = $true
                $null = $setFields.Add($columnName)
            }
        }

        $clearRequested = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ClearAttributes)
        foreach ($clearName in $clearRequested) {
            if ($setFields.Contains($clearName)) {
                throw "Field '$clearName' is set and cleared in the same row. Use only one behavior per field."
            }

            if ($clearFieldMap.ContainsKey($clearName)) {
                $propertyName = $clearFieldMap[$clearName]
                if ($propertyName -in @('businessPhones', 'otherMails')) {
                    $body[$propertyName] = @()
                }
                else {
                    $body[$propertyName] = $null
                }

                continue
            }

            if ($clearName -match '^ExtensionAttribute([1-9]|1[0-5])$') {
                $attributeName = 'extensionAttribute{0}' -f $Matches[1]
                $extensionAttributes[$attributeName] = $null
                $extensionTouched = $true
                continue
            }

            throw "ClearAttributes value '$clearName' is not supported."
        }

        if ($extensionTouched) {
            $body['onPremisesExtensionAttributes'] = $extensionAttributes
        }

        if ($body.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'UpdateUser' -Status 'Skipped' -Message 'No updates were requested.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($upn, 'Update Entra ID user attributes')) {
            Invoke-WithRetry -OperationName "Update user $upn" -ScriptBlock {
                Update-MgUser -UserId $user.Id -BodyParameter $body -ErrorAction Stop | Out-Null
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'UpdateUser' -Status 'Updated' -Message 'User updated successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'UpdateUser' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'UpdateUser' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID user update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
