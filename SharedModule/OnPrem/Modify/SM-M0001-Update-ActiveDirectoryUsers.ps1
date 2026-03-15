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
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M0001-Update-ActiveDirectoryUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Resolve-TargetAdUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IdentityType,

        [Parameter(Mandatory)]
        [string]$IdentityValue
    )

    $normalizedType = $IdentityType.Trim().ToLowerInvariant()
    switch ($normalizedType) {
        'samaccountname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADUser -Filter "SamAccountName -eq '$escaped'" -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'userprincipalname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADUser -Filter "UserPrincipalName -eq '$escaped'" -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'distinguishedname' {
            return Get-ADUser -Identity $IdentityValue -ErrorAction SilentlyContinue
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            return Get-ADUser -Identity $guid -ErrorAction SilentlyContinue
        }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use SamAccountName, UserPrincipalName, DistinguishedName, or ObjectGuid."
        }
    }
}

function Add-SetUserField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$SetParams,

        [Parameter(Mandatory)]
        [string]$ParameterName,

        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        $SetParams[$ParameterName] = $text
        return $true
    }

    return $false
}

function Add-SetUserBoolField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$SetParams,

        [Parameter(Mandatory)]
        [string]$ParameterName,

        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $boolValue = Get-NullableBool -Value $Value
    if ($null -ne $boolValue) {
        $SetParams[$ParameterName] = $boolValue
        return $true
    }

    return $false
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'IdentityType',
    'IdentityValue',
    'ClearAttributes',
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

$clearAttributeMap = @{
    SamAccountName = 'sAMAccountName'
    UserPrincipalName = 'userPrincipalName'
    GivenName = 'givenName'
    Initials = 'initials'
    Surname = 'sn'
    DisplayName = 'displayName'
    Description = 'description'
    EmployeeID = 'employeeID'
    EmployeeNumber = 'employeeNumber'
    EmployeeType = 'employeeType'
    UserWorkstations = 'userWorkstations'
    ScriptPath = 'scriptPath'
    ProfilePath = 'profilePath'
    HomeDirectory = 'homeDirectory'
    HomeDrive = 'homeDrive'
    Title = 'title'
    Department = 'department'
    Company = 'company'
    Division = 'division'
    Office = 'physicalDeliveryOfficeName'
    Manager = 'manager'
    OfficePhone = 'telephoneNumber'
    MobilePhone = 'mobile'
    Mail = 'mail'
    MailNickname = 'mailNickname'
    ProxyAddresses = 'proxyAddresses'
    TargetAddress = 'targetAddress'
    HideFromAddressLists = 'msExchHideFromAddressLists'
    StreetAddress = 'streetAddress'
    PostOfficeBox = 'postOfficeBox'
    City = 'l'
    StateOrProvince = 'st'
    PostalCode = 'postalCode'
    CountryCode = 'c'
    CountryName = 'co'
    CountryNumericCode = 'countryCode'
    HomePhone = 'homePhone'
    IpPhone = 'ipPhone'
    Fax = 'facsimileTelephoneNumber'
    Pager = 'pager'
    WebPage = 'wWWHomePage'
    AccountExpirationDate = 'accountExpires'
}

for ($i = 1; $i -le 15; $i++) {
    $clearAttributeMap["ExtensionAttribute$i"] = "extensionAttribute$i"
}

Write-Status -Message 'Starting Active Directory user update script.'
Ensure-ActiveDirectoryConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $identityType = Get-TrimmedValue -Value $row.IdentityType
    $identityValue = Get-TrimmedValue -Value $row.IdentityValue
    $primaryKey = "${identityType}:$identityValue"

    try {
        if ([string]::IsNullOrWhiteSpace($identityType) -or [string]::IsNullOrWhiteSpace($identityValue)) {
            throw 'IdentityType and IdentityValue are required.'
        }

        $targetUser = Invoke-WithRetry -OperationName "Resolve AD user $primaryKey" -ScriptBlock {
            Resolve-TargetAdUser -IdentityType $identityType -IdentityValue $identityValue
        }

        if (-not $targetUser) {
            throw 'Target user was not found.'
        }

        $resolvedKey = if (-not [string]::IsNullOrWhiteSpace($targetUser.UserPrincipalName)) { $targetUser.UserPrincipalName } else { $targetUser.SamAccountName }
        $messages = [System.Collections.Generic.List[string]]::new()
        $changeCount = 0

        $setParams = @{
            Identity = $targetUser.DistinguishedName
        }

        if (Add-SetUserField -SetParams $setParams -ParameterName 'SamAccountName' -Value $row.SamAccountName) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'UserPrincipalName' -Value $row.UserPrincipalName) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'GivenName' -Value $row.GivenName) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'Initials' -Value $row.Initials) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'Surname' -Value $row.Surname) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'DisplayName' -Value $row.DisplayName) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'Description' -Value $row.Description) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'EmployeeID' -Value $row.EmployeeID) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'EmployeeNumber' -Value $row.EmployeeNumber) { $changeCount++ }
        if (Add-SetUserBoolField -SetParams $setParams -ParameterName 'ChangePasswordAtLogon' -Value $row.ChangePasswordAtLogon) { $changeCount++ }
        if (Add-SetUserBoolField -SetParams $setParams -ParameterName 'PasswordNeverExpires' -Value $row.PasswordNeverExpires) { $changeCount++ }
        if (Add-SetUserBoolField -SetParams $setParams -ParameterName 'CannotChangePassword' -Value $row.CannotChangePassword) { $changeCount++ }
        if (Add-SetUserBoolField -SetParams $setParams -ParameterName 'SmartcardLogonRequired' -Value $row.SmartcardLogonRequired) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'LogonWorkstations' -Value $row.UserWorkstations) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'ScriptPath' -Value $row.ScriptPath) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'ProfilePath' -Value $row.ProfilePath) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'HomeDirectory' -Value $row.HomeDirectory) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'HomeDrive' -Value $row.HomeDrive) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'Title' -Value $row.Title) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'Department' -Value $row.Department) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'Company' -Value $row.Company) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'Division' -Value $row.Division) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'Office' -Value $row.Office) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'Manager' -Value $row.Manager) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'OfficePhone' -Value $row.OfficePhone) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'MobilePhone' -Value $row.MobilePhone) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'HomePhone' -Value $row.HomePhone) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'EmailAddress' -Value $row.Mail) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'StreetAddress' -Value $row.StreetAddress) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'POBox' -Value $row.PostOfficeBox) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'City' -Value $row.City) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'State' -Value $row.StateOrProvince) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'PostalCode' -Value $row.PostalCode) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'Fax' -Value $row.Fax) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'Pager' -Value $row.Pager) { $changeCount++ }
        if (Add-SetUserField -SetParams $setParams -ParameterName 'HomePage' -Value $row.WebPage) { $changeCount++ }

        $replaceAttributes = @{}

        $mailNickname = Get-TrimmedValue -Value $row.MailNickname
        if (-not [string]::IsNullOrWhiteSpace($mailNickname)) {
            $replaceAttributes['mailNickname'] = $mailNickname
            $changeCount++
        }

        $targetAddress = Get-TrimmedValue -Value $row.TargetAddress
        if (-not [string]::IsNullOrWhiteSpace($targetAddress)) {
            $replaceAttributes['targetAddress'] = $targetAddress
            $changeCount++
        }

        $ipPhone = Get-TrimmedValue -Value $row.IpPhone
        if (-not [string]::IsNullOrWhiteSpace($ipPhone)) {
            $replaceAttributes['ipPhone'] = $ipPhone
            $changeCount++
        }

        $employeeType = Get-TrimmedValue -Value $row.EmployeeType
        if (-not [string]::IsNullOrWhiteSpace($employeeType)) {
            $replaceAttributes['employeeType'] = $employeeType
            $changeCount++
        }

        $proxyAddresses = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ProxyAddresses)
        if ($proxyAddresses.Count -gt 0) {
            $replaceAttributes['proxyAddresses'] = [string[]]$proxyAddresses
            $changeCount++
        }

        $hideFromAddressLists = Get-NullableBool -Value $row.HideFromAddressLists
        if ($null -ne $hideFromAddressLists) {
            $replaceAttributes['msExchHideFromAddressLists'] = $hideFromAddressLists
            $changeCount++
        }

        $countryCode = Get-TrimmedValue -Value $row.CountryCode
        if (-not [string]::IsNullOrWhiteSpace($countryCode)) {
            $replaceAttributes['c'] = $countryCode
            $changeCount++
        }

        $countryName = Get-TrimmedValue -Value $row.CountryName
        if (-not [string]::IsNullOrWhiteSpace($countryName)) {
            $replaceAttributes['co'] = $countryName
            $changeCount++
        }

        $countryNumericCode = Get-TrimmedValue -Value $row.CountryNumericCode
        if (-not [string]::IsNullOrWhiteSpace($countryNumericCode)) {
            try {
                $replaceAttributes['countryCode'] = [int]$countryNumericCode
                $changeCount++
            }
            catch {
                throw "CountryNumericCode '$countryNumericCode' must be an integer value."
            }
        }

        for ($i = 1; $i -le 15; $i++) {
            $columnName = "ExtensionAttribute$i"
            $attributeName = "extensionAttribute$i"
            $value = Get-TrimmedValue -Value $row.$columnName
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $replaceAttributes[$attributeName] = $value
                $changeCount++
            }
        }

        if ($replaceAttributes.Count -gt 0) {
            $setParams['Replace'] = $replaceAttributes
        }

        $clearAttributes = [System.Collections.Generic.List[string]]::new()
        $clearRequested = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ClearAttributes)
        foreach ($clearRequestedName in $clearRequested) {
            if ($clearAttributeMap.ContainsKey($clearRequestedName)) {
                $mapped = $clearAttributeMap[$clearRequestedName]
                if (-not $clearAttributes.Contains($mapped)) {
                    $clearAttributes.Add($mapped)
                }
            }
            else {
                if (-not $clearAttributes.Contains($clearRequestedName)) {
                    $clearAttributes.Add($clearRequestedName)
                }
            }
        }

        if ($clearAttributes.Count -gt 0) {
            $setParams['Clear'] = @($clearAttributes)
            $changeCount++
        }

        $accountExpirationDate = ConvertTo-NullableDateTime -Value $row.AccountExpirationDate
        if ($null -ne $accountExpirationDate) {
            $setParams['AccountExpirationDate'] = $accountExpirationDate
            $changeCount++
        }

        if ($setParams.Count -gt 1) {
            if ($PSCmdlet.ShouldProcess($resolvedKey, 'Update Active Directory user attributes')) {
                Invoke-WithRetry -OperationName "Update AD user attributes $resolvedKey" -ScriptBlock {
                    Set-ADUser @setParams -ErrorAction Stop
                }

                $messages.Add('Attributes updated.')
            }
            else {
                $messages.Add('Attribute updates skipped due to WhatIf.')
            }
        }

        $enabled = Get-NullableBool -Value $row.Enabled
        if ($null -ne $enabled) {
            $changeCount++
            if ($enabled -and -not [bool]$targetUser.Enabled) {
                if ($PSCmdlet.ShouldProcess($resolvedKey, 'Enable AD account')) {
                    Invoke-WithRetry -OperationName "Enable AD account $resolvedKey" -ScriptBlock {
                        Enable-ADAccount -Identity $targetUser.DistinguishedName -ErrorAction Stop
                    }
                    $messages.Add('Account enabled.')
                }
                else {
                    $messages.Add('Enable account skipped due to WhatIf.')
                }
            }
            elseif ((-not $enabled) -and [bool]$targetUser.Enabled) {
                if ($PSCmdlet.ShouldProcess($resolvedKey, 'Disable AD account')) {
                    Invoke-WithRetry -OperationName "Disable AD account $resolvedKey" -ScriptBlock {
                        Disable-ADAccount -Identity $targetUser.DistinguishedName -ErrorAction Stop
                    }
                    $messages.Add('Account disabled.')
                }
                else {
                    $messages.Add('Disable account skipped due to WhatIf.')
                }
            }
            else {
                $messages.Add('Account enabled state already matches requested value.')
            }
        }

        $accountPassword = Get-TrimmedValue -Value $row.AccountPassword
        if (-not [string]::IsNullOrWhiteSpace($accountPassword)) {
            $changeCount++
            if ($PSCmdlet.ShouldProcess($resolvedKey, 'Reset AD account password')) {
                $securePassword = ConvertTo-SecureString -String $accountPassword -AsPlainText -Force
                Invoke-WithRetry -OperationName "Reset AD account password $resolvedKey" -ScriptBlock {
                    Set-ADAccountPassword -Identity $targetUser.DistinguishedName -Reset -NewPassword $securePassword -ErrorAction Stop
                }
                $messages.Add('Password reset completed.')
            }
            else {
                $messages.Add('Password reset skipped due to WhatIf.')
            }
        }

        $targetPath = Get-TrimmedValue -Value $row.Path
        if (-not [string]::IsNullOrWhiteSpace($targetPath)) {
            $changeCount++
            $currentParentPath = ($targetUser.DistinguishedName -split ',', 2)[1]
            if ($targetPath -ieq $currentParentPath) {
                $messages.Add('User already in requested OU path.')
            }
            else {
                if ($PSCmdlet.ShouldProcess($resolvedKey, "Move AD object to '$targetPath'")) {
                    Invoke-WithRetry -OperationName "Move AD user $resolvedKey" -ScriptBlock {
                        Move-ADObject -Identity $targetUser.DistinguishedName -TargetPath $targetPath -ErrorAction Stop
                    }
                    $messages.Add("User moved to '$targetPath'.")
                }
                else {
                    $messages.Add('OU move skipped due to WhatIf.')
                }
            }
        }

        if ($changeCount -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'UpdateActiveDirectoryUser' -Status 'Skipped' -Message 'No changes were requested for this row.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'UpdateActiveDirectoryUser' -Status 'Completed' -Message ($messages -join ' ')))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'UpdateActiveDirectoryUser' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory user update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
