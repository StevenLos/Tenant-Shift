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
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M0002-Update-ActiveDirectoryContacts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Test-IsAdContact {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$AdObject
    )

    if ($null -eq $AdObject) {
        return $false
    }

    $classes = @($AdObject.ObjectClass | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() })
    return $classes -contains 'contact'
}

function Resolve-TargetAdContact {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IdentityType,

        [Parameter(Mandatory)]
        [string]$IdentityValue
    )

    $normalizedType = $IdentityType.Trim().ToLowerInvariant()
    switch ($normalizedType) {
        'name' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADObject -Filter "ObjectClass -eq 'contact' -and Name -eq '$escaped'" -Properties * -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'mail' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADObject -Filter "ObjectClass -eq 'contact' -and mail -eq '$escaped'" -Properties * -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'proxyaddress' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADObject -Filter "ObjectClass -eq 'contact' -and proxyAddresses -eq '$escaped'" -Properties * -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'distinguishedname' {
            $candidate = Get-ADObject -Identity $IdentityValue -Properties * -ErrorAction SilentlyContinue
            if (Test-IsAdContact -AdObject $candidate) {
                return $candidate
            }

            return $null
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            $candidate = Get-ADObject -Identity $guid -Properties * -ErrorAction SilentlyContinue
            if (Test-IsAdContact -AdObject $candidate) {
                return $candidate
            }

            return $null
        }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use Name, Mail, ProxyAddress, DistinguishedName, or ObjectGuid."
        }
    }
}

function Add-ReplaceField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$ReplaceAttributes,

        [Parameter(Mandatory)]
        [string]$AttributeName,

        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        $ReplaceAttributes[$AttributeName] = $text
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

$clearAttributeMap = @{
    DisplayName = 'displayName'
    GivenName = 'givenName'
    Initials = 'initials'
    Surname = 'sn'
    Description = 'description'
    Company = 'company'
    Department = 'department'
    Division = 'division'
    Title = 'title'
    Office = 'physicalDeliveryOfficeName'
    Manager = 'manager'
    OfficePhone = 'telephoneNumber'
    MobilePhone = 'mobile'
    HomePhone = 'homePhone'
    IpPhone = 'ipPhone'
    Fax = 'facsimileTelephoneNumber'
    Pager = 'pager'
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
    WebPage = 'wWWHomePage'
}

for ($i = 1; $i -le 15; $i++) {
    $clearAttributeMap["ExtensionAttribute$i"] = "extensionAttribute$i"
}

Write-Status -Message 'Starting Active Directory contact update script.'
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

        $targetContact = Invoke-WithRetry -OperationName "Resolve AD contact $primaryKey" -ScriptBlock {
            Resolve-TargetAdContact -IdentityType $identityType -IdentityValue $identityValue
        }

        if (-not $targetContact) {
            throw 'Target contact was not found.'
        }

        $resolvedKey = if (-not [string]::IsNullOrWhiteSpace((Get-TrimmedValue -Value $targetContact.mail))) { (Get-TrimmedValue -Value $targetContact.mail) } else { (Get-TrimmedValue -Value $targetContact.Name) }
        $messages = [System.Collections.Generic.List[string]]::new()
        $changeCount = 0

        $setParams = @{
            Identity = $targetContact.ObjectGuid
        }

        $replaceAttributes = @{}
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'displayName' -Value $row.DisplayName) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'givenName' -Value $row.GivenName) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'initials' -Value $row.Initials) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'sn' -Value $row.Surname) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'description' -Value $row.Description) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'company' -Value $row.Company) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'department' -Value $row.Department) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'division' -Value $row.Division) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'title' -Value $row.Title) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'physicalDeliveryOfficeName' -Value $row.Office) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'manager' -Value $row.Manager) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'telephoneNumber' -Value $row.OfficePhone) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'mobile' -Value $row.MobilePhone) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'homePhone' -Value $row.HomePhone) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'ipPhone' -Value $row.IpPhone) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'facsimileTelephoneNumber' -Value $row.Fax) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'pager' -Value $row.Pager) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'mail' -Value $row.Mail) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'mailNickname' -Value $row.MailNickname) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'targetAddress' -Value $row.TargetAddress) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'streetAddress' -Value $row.StreetAddress) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'postOfficeBox' -Value $row.PostOfficeBox) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'l' -Value $row.City) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'st' -Value $row.StateOrProvince) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'postalCode' -Value $row.PostalCode) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'c' -Value $row.CountryCode) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'co' -Value $row.CountryName) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'wWWHomePage' -Value $row.WebPage) { $changeCount++ }

        $proxyAddresses = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ProxyAddresses)
        if ($proxyAddresses.Count -gt 0) {
            $replaceAttributes['proxyAddresses'] = $proxyAddresses
            $changeCount++
        }

        $hideFromAddressLists = Get-NullableBool -Value $row.HideFromAddressLists
        if ($null -ne $hideFromAddressLists) {
            $replaceAttributes['msExchHideFromAddressLists'] = $hideFromAddressLists
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

        if ($setParams.Count -gt 1) {
            if ($PSCmdlet.ShouldProcess($resolvedKey, 'Update Active Directory contact attributes')) {
                Invoke-WithRetry -OperationName "Update AD contact attributes $resolvedKey" -ScriptBlock {
                    Set-ADObject @setParams -ErrorAction Stop
                }

                $messages.Add('Attributes updated.')
            }
            else {
                $messages.Add('Attribute updates skipped due to WhatIf.')
            }
        }

        $newName = Get-TrimmedValue -Value $row.Name
        if (-not [string]::IsNullOrWhiteSpace($newName) -and $newName -ne (Get-TrimmedValue -Value $targetContact.Name)) {
            $changeCount++

            if ($PSCmdlet.ShouldProcess($resolvedKey, "Rename AD contact to '$newName'")) {
                Invoke-WithRetry -OperationName "Rename AD contact $resolvedKey" -ScriptBlock {
                    Rename-ADObject -Identity $targetContact.ObjectGuid -NewName $newName -ErrorAction Stop
                }

                $messages.Add("Contact renamed to '$newName'.")
            }
            else {
                $messages.Add('Rename skipped due to WhatIf.')
            }
        }

        $targetPath = Get-TrimmedValue -Value $row.Path
        if (-not [string]::IsNullOrWhiteSpace($targetPath)) {
            $currentContact = Invoke-WithRetry -OperationName "Reload AD contact $resolvedKey" -ScriptBlock {
                Get-ADObject -Identity $targetContact.ObjectGuid -Properties DistinguishedName -ErrorAction Stop
            }

            $currentParentPath = ($currentContact.DistinguishedName -split ',', 2)[1]
            if ($targetPath -ieq $currentParentPath) {
                $messages.Add('Contact already in requested OU path.')
            }
            else {
                $changeCount++
                if ($PSCmdlet.ShouldProcess($resolvedKey, "Move AD contact to '$targetPath'")) {
                    Invoke-WithRetry -OperationName "Move AD contact $resolvedKey" -ScriptBlock {
                        Move-ADObject -Identity $targetContact.ObjectGuid -TargetPath $targetPath -ErrorAction Stop
                    }

                    $messages.Add("Contact moved to '$targetPath'.")
                }
                else {
                    $messages.Add('OU move skipped due to WhatIf.')
                }
            }
        }

        if ($changeCount -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'UpdateActiveDirectoryContact' -Status 'Skipped' -Message 'No changes were requested for this row.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'UpdateActiveDirectoryContact' -Status 'Completed' -Message ($messages -join ' ')))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'UpdateActiveDirectoryContact' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory contact update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
