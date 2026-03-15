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

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$SearchBase,

    [Parameter(ParameterSetName = 'FromCsv')]
    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$Server,

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [ValidateRange(0, 10000000)]
    [int]$MaxObjects = 0,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR0002-Get-ActiveDirectoryContacts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Escape-LdapFilterValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    $builder = [System.Text.StringBuilder]::new()
    foreach ($char in $Value.ToCharArray()) {
        switch ($char) {
            '\\' { [void]$builder.Append('\\5c') }
            '*' { [void]$builder.Append('\\2a') }
            '(' { [void]$builder.Append('\\28') }
            ')' { [void]$builder.Append('\\29') }
            ([char]0) { [void]$builder.Append('\\00') }
            default { [void]$builder.Append($char) }
        }
    }

    return $builder.ToString()
}

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

function Resolve-ContactsByScope {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IdentityType,

        [Parameter(Mandatory)]
        [string]$IdentityValue,

        [Parameter(Mandatory)]
        [string[]]$PropertyNames,

        [AllowEmptyString()]
        [string]$SearchBase,

        [AllowEmptyString()]
        [string]$Server
    )

    $normalizedType = $IdentityType.Trim().ToLowerInvariant()

    switch ($normalizedType) {
        'all' {
            if ($IdentityValue -ne '*') {
                throw "IdentityValue must be '*' when IdentityType is 'All'."
            }

            $params = @{
                LDAPFilter = '(objectClass=contact)'
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADObject @params)
        }
        'name' {
            $escaped = Escape-LdapFilterValue -Value $IdentityValue
            $params = @{
                LDAPFilter = "(&(objectClass=contact)(name=$escaped))"
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADObject @params)
        }
        'mail' {
            $escaped = Escape-LdapFilterValue -Value $IdentityValue
            $params = @{
                LDAPFilter = "(&(objectClass=contact)(mail=$escaped))"
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADObject @params)
        }
        'distinguishedname' {
            $params = @{
                Identity = $IdentityValue
                Properties  = '*'
                ErrorAction = 'SilentlyContinue'
            }

            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            $contact = Get-ADObject @params
            if ($contact -and (Test-IsAdContact -AdObject $contact)) {
                return @($contact)
            }

            return @()
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            $params = @{
                Identity = $guid
                Properties  = '*'
                ErrorAction = 'SilentlyContinue'
            }

            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            $contact = Get-ADObject @params
            if ($contact -and (Test-IsAdContact -AdObject $contact)) {
                return @($contact)
            }

            return @()
        }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use All, Name, Mail, DistinguishedName, or ObjectGuid."
        }
    }
}

function Convert-MultiValueToString {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [string]) {
        return ([string]$Value).Trim()
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        $items = [System.Collections.Generic.List[string]]::new()
        foreach ($item in $Value) {
            $text = ([string]$item).Trim()
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $items.Add($text)
            }
        }

        return (@($items | Sort-Object -Unique) -join ';')
    }

    return ([string]$Value).Trim()
}

function New-InventoryResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$RowNumber,

        [Parameter(Mandatory)]
        [string]$PrimaryKey,

        [Parameter(Mandatory)]
        [string]$Action,

        [Parameter(Mandatory)]
        [string]$Status,

        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory)]
        [hashtable]$Data
    )

    $base = New-ResultObject -RowNumber $RowNumber -PrimaryKey $PrimaryKey -Action $Action -Status $Status -Message $Message
    $ordered = [ordered]@{}

    foreach ($prop in $base.PSObject.Properties.Name) {
        $ordered[$prop] = $base.$prop
    }

    foreach ($key in $Data.Keys) {
        $ordered[$key] = $Data[$key]
    }

    return [PSCustomObject]$ordered
}

$requiredHeaders = @(
    'IdentityType',
    'IdentityValue'
)

$propertyNames = @(
    'DistinguishedName',
    'ObjectGuid',
    'objectClass',
    'Name',
    'displayName',
    'givenName',
    'initials',
    'sn',
    'description',
    'company',
    'department',
    'division',
    'title',
    'physicalDeliveryOfficeName',
    'manager',
    'telephoneNumber',
    'mobile',
    'homePhone',
    'ipPhone',
    'facsimileTelephoneNumber',
    'pager',
    'mail',
    'mailNickname',
    'proxyAddresses',
    'targetAddress',
    'msExchHideFromAddressLists',
    'streetAddress',
    'postOfficeBox',
    'l',
    'st',
    'postalCode',
    'c',
    'co',
    'countryCode',
    'wWWHomePage',
    'extensionAttribute1',
    'extensionAttribute2',
    'extensionAttribute3',
    'extensionAttribute4',
    'extensionAttribute5',
    'extensionAttribute6',
    'extensionAttribute7',
    'extensionAttribute8',
    'extensionAttribute9',
    'extensionAttribute10',
    'extensionAttribute11',
    'extensionAttribute12',
    'extensionAttribute13',
    'extensionAttribute14',
    'extensionAttribute15',
    'CanonicalName',
    'whenCreated',
    'whenChanged',
    'showInAddressBook',
    'legacyExchangeDN',
    'memberOf'
)

function New-EmptyContactData {
    return [ordered]@{
        IdentityTypeRequested = ''
        IdentityValueRequested = ''
        DistinguishedName = ''
        ObjectGuid = ''
        ObjectClass = ''
        Name = ''
        DisplayName = ''
        GivenName = ''
        Initials = ''
        Surname = ''
        Description = ''
        Company = ''
        Department = ''
        Division = ''
        Title = ''
        Office = ''
        Manager = ''
        OfficePhone = ''
        MobilePhone = ''
        HomePhone = ''
        IpPhone = ''
        Fax = ''
        Pager = ''
        Mail = ''
        MailNickname = ''
        ProxyAddresses = ''
        TargetAddress = ''
        HideFromAddressLists = ''
        StreetAddress = ''
        PostOfficeBox = ''
        City = ''
        StateOrProvince = ''
        PostalCode = ''
        CountryCode = ''
        CountryName = ''
        CountryNumericCode = ''
        WebPage = ''
        ExtensionAttribute1 = ''
        ExtensionAttribute2 = ''
        ExtensionAttribute3 = ''
        ExtensionAttribute4 = ''
        ExtensionAttribute5 = ''
        ExtensionAttribute6 = ''
        ExtensionAttribute7 = ''
        ExtensionAttribute8 = ''
        ExtensionAttribute9 = ''
        ExtensionAttribute10 = ''
        ExtensionAttribute11 = ''
        ExtensionAttribute12 = ''
        ExtensionAttribute13 = ''
        ExtensionAttribute14 = ''
        ExtensionAttribute15 = ''
        CanonicalName = ''
        WhenCreated = ''
        WhenChanged = ''
        ShowInAddressBook = ''
        LegacyExchangeDn = ''
        MemberOf = ''
    }
}

Write-Status -Message 'Starting Active Directory contact inventory script.'
Ensure-ActiveDirectoryConnection

$scopeMode = 'Csv'
$resolvedServer = Get-TrimmedValue -Value $Server
$resolvedSearchBase = ''
$runWasTruncated = $false

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    $resolvedSearchBase = Get-TrimmedValue -Value $SearchBase
    if ([string]::IsNullOrWhiteSpace($resolvedSearchBase)) {
        $domainParams = @{
            ErrorAction = 'Stop'
        }

        if (-not [string]::IsNullOrWhiteSpace($resolvedServer)) {
            $domainParams['Server'] = $resolvedServer
        }

        $resolvedSearchBase = (Get-ADDomain @domainParams).DistinguishedName
    }

    if ([string]::IsNullOrWhiteSpace($resolvedSearchBase)) {
        throw 'Unable to determine SearchBase for DiscoverAll mode.'
    }

    Write-Status -Message "DiscoverAll enabled for Active Directory contacts. SearchBase='$resolvedSearchBase'." -Level WARN
    $rows = @([PSCustomObject]@{
            IdentityType = 'All'
            IdentityValue = '*'
        })
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

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

        $effectiveSearchBase = if ($scopeMode -eq 'DiscoverAll') { $resolvedSearchBase } else { '' }
        $contacts = Invoke-WithRetry -OperationName "Load contacts for $primaryKey" -ScriptBlock {
            Resolve-ContactsByScope -IdentityType $identityType -IdentityValue $identityValue -PropertyNames $propertyNames -SearchBase $effectiveSearchBase -Server $resolvedServer
        }

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $contacts.Count -gt $MaxObjects) {
            $contacts = @($contacts | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($contacts.Count -eq 0) {
            $data = New-EmptyContactData
            $data['IdentityTypeRequested'] = $identityType
            $data['IdentityValueRequested'] = $identityValue
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryContact' -Status 'NotFound' -Message 'No matching contacts were found.' -Data $data))

            $rowNumber++
            continue
        }

        foreach ($contact in @($contacts | Sort-Object -Property mail, Name, DistinguishedName)) {
            $contactPrimaryKey = if (-not [string]::IsNullOrWhiteSpace((Get-TrimmedValue -Value $contact.mail))) { (Get-TrimmedValue -Value $contact.mail) } else { (Get-TrimmedValue -Value $contact.Name) }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $contactPrimaryKey -Action 'GetActiveDirectoryContact' -Status 'Completed' -Message 'Contact exported.' -Data ([ordered]@{
                            IdentityTypeRequested = $identityType
                            IdentityValueRequested = $identityValue
                            DistinguishedName = Get-TrimmedValue -Value $contact.DistinguishedName
                            ObjectGuid = Get-TrimmedValue -Value $contact.ObjectGuid
                            ObjectClass = Convert-MultiValueToString -Value $contact.objectClass
                            Name = Get-TrimmedValue -Value $contact.Name
                            DisplayName = Get-TrimmedValue -Value $contact.displayName
                            GivenName = Get-TrimmedValue -Value $contact.givenName
                            Initials = Get-TrimmedValue -Value $contact.initials
                            Surname = Get-TrimmedValue -Value $contact.sn
                            Description = Get-TrimmedValue -Value $contact.description
                            Company = Get-TrimmedValue -Value $contact.company
                            Department = Get-TrimmedValue -Value $contact.department
                            Division = Get-TrimmedValue -Value $contact.division
                            Title = Get-TrimmedValue -Value $contact.title
                            Office = Get-TrimmedValue -Value $contact.physicalDeliveryOfficeName
                            Manager = Get-TrimmedValue -Value $contact.manager
                            OfficePhone = Get-TrimmedValue -Value $contact.telephoneNumber
                            MobilePhone = Get-TrimmedValue -Value $contact.mobile
                            HomePhone = Get-TrimmedValue -Value $contact.homePhone
                            IpPhone = Get-TrimmedValue -Value $contact.ipPhone
                            Fax = Get-TrimmedValue -Value $contact.facsimileTelephoneNumber
                            Pager = Get-TrimmedValue -Value $contact.pager
                            Mail = Get-TrimmedValue -Value $contact.mail
                            MailNickname = Get-TrimmedValue -Value $contact.mailNickname
                            ProxyAddresses = Convert-MultiValueToString -Value $contact.proxyAddresses
                            TargetAddress = Get-TrimmedValue -Value $contact.targetAddress
                            HideFromAddressLists = [string]$contact.msExchHideFromAddressLists
                            StreetAddress = Get-TrimmedValue -Value $contact.streetAddress
                            PostOfficeBox = Get-TrimmedValue -Value $contact.postOfficeBox
                            City = Get-TrimmedValue -Value $contact.l
                            StateOrProvince = Get-TrimmedValue -Value $contact.st
                            PostalCode = Get-TrimmedValue -Value $contact.postalCode
                            CountryCode = Get-TrimmedValue -Value $contact.c
                            CountryName = Get-TrimmedValue -Value $contact.co
                            CountryNumericCode = Get-TrimmedValue -Value $contact.countryCode
                            WebPage = Get-TrimmedValue -Value $contact.wWWHomePage
                            ExtensionAttribute1 = Get-TrimmedValue -Value $contact.extensionAttribute1
                            ExtensionAttribute2 = Get-TrimmedValue -Value $contact.extensionAttribute2
                            ExtensionAttribute3 = Get-TrimmedValue -Value $contact.extensionAttribute3
                            ExtensionAttribute4 = Get-TrimmedValue -Value $contact.extensionAttribute4
                            ExtensionAttribute5 = Get-TrimmedValue -Value $contact.extensionAttribute5
                            ExtensionAttribute6 = Get-TrimmedValue -Value $contact.extensionAttribute6
                            ExtensionAttribute7 = Get-TrimmedValue -Value $contact.extensionAttribute7
                            ExtensionAttribute8 = Get-TrimmedValue -Value $contact.extensionAttribute8
                            ExtensionAttribute9 = Get-TrimmedValue -Value $contact.extensionAttribute9
                            ExtensionAttribute10 = Get-TrimmedValue -Value $contact.extensionAttribute10
                            ExtensionAttribute11 = Get-TrimmedValue -Value $contact.extensionAttribute11
                            ExtensionAttribute12 = Get-TrimmedValue -Value $contact.extensionAttribute12
                            ExtensionAttribute13 = Get-TrimmedValue -Value $contact.extensionAttribute13
                            ExtensionAttribute14 = Get-TrimmedValue -Value $contact.extensionAttribute14
                            ExtensionAttribute15 = Get-TrimmedValue -Value $contact.extensionAttribute15
                            CanonicalName = Get-TrimmedValue -Value $contact.CanonicalName
                            WhenCreated = Get-TrimmedValue -Value $contact.whenCreated
                            WhenChanged = Get-TrimmedValue -Value $contact.whenChanged
                            ShowInAddressBook = Convert-MultiValueToString -Value $contact.showInAddressBook
                            LegacyExchangeDn = Get-TrimmedValue -Value $contact.legacyExchangeDN
                            MemberOf = Convert-MultiValueToString -Value $contact.memberOf
                        })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $data = New-EmptyContactData
        $data['IdentityTypeRequested'] = $identityType
        $data['IdentityValueRequested'] = $identityValue
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryContact' -Status 'Failed' -Message $_.Exception.Message -Data $data))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeSearchBase' -NotePropertyValue $resolvedSearchBase -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeServer' -NotePropertyValue $resolvedServer -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeMaxObjects' -NotePropertyValue ($(if ($scopeMode -eq 'DiscoverAll') { [string]$MaxObjects } else { '' })) -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeWasTruncated' -NotePropertyValue ([string]$runWasTruncated) -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory contact inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
