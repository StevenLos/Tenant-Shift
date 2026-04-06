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
    Gets ActiveDirectoryUsers and exports results to CSV.

.DESCRIPTION
    Gets ActiveDirectoryUsers from Active Directory and writes the results to a CSV file.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all objects in scope (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all objects in scope rather than processing from an input CSV file. Uses the DiscoverAll parameter set.

.PARAMETER SearchBase
    Distinguished name of the Active Directory OU to scope the discovery. If omitted, searches the entire domain.

.PARAMETER Server
    Active Directory domain controller to target. If omitted, uses the default DC for the current domain.

.PARAMETER MaxObjects
    Maximum number of objects to retrieve. 0 (default) means no limit.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-D0001-Get-ActiveDirectoryUsers.ps1 -InputCsvPath .\0001.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D0001-Get-ActiveDirectoryUsers.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: ActiveDirectory
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_SM-D0001-Get-ActiveDirectoryUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Resolve-UsersByScope {
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
                Filter     = '*'
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADUser @params)
        }
        'samaccountname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            $params = @{
                Filter      = "SamAccountName -eq '$escaped'"
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADUser @params)
        }
        'userprincipalname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            $params = @{
                Filter      = "UserPrincipalName -eq '$escaped'"
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADUser @params)
        }
        'distinguishedname' {
            $params = @{
                Identity    = $IdentityValue
                Properties  = '*'
                ErrorAction = 'SilentlyContinue'
            }

            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            $user = Get-ADUser @params
            if ($user) {
                return @($user)
            }

            return @()
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            $params = @{
                Identity    = $guid
                Properties  = '*'
                ErrorAction = 'SilentlyContinue'
            }

            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            $user = Get-ADUser @params
            if ($user) {
                return @($user)
            }

            return @()
        }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use All, SamAccountName, UserPrincipalName, DistinguishedName, or ObjectGuid."
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
    'SID',
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
    'ChangePasswordAtLogon',
    'PasswordNeverExpires',
    'CannotChangePassword',
    'SmartcardLogonRequired',
    'AccountExpirationDate',
    'LogonWorkstations',
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
    'mailNickname',
    'proxyAddresses',
    'targetAddress',
    'msExchHideFromAddressLists',
    'StreetAddress',
    'POBox',
    'City',
    'State',
    'PostalCode',
    'c',
    'co',
    'countryCode',
    'HomePhone',
    'ipPhone',
    'Fax',
    'Pager',
    'HomePage',
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
    'LastLogonDate',
    'LastBadPasswordAttempt',
    'PasswordLastSet',
    'LockedOut',
    'MemberOf'
)

Write-Status -Message 'Starting Active Directory user inventory script.'
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

    Write-Status -Message "DiscoverAll enabled for Active Directory users. SearchBase='$resolvedSearchBase'." -Level WARN
    $rows = @([PSCustomObject]@{
            IdentityType  = 'All'
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
        $users = Invoke-WithRetry -OperationName "Load users for $primaryKey" -ScriptBlock {
            Resolve-UsersByScope -IdentityType $identityType -IdentityValue $identityValue -PropertyNames $propertyNames -SearchBase $effectiveSearchBase -Server $resolvedServer
        }

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $users.Count -gt $MaxObjects) {
            $users = @($users | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($users.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryUser' -Status 'NotFound' -Message 'No matching users were found.' -Data ([ordered]@{
                            IdentityTypeRequested = $identityType
                            IdentityValueRequested = $identityValue
                            DistinguishedName = ''
                            ObjectGuid = ''
                            SID = ''
                            SamAccountName = ''
                            UserPrincipalName = ''
                            GivenName = ''
                            Initials = ''
                            Surname = ''
                            DisplayName = ''
                            Name = ''
                            Description = ''
                            EmployeeID = ''
                            EmployeeNumber = ''
                            EmployeeType = ''
                            Enabled = ''
                            ChangePasswordAtLogon = ''
                            PasswordNeverExpires = ''
                            CannotChangePassword = ''
                            SmartcardLogonRequired = ''
                            AccountExpirationDate = ''
                            UserWorkstations = ''
                            ScriptPath = ''
                            ProfilePath = ''
                            HomeDirectory = ''
                            HomeDrive = ''
                            Title = ''
                            Department = ''
                            Company = ''
                            Division = ''
                            Office = ''
                            Manager = ''
                            OfficePhone = ''
                            MobilePhone = ''
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
                            HomePhone = ''
                            IpPhone = ''
                            Fax = ''
                            Pager = ''
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
                            LastLogonDate = ''
                            LastBadPasswordAttempt = ''
                            PasswordLastSet = ''
                            LockedOut = ''
                            MemberOf = ''
                        })))

            $rowNumber++
            continue
        }

        foreach ($user in @($users | Sort-Object -Property UserPrincipalName, SamAccountName, DistinguishedName)) {
            $userPrimaryKey = if (-not [string]::IsNullOrWhiteSpace([string]$user.UserPrincipalName)) { ([string]$user.UserPrincipalName).Trim() } else { ([string]$user.SamAccountName).Trim() }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrimaryKey -Action 'GetActiveDirectoryUser' -Status 'Completed' -Message 'User exported.' -Data ([ordered]@{
                            IdentityTypeRequested = $identityType
                            IdentityValueRequested = $identityValue
                            DistinguishedName = Get-TrimmedValue -Value $user.DistinguishedName
                            ObjectGuid = Get-TrimmedValue -Value $user.ObjectGuid
                            SID = Get-TrimmedValue -Value $user.SID
                            SamAccountName = Get-TrimmedValue -Value $user.SamAccountName
                            UserPrincipalName = Get-TrimmedValue -Value $user.UserPrincipalName
                            GivenName = Get-TrimmedValue -Value $user.GivenName
                            Initials = Get-TrimmedValue -Value $user.Initials
                            Surname = Get-TrimmedValue -Value $user.Surname
                            DisplayName = Get-TrimmedValue -Value $user.DisplayName
                            Name = Get-TrimmedValue -Value $user.Name
                            Description = Get-TrimmedValue -Value $user.Description
                            EmployeeID = Get-TrimmedValue -Value $user.EmployeeID
                            EmployeeNumber = Get-TrimmedValue -Value $user.EmployeeNumber
                            EmployeeType = Get-TrimmedValue -Value $user.EmployeeType
                            Enabled = [string]$user.Enabled
                            ChangePasswordAtLogon = [string]$user.ChangePasswordAtLogon
                            PasswordNeverExpires = [string]$user.PasswordNeverExpires
                            CannotChangePassword = [string]$user.CannotChangePassword
                            SmartcardLogonRequired = [string]$user.SmartcardLogonRequired
                            AccountExpirationDate = Get-TrimmedValue -Value $user.AccountExpirationDate
                            UserWorkstations = Get-TrimmedValue -Value $user.LogonWorkstations
                            ScriptPath = Get-TrimmedValue -Value $user.ScriptPath
                            ProfilePath = Get-TrimmedValue -Value $user.ProfilePath
                            HomeDirectory = Get-TrimmedValue -Value $user.HomeDirectory
                            HomeDrive = Get-TrimmedValue -Value $user.HomeDrive
                            Title = Get-TrimmedValue -Value $user.Title
                            Department = Get-TrimmedValue -Value $user.Department
                            Company = Get-TrimmedValue -Value $user.Company
                            Division = Get-TrimmedValue -Value $user.Division
                            Office = Get-TrimmedValue -Value $user.Office
                            Manager = Get-TrimmedValue -Value $user.Manager
                            OfficePhone = Get-TrimmedValue -Value $user.OfficePhone
                            MobilePhone = Get-TrimmedValue -Value $user.MobilePhone
                            Mail = Get-TrimmedValue -Value $user.Mail
                            MailNickname = Get-TrimmedValue -Value $user.mailNickname
                            ProxyAddresses = Convert-MultiValueToString -Value $user.proxyAddresses
                            TargetAddress = Get-TrimmedValue -Value $user.targetAddress
                            HideFromAddressLists = [string]$user.msExchHideFromAddressLists
                            StreetAddress = Get-TrimmedValue -Value $user.StreetAddress
                            PostOfficeBox = Get-TrimmedValue -Value $user.POBox
                            City = Get-TrimmedValue -Value $user.City
                            StateOrProvince = Get-TrimmedValue -Value $user.State
                            PostalCode = Get-TrimmedValue -Value $user.PostalCode
                            CountryCode = Get-TrimmedValue -Value $user.c
                            CountryName = Get-TrimmedValue -Value $user.co
                            CountryNumericCode = Get-TrimmedValue -Value $user.countryCode
                            HomePhone = Get-TrimmedValue -Value $user.HomePhone
                            IpPhone = Get-TrimmedValue -Value $user.ipPhone
                            Fax = Get-TrimmedValue -Value $user.Fax
                            Pager = Get-TrimmedValue -Value $user.Pager
                            WebPage = Get-TrimmedValue -Value $user.HomePage
                            ExtensionAttribute1 = Get-TrimmedValue -Value $user.extensionAttribute1
                            ExtensionAttribute2 = Get-TrimmedValue -Value $user.extensionAttribute2
                            ExtensionAttribute3 = Get-TrimmedValue -Value $user.extensionAttribute3
                            ExtensionAttribute4 = Get-TrimmedValue -Value $user.extensionAttribute4
                            ExtensionAttribute5 = Get-TrimmedValue -Value $user.extensionAttribute5
                            ExtensionAttribute6 = Get-TrimmedValue -Value $user.extensionAttribute6
                            ExtensionAttribute7 = Get-TrimmedValue -Value $user.extensionAttribute7
                            ExtensionAttribute8 = Get-TrimmedValue -Value $user.extensionAttribute8
                            ExtensionAttribute9 = Get-TrimmedValue -Value $user.extensionAttribute9
                            ExtensionAttribute10 = Get-TrimmedValue -Value $user.extensionAttribute10
                            ExtensionAttribute11 = Get-TrimmedValue -Value $user.extensionAttribute11
                            ExtensionAttribute12 = Get-TrimmedValue -Value $user.extensionAttribute12
                            ExtensionAttribute13 = Get-TrimmedValue -Value $user.extensionAttribute13
                            ExtensionAttribute14 = Get-TrimmedValue -Value $user.extensionAttribute14
                            ExtensionAttribute15 = Get-TrimmedValue -Value $user.extensionAttribute15
                            CanonicalName = Get-TrimmedValue -Value $user.CanonicalName
                            WhenCreated = Get-TrimmedValue -Value $user.whenCreated
                            WhenChanged = Get-TrimmedValue -Value $user.whenChanged
                            LastLogonDate = Get-TrimmedValue -Value $user.LastLogonDate
                            LastBadPasswordAttempt = Get-TrimmedValue -Value $user.LastBadPasswordAttempt
                            PasswordLastSet = Get-TrimmedValue -Value $user.PasswordLastSet
                            LockedOut = [string]$user.LockedOut
                            MemberOf = Convert-MultiValueToString -Value $user.MemberOf
                        })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryUser' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                        IdentityTypeRequested = $identityType
                        IdentityValueRequested = $identityValue
                        DistinguishedName = ''
                        ObjectGuid = ''
                        SID = ''
                        SamAccountName = ''
                        UserPrincipalName = ''
                        GivenName = ''
                        Initials = ''
                        Surname = ''
                        DisplayName = ''
                        Name = ''
                        Description = ''
                        EmployeeID = ''
                        EmployeeNumber = ''
                        EmployeeType = ''
                        Enabled = ''
                        ChangePasswordAtLogon = ''
                        PasswordNeverExpires = ''
                        CannotChangePassword = ''
                        SmartcardLogonRequired = ''
                        AccountExpirationDate = ''
                        UserWorkstations = ''
                        ScriptPath = ''
                        ProfilePath = ''
                        HomeDirectory = ''
                        HomeDrive = ''
                        Title = ''
                        Department = ''
                        Company = ''
                        Division = ''
                        Office = ''
                        Manager = ''
                        OfficePhone = ''
                        MobilePhone = ''
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
                        HomePhone = ''
                        IpPhone = ''
                        Fax = ''
                        Pager = ''
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
                        LastLogonDate = ''
                        LastBadPasswordAttempt = ''
                        PasswordLastSet = ''
                        LockedOut = ''
                        MemberOf = ''
                    })))
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
Write-Status -Message 'Active Directory user inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
