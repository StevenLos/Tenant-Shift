<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-153500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets EntraUsers and exports results to CSV.

.DESCRIPTION
    Gets EntraUsers from Microsoft 365 and writes the results to a CSV file.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all objects in scope (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all objects in scope rather than processing from an input CSV file. Uses the DiscoverAll parameter set.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-D3001-Get-EntraUsers.ps1 -InputCsvPath .\3001.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D3001-Get-EntraUsers.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Users
    Required roles:   Global Administrator or appropriate workload-specific role
    Limitations:      None known.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-MEID-0010-Get-EntraUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

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

function Get-UserExtensionAttributeValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$User,

        [Parameter(Mandatory)]
        [string]$AttributeName
    )

    $extensionObject = $null
    if ($User.PSObject.Properties.Name -contains 'OnPremisesExtensionAttributes') {
        $extensionObject = $User.OnPremisesExtensionAttributes
    }

    if (-not $extensionObject) {
        return ''
    }

    if ($extensionObject.PSObject.Properties.Name -contains $AttributeName) {
        return Get-TrimmedValue -Value $extensionObject.$AttributeName
    }

    if ($extensionObject.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $additional = $extensionObject.AdditionalProperties
        if ($additional) {
            try {
                if ($additional.ContainsKey($AttributeName)) {
                    return Get-TrimmedValue -Value $additional[$AttributeName]
                }
            }
            catch {
                # Best effort only.
            }
        }
    }

    return ''
}

function Get-AssignmentSkuIds {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$AssignedLicenses
    )

    if ($null -eq $AssignedLicenses) {
        return ''
    }

    $ids = [System.Collections.Generic.List[string]]::new()
    foreach ($assignment in @($AssignedLicenses)) {
        $skuId = Get-TrimmedValue -Value $assignment.SkuId
        if (-not [string]::IsNullOrWhiteSpace($skuId) -and -not $ids.Contains($skuId)) {
            $ids.Add($skuId)
        }
    }

    return (@($ids | Sort-Object) -join ';')
}

function Get-AssignedPlanNames {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$AssignedPlans
    )

    if ($null -eq $AssignedPlans) {
        return ''
    }

    $planNames = [System.Collections.Generic.List[string]]::new()
    foreach ($plan in @($AssignedPlans)) {
        $name = Get-TrimmedValue -Value $plan.ServicePlanId
        if (-not [string]::IsNullOrWhiteSpace($name) -and -not $planNames.Contains($name)) {
            $planNames.Add($name)
        }
    }

    return (@($planNames | Sort-Object) -join ';')
}

$requiredHeaders = @(
    'UserPrincipalName'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'UserId',
    'UserPrincipalName',
    'Mail',
    'MailNickname',
    'DisplayName',
    'GivenName',
    'Surname',
    'UserType',
    'AccountEnabled',
    'Department',
    'JobTitle',
    'CompanyName',
    'OfficeLocation',
    'EmployeeId',
    'EmployeeType',
    'EmployeeHireDate',
    'UsageLocation',
    'PreferredLanguage',
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
    'CreatedDateTime',
    'LastPasswordChangeDateTime',
    'AssignedLicenseSkuIds',
    'AssignedPlanServiceIds',
    'OnPremisesSyncEnabled',
    'OnPremisesImmutableId',
    'OnPremisesDistinguishedName',
    'OnPremisesDomainName',
    'OnPremisesSamAccountName',
    'OnPremisesSecurityIdentifier',
    'OnPremisesLastSyncDateTime',
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

Write-Status -Message 'Starting Entra ID user inventory script (expanded field model).'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('User.Read.All', 'Directory.Read.All')

$scopeMode = 'Csv'
if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. CSV input is bypassed.' -Level WARN

    $discoverRow = [ordered]@{}
    foreach ($header in $requiredHeaders) {
        $discoverRow[$header] = '*'
    }

    $rows = @([PSCustomObject]$discoverRow)
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}
$results = [System.Collections.Generic.List[object]]::new()

$userSelect = 'id,userPrincipalName,displayName,givenName,surname,mailNickname,userType,accountEnabled,usageLocation,preferredLanguage,department,jobTitle,companyName,officeLocation,employeeId,employeeType,employeeHireDate,mobilePhone,businessPhones,faxNumber,otherMails,streetAddress,city,state,postalCode,country,passwordPolicies,mail,onPremisesExtensionAttributes,createdDateTime,lastPasswordChangeDateTime,onPremisesSyncEnabled,onPremisesImmutableId,onPremisesDistinguishedName,onPremisesDomainName,onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesLastSyncDateTime,assignedLicenses,assignedPlans'

$rowNumber = 1
foreach ($row in $rows) {
    $userPrincipalName = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
            throw 'UserPrincipalName is required. Use * to inventory all users.'
        }

        $users = @()
        if ($userPrincipalName -eq '*') {
            $users = @(Invoke-WithRetry -OperationName 'Load all users' -ScriptBlock {
                Get-MgUser -All -Property $userSelect -ErrorAction Stop
            })
        }
        else {
            $escapedUpn = Escape-ODataString -Value $userPrincipalName
            $users = @(Invoke-WithRetry -OperationName "Lookup user $userPrincipalName" -ScriptBlock {
                Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -Property $userSelect -ErrorAction Stop
            })
        }

        if ($users.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'GetEntraUser' -Status 'NotFound' -Message 'No matching users were found.' -Data ([ordered]@{
                        UserId                         = ''
                        UserPrincipalName              = $userPrincipalName
                        DisplayName                    = ''
                        GivenName                      = ''
                        Surname                        = ''
                        MailNickname                   = ''
                        UserType                       = ''
                        AccountEnabled                 = ''
                        UsageLocation                  = ''
                        PreferredLanguage              = ''
                        Department                     = ''
                        JobTitle                       = ''
                        CompanyName                    = ''
                        OfficeLocation                 = ''
                        EmployeeId                     = ''
                        EmployeeType                   = ''
                        EmployeeHireDate               = ''
                        MobilePhone                    = ''
                        BusinessPhones                 = ''
                        FaxNumber                      = ''
                        OtherMails                     = ''
                        StreetAddress                  = ''
                        City                           = ''
                        State                          = ''
                        PostalCode                     = ''
                        Country                        = ''
                        PasswordPolicies               = ''
                        ExtensionAttribute1            = ''
                        ExtensionAttribute2            = ''
                        ExtensionAttribute3            = ''
                        ExtensionAttribute4            = ''
                        ExtensionAttribute5            = ''
                        ExtensionAttribute6            = ''
                        ExtensionAttribute7            = ''
                        ExtensionAttribute8            = ''
                        ExtensionAttribute9            = ''
                        ExtensionAttribute10           = ''
                        ExtensionAttribute11           = ''
                        ExtensionAttribute12           = ''
                        ExtensionAttribute13           = ''
                        ExtensionAttribute14           = ''
                        ExtensionAttribute15           = ''
                        Mail                           = ''
                        CreatedDateTime                = ''
                        LastPasswordChangeDateTime     = ''
                        OnPremisesSyncEnabled          = ''
                        OnPremisesImmutableId          = ''
                        OnPremisesDistinguishedName    = ''
                        OnPremisesDomainName           = ''
                        OnPremisesSamAccountName       = ''
                        OnPremisesSecurityIdentifier   = ''
                        OnPremisesLastSyncDateTime     = ''
                        AssignedLicenseSkuIds          = ''
                        AssignedPlanServiceIds         = ''
                    })))
            $rowNumber++
            continue
        }

        $sortedUsers = @($users | Sort-Object -Property UserPrincipalName, Id)
        foreach ($user in $sortedUsers) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey ([string]$user.UserPrincipalName) -Action 'GetEntraUser' -Status 'Completed' -Message 'User exported.' -Data ([ordered]@{
                        UserId                         = Get-TrimmedValue -Value $user.Id
                        UserPrincipalName              = Get-TrimmedValue -Value $user.UserPrincipalName
                        DisplayName                    = Get-TrimmedValue -Value $user.DisplayName
                        GivenName                      = Get-TrimmedValue -Value $user.GivenName
                        Surname                        = Get-TrimmedValue -Value $user.Surname
                        MailNickname                   = Get-TrimmedValue -Value $user.MailNickname
                        UserType                       = Get-TrimmedValue -Value $user.UserType
                        AccountEnabled                 = [string]$user.AccountEnabled
                        UsageLocation                  = Get-TrimmedValue -Value $user.UsageLocation
                        PreferredLanguage              = Get-TrimmedValue -Value $user.PreferredLanguage
                        Department                     = Get-TrimmedValue -Value $user.Department
                        JobTitle                       = Get-TrimmedValue -Value $user.JobTitle
                        CompanyName                    = Get-TrimmedValue -Value $user.CompanyName
                        OfficeLocation                 = Get-TrimmedValue -Value $user.OfficeLocation
                        EmployeeId                     = Get-TrimmedValue -Value $user.EmployeeId
                        EmployeeType                   = Get-TrimmedValue -Value $user.EmployeeType
                        EmployeeHireDate               = Get-TrimmedValue -Value $user.EmployeeHireDate
                        MobilePhone                    = Get-TrimmedValue -Value $user.MobilePhone
                        BusinessPhones                 = Convert-MultiValueToString -Value $user.BusinessPhones
                        FaxNumber                      = Get-TrimmedValue -Value $user.FaxNumber
                        OtherMails                     = Convert-MultiValueToString -Value $user.OtherMails
                        StreetAddress                  = Get-TrimmedValue -Value $user.StreetAddress
                        City                           = Get-TrimmedValue -Value $user.City
                        State                          = Get-TrimmedValue -Value $user.State
                        PostalCode                     = Get-TrimmedValue -Value $user.PostalCode
                        Country                        = Get-TrimmedValue -Value $user.Country
                        PasswordPolicies               = Get-TrimmedValue -Value $user.PasswordPolicies
                        ExtensionAttribute1            = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute1'
                        ExtensionAttribute2            = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute2'
                        ExtensionAttribute3            = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute3'
                        ExtensionAttribute4            = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute4'
                        ExtensionAttribute5            = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute5'
                        ExtensionAttribute6            = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute6'
                        ExtensionAttribute7            = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute7'
                        ExtensionAttribute8            = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute8'
                        ExtensionAttribute9            = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute9'
                        ExtensionAttribute10           = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute10'
                        ExtensionAttribute11           = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute11'
                        ExtensionAttribute12           = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute12'
                        ExtensionAttribute13           = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute13'
                        ExtensionAttribute14           = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute14'
                        ExtensionAttribute15           = Get-UserExtensionAttributeValue -User $user -AttributeName 'extensionAttribute15'
                        Mail                           = Get-TrimmedValue -Value $user.Mail
                        CreatedDateTime                = [string]$user.CreatedDateTime
                        LastPasswordChangeDateTime     = [string]$user.LastPasswordChangeDateTime
                        OnPremisesSyncEnabled          = [string]$user.OnPremisesSyncEnabled
                        OnPremisesImmutableId          = Get-TrimmedValue -Value $user.OnPremisesImmutableId
                        OnPremisesDistinguishedName    = Get-TrimmedValue -Value $user.OnPremisesDistinguishedName
                        OnPremisesDomainName           = Get-TrimmedValue -Value $user.OnPremisesDomainName
                        OnPremisesSamAccountName       = Get-TrimmedValue -Value $user.OnPremisesSamAccountName
                        OnPremisesSecurityIdentifier   = Get-TrimmedValue -Value $user.OnPremisesSecurityIdentifier
                        OnPremisesLastSyncDateTime     = [string]$user.OnPremisesLastSyncDateTime
                        AssignedLicenseSkuIds          = Get-AssignmentSkuIds -AssignedLicenses @($user.AssignedLicenses)
                        AssignedPlanServiceIds         = Get-AssignedPlanNames -AssignedPlans @($user.AssignedPlans)
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($userPrincipalName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'GetEntraUser' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    UserId                         = ''
                    UserPrincipalName              = $userPrincipalName
                    DisplayName                    = ''
                    GivenName                      = ''
                    Surname                        = ''
                    MailNickname                   = ''
                    UserType                       = ''
                    AccountEnabled                 = ''
                    UsageLocation                  = ''
                    PreferredLanguage              = ''
                    Department                     = ''
                    JobTitle                       = ''
                    CompanyName                    = ''
                    OfficeLocation                 = ''
                    EmployeeId                     = ''
                    EmployeeType                   = ''
                    EmployeeHireDate               = ''
                    MobilePhone                    = ''
                    BusinessPhones                 = ''
                    FaxNumber                      = ''
                    OtherMails                     = ''
                    StreetAddress                  = ''
                    City                           = ''
                    State                          = ''
                    PostalCode                     = ''
                    Country                        = ''
                    PasswordPolicies               = ''
                    ExtensionAttribute1            = ''
                    ExtensionAttribute2            = ''
                    ExtensionAttribute3            = ''
                    ExtensionAttribute4            = ''
                    ExtensionAttribute5            = ''
                    ExtensionAttribute6            = ''
                    ExtensionAttribute7            = ''
                    ExtensionAttribute8            = ''
                    ExtensionAttribute9            = ''
                    ExtensionAttribute10           = ''
                    ExtensionAttribute11           = ''
                    ExtensionAttribute12           = ''
                    ExtensionAttribute13           = ''
                    ExtensionAttribute14           = ''
                    ExtensionAttribute15           = ''
                    Mail                           = ''
                    CreatedDateTime                = ''
                    LastPasswordChangeDateTime     = ''
                    OnPremisesSyncEnabled          = ''
                    OnPremisesImmutableId          = ''
                    OnPremisesDistinguishedName    = ''
                    OnPremisesDomainName           = ''
                    OnPremisesSamAccountName       = ''
                    OnPremisesSecurityIdentifier   = ''
                    OnPremisesLastSyncDateTime     = ''
                    AssignedLicenseSkuIds          = ''
                    AssignedPlanServiceIds         = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID user inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}




