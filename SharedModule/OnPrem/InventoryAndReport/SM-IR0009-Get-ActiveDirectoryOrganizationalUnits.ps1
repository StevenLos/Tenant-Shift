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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR0009-Get-ActiveDirectoryOrganizationalUnits_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Resolve-OusByScope {
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
                Filter = '*'
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADOrganizationalUnit @params)
        }
        'name' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            $params = @{
                Filter = "Name -eq '$escaped'"
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADOrganizationalUnit @params)
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

            $ou = Get-ADOrganizationalUnit @params
            if ($ou) {
                return @($ou)
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

            $ou = Get-ADOrganizationalUnit @params
            if ($ou) {
                return @($ou)
            }

            return @()
        }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use All, Name, DistinguishedName, or ObjectGuid."
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
    'Name',
    'Description',
    'DisplayName',
    'ManagedBy',
    'StreetAddress',
    'City',
    'State',
    'PostalCode',
    'Country',
    'co',
    'countryCode',
    'CanonicalName',
    'whenCreated',
    'whenChanged',
    'ProtectedFromAccidentalDeletion',
    'gPLink',
    'gPOptions',
    'seeAlso'
)

function New-EmptyOuData {
    return [ordered]@{
        IdentityTypeRequested = ''
        IdentityValueRequested = ''
        DistinguishedName = ''
        ObjectGuid = ''
        Name = ''
        Description = ''
        DisplayName = ''
        ManagedBy = ''
        StreetAddress = ''
        City = ''
        StateOrProvince = ''
        PostalCode = ''
        CountryCode = ''
        CountryName = ''
        CountryNumericCode = ''
        CanonicalName = ''
        WhenCreated = ''
        WhenChanged = ''
        ProtectedFromAccidentalDeletion = ''
        GpLink = ''
        GpOptions = ''
        SeeAlso = ''
    }
}

Write-Status -Message 'Starting Active Directory organizational unit inventory script.'
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

    Write-Status -Message "DiscoverAll enabled for Active Directory organizational units. SearchBase='$resolvedSearchBase'." -Level WARN
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
        $ous = Invoke-WithRetry -OperationName "Load organizational units for $primaryKey" -ScriptBlock {
            Resolve-OusByScope -IdentityType $identityType -IdentityValue $identityValue -PropertyNames $propertyNames -SearchBase $effectiveSearchBase -Server $resolvedServer
        }

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $ous.Count -gt $MaxObjects) {
            $ous = @($ous | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($ous.Count -eq 0) {
            $data = New-EmptyOuData
            $data['IdentityTypeRequested'] = $identityType
            $data['IdentityValueRequested'] = $identityValue
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryOrganizationalUnit' -Status 'NotFound' -Message 'No matching organizational units were found.' -Data $data))

            $rowNumber++
            continue
        }

        foreach ($ou in @($ous | Sort-Object -Property CanonicalName, DistinguishedName)) {
            $ouPrimaryKey = Get-TrimmedValue -Value $ou.DistinguishedName
            if ([string]::IsNullOrWhiteSpace($ouPrimaryKey)) {
                $ouPrimaryKey = Get-TrimmedValue -Value $ou.Name
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $ouPrimaryKey -Action 'GetActiveDirectoryOrganizationalUnit' -Status 'Completed' -Message 'Organizational unit exported.' -Data ([ordered]@{
                            IdentityTypeRequested = $identityType
                            IdentityValueRequested = $identityValue
                            DistinguishedName = Get-TrimmedValue -Value $ou.DistinguishedName
                            ObjectGuid = Get-TrimmedValue -Value $ou.ObjectGuid
                            Name = Get-TrimmedValue -Value $ou.Name
                            Description = Get-TrimmedValue -Value $ou.Description
                            DisplayName = Get-TrimmedValue -Value $ou.DisplayName
                            ManagedBy = Get-TrimmedValue -Value $ou.ManagedBy
                            StreetAddress = Get-TrimmedValue -Value $ou.StreetAddress
                            City = Get-TrimmedValue -Value $ou.City
                            StateOrProvince = Get-TrimmedValue -Value $ou.State
                            PostalCode = Get-TrimmedValue -Value $ou.PostalCode
                            CountryCode = Get-TrimmedValue -Value $ou.Country
                            CountryName = Get-TrimmedValue -Value $ou.co
                            CountryNumericCode = Get-TrimmedValue -Value $ou.countryCode
                            CanonicalName = Get-TrimmedValue -Value $ou.CanonicalName
                            WhenCreated = Get-TrimmedValue -Value $ou.whenCreated
                            WhenChanged = Get-TrimmedValue -Value $ou.whenChanged
                            ProtectedFromAccidentalDeletion = [string]$ou.ProtectedFromAccidentalDeletion
                            GpLink = Get-TrimmedValue -Value $ou.gPLink
                            GpOptions = Get-TrimmedValue -Value $ou.gPOptions
                            SeeAlso = Convert-MultiValueToString -Value $ou.seeAlso
                        })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $data = New-EmptyOuData
        $data['IdentityTypeRequested'] = $identityType
        $data['IdentityValueRequested'] = $identityValue
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryOrganizationalUnit' -Status 'Failed' -Message $_.Exception.Message -Data $data))
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
Write-Status -Message 'Active Directory organizational unit inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
