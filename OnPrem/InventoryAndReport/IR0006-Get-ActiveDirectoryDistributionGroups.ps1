<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260302-201500

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR0006-Get-ActiveDirectoryDistributionGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$expectedGroupCategory = 'Distribution'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Resolve-GroupsByScope {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IdentityType,

        [Parameter(Mandatory)]
        [string]$IdentityValue,

        [Parameter(Mandatory)]
        [string]$GroupCategory,

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
                Filter      = "GroupCategory -eq '$GroupCategory'"
                Properties  = $PropertyNames
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADGroup @params)
        }
        'samaccountname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            $params = @{
                Filter      = "GroupCategory -eq '$GroupCategory' -and SamAccountName -eq '$escaped'"
                Properties  = $PropertyNames
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADGroup @params)
        }
        'name' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            $params = @{
                Filter      = "GroupCategory -eq '$GroupCategory' -and Name -eq '$escaped'"
                Properties  = $PropertyNames
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADGroup @params)
        }
        'distinguishedname' {
            $params = @{
                Identity    = $IdentityValue
                Properties  = $PropertyNames
                ErrorAction = 'SilentlyContinue'
            }

            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            $group = Get-ADGroup @params
            if ($group -and (Get-TrimmedValue -Value $group.GroupCategory) -eq $GroupCategory) {
                return @($group)
            }

            return @()
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            $params = @{
                Identity    = $guid
                Properties  = $PropertyNames
                ErrorAction = 'SilentlyContinue'
            }

            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            $group = Get-ADGroup @params
            if ($group -and (Get-TrimmedValue -Value $group.GroupCategory) -eq $GroupCategory) {
                return @($group)
            }

            return @()
        }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use All, SamAccountName, Name, DistinguishedName, or ObjectGuid."
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
    'Name',
    'SamAccountName',
    'DisplayName',
    'Description',
    'GroupCategory',
    'GroupScope',
    'ManagedBy',
    'mail',
    'mailNickname',
    'proxyAddresses',
    'msExchHideFromAddressLists',
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
    'Member'
)

Write-Status -Message 'Starting Active Directory distribution group inventory script.'
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

    Write-Status -Message "DiscoverAll enabled for Active Directory groups. SearchBase='$resolvedSearchBase'." -Level WARN
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
        $groups = Invoke-WithRetry -OperationName "Load groups for $primaryKey" -ScriptBlock {
            Resolve-GroupsByScope -IdentityType $identityType -IdentityValue $identityValue -GroupCategory $expectedGroupCategory -PropertyNames $propertyNames -SearchBase $effectiveSearchBase -Server $resolvedServer
        }

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $groups.Count -gt $MaxObjects) {
            $groups = @($groups | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryDistributionGroup' -Status 'NotFound' -Message 'No matching groups were found.' -Data ([ordered]@{
                            IdentityTypeRequested = $identityType
                            IdentityValueRequested = $identityValue
                            DistinguishedName = ''
                            ObjectGuid = ''
                            SID = ''
                            Name = ''
                            SamAccountName = ''
                            DisplayName = ''
                            Description = ''
                            GroupCategory = ''
                            GroupScope = ''
                            ManagedBy = ''
                            Mail = ''
                            MailNickname = ''
                            ProxyAddresses = ''
                            HideFromAddressLists = ''
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
                            MemberCount = ''
                        })))

            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property SamAccountName, Name, DistinguishedName)) {
            $groupKey = if (-not [string]::IsNullOrWhiteSpace((Get-TrimmedValue -Value $group.SamAccountName))) { (Get-TrimmedValue -Value $group.SamAccountName) } else { (Get-TrimmedValue -Value $group.Name) }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupKey -Action 'GetActiveDirectoryDistributionGroup' -Status 'Completed' -Message 'Group exported.' -Data ([ordered]@{
                            IdentityTypeRequested = $identityType
                            IdentityValueRequested = $identityValue
                            DistinguishedName = Get-TrimmedValue -Value $group.DistinguishedName
                            ObjectGuid = Get-TrimmedValue -Value $group.ObjectGuid
                            SID = Get-TrimmedValue -Value $group.SID
                            Name = Get-TrimmedValue -Value $group.Name
                            SamAccountName = Get-TrimmedValue -Value $group.SamAccountName
                            DisplayName = Get-TrimmedValue -Value $group.DisplayName
                            Description = Get-TrimmedValue -Value $group.Description
                            GroupCategory = Get-TrimmedValue -Value $group.GroupCategory
                            GroupScope = Get-TrimmedValue -Value $group.GroupScope
                            ManagedBy = Get-TrimmedValue -Value $group.ManagedBy
                            Mail = Get-TrimmedValue -Value $group.mail
                            MailNickname = Get-TrimmedValue -Value $group.mailNickname
                            ProxyAddresses = Convert-MultiValueToString -Value $group.proxyAddresses
                            HideFromAddressLists = [string]$group.msExchHideFromAddressLists
                            ExtensionAttribute1 = Get-TrimmedValue -Value $group.extensionAttribute1
                            ExtensionAttribute2 = Get-TrimmedValue -Value $group.extensionAttribute2
                            ExtensionAttribute3 = Get-TrimmedValue -Value $group.extensionAttribute3
                            ExtensionAttribute4 = Get-TrimmedValue -Value $group.extensionAttribute4
                            ExtensionAttribute5 = Get-TrimmedValue -Value $group.extensionAttribute5
                            ExtensionAttribute6 = Get-TrimmedValue -Value $group.extensionAttribute6
                            ExtensionAttribute7 = Get-TrimmedValue -Value $group.extensionAttribute7
                            ExtensionAttribute8 = Get-TrimmedValue -Value $group.extensionAttribute8
                            ExtensionAttribute9 = Get-TrimmedValue -Value $group.extensionAttribute9
                            ExtensionAttribute10 = Get-TrimmedValue -Value $group.extensionAttribute10
                            ExtensionAttribute11 = Get-TrimmedValue -Value $group.extensionAttribute11
                            ExtensionAttribute12 = Get-TrimmedValue -Value $group.extensionAttribute12
                            ExtensionAttribute13 = Get-TrimmedValue -Value $group.extensionAttribute13
                            ExtensionAttribute14 = Get-TrimmedValue -Value $group.extensionAttribute14
                            ExtensionAttribute15 = Get-TrimmedValue -Value $group.extensionAttribute15
                            CanonicalName = Get-TrimmedValue -Value $group.CanonicalName
                            WhenCreated = Get-TrimmedValue -Value $group.whenCreated
                            WhenChanged = Get-TrimmedValue -Value $group.whenChanged
                            MemberCount = [string]@($group.Member).Count
                        })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryDistributionGroup' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                        IdentityTypeRequested = $identityType
                        IdentityValueRequested = $identityValue
                        DistinguishedName = ''
                        ObjectGuid = ''
                        SID = ''
                        Name = ''
                        SamAccountName = ''
                        DisplayName = ''
                        Description = ''
                        GroupCategory = ''
                        GroupScope = ''
                        ManagedBy = ''
                        Mail = ''
                        MailNickname = ''
                        ProxyAddresses = ''
                        HideFromAddressLists = ''
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
                        MemberCount = ''
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
Write-Status -Message 'Active Directory distribution group inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
