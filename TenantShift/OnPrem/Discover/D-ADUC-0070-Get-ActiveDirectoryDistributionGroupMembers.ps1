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

.SYNOPSIS
    Gets ActiveDirectoryDistributionGroupMembers and exports results to CSV.

.DESCRIPTION
    Gets ActiveDirectoryDistributionGroupMembers from Active Directory and writes the results to a CSV file.
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
    .\SM-D0008-Get-ActiveDirectoryDistributionGroupMembers.ps1 -InputCsvPath .\0008.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D0008-Get-ActiveDirectoryDistributionGroupMembers.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_SM-D0008-Get-ActiveDirectoryDistributionGroupMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
                Properties  = 'GroupCategory', 'GroupScope', 'SamAccountName', 'Name', 'DistinguishedName', 'ObjectGuid'
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
                Properties  = 'GroupCategory', 'GroupScope', 'SamAccountName', 'Name', 'DistinguishedName', 'ObjectGuid'
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
                Properties  = 'GroupCategory', 'GroupScope', 'SamAccountName', 'Name', 'DistinguishedName', 'ObjectGuid'
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
                Properties  = 'GroupCategory', 'GroupScope', 'SamAccountName', 'Name', 'DistinguishedName', 'ObjectGuid'
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
                Properties  = 'GroupCategory', 'GroupScope', 'SamAccountName', 'Name', 'DistinguishedName', 'ObjectGuid'
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

function Get-MemberDetail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Member,

        [AllowEmptyString()]
        [string]$Server
    )

    $memberDn = Get-TrimmedValue -Value $Member.DistinguishedName
    if ([string]::IsNullOrWhiteSpace($memberDn)) {
        return [PSCustomObject]@{
            DistinguishedName = ''
            ObjectGuid = ''
            ObjectClass = Get-TrimmedValue -Value $Member.objectClass
            SamAccountName = ''
            UserPrincipalName = ''
            Name = Get-TrimmedValue -Value $Member.Name
            SID = Get-TrimmedValue -Value $Member.SID
        }
    }

    $params = @{
        Identity    = $memberDn
        Properties  = 'objectClass', 'samAccountName', 'userPrincipalName', 'name', 'objectGuid', 'sID'
        ErrorAction = 'SilentlyContinue'
    }

    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $params['Server'] = $Server
    }

    $detail = Get-ADObject @params
    if ($detail) {
        return [PSCustomObject]@{
            DistinguishedName = Get-TrimmedValue -Value $detail.DistinguishedName
            ObjectGuid = Get-TrimmedValue -Value $detail.ObjectGuid
            ObjectClass = Get-TrimmedValue -Value $detail.objectClass
            SamAccountName = Get-TrimmedValue -Value $detail.samAccountName
            UserPrincipalName = Get-TrimmedValue -Value $detail.userPrincipalName
            Name = Get-TrimmedValue -Value $detail.Name
            SID = Get-TrimmedValue -Value $detail.SID
        }
    }

    return [PSCustomObject]@{
        DistinguishedName = $memberDn
        ObjectGuid = Get-TrimmedValue -Value $Member.ObjectGuid
        ObjectClass = Get-TrimmedValue -Value $Member.objectClass
        SamAccountName = Get-TrimmedValue -Value $Member.SamAccountName
        UserPrincipalName = Get-TrimmedValue -Value $Member.UserPrincipalName
        Name = Get-TrimmedValue -Value $Member.Name
        SID = Get-TrimmedValue -Value $Member.SID
    }
}

$requiredHeaders = @(
    'IdentityType',
    'IdentityValue'
)

Write-Status -Message 'Starting Active Directory distribution group membership inventory script.'
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

    Write-Status -Message "DiscoverAll enabled for Active Directory distribution group members. SearchBase='$resolvedSearchBase'." -Level WARN
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
            Resolve-GroupsByScope -IdentityType $identityType -IdentityValue $identityValue -GroupCategory $expectedGroupCategory -SearchBase $effectiveSearchBase -Server $resolvedServer
        }

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $groups.Count -gt $MaxObjects) {
            $groups = @($groups | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryDistributionGroupMembers' -Status 'NotFound' -Message 'No matching groups were found.' -Data ([ordered]@{
                            IdentityTypeRequested = $identityType
                            IdentityValueRequested = $identityValue
                            GroupDistinguishedName = ''
                            GroupSamAccountName = ''
                            GroupName = ''
                            GroupCategory = ''
                            GroupScope = ''
                            MemberDistinguishedName = ''
                            MemberObjectGuid = ''
                            MemberObjectClass = ''
                            MemberSamAccountName = ''
                            MemberUserPrincipalName = ''
                            MemberName = ''
                            MemberSid = ''
                        })))

            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property SamAccountName, Name, DistinguishedName)) {
            $groupKey = if (-not [string]::IsNullOrWhiteSpace((Get-TrimmedValue -Value $group.SamAccountName))) { (Get-TrimmedValue -Value $group.SamAccountName) } else { (Get-TrimmedValue -Value $group.Name) }

            $members = Invoke-WithRetry -OperationName "Load members for group $groupKey" -ScriptBlock {
                $memberParams = @{
                    Identity    = $group.ObjectGuid
                    ErrorAction = 'Stop'
                }

                if (-not [string]::IsNullOrWhiteSpace($resolvedServer)) {
                    $memberParams['Server'] = $resolvedServer
                }

                @(Get-ADGroupMember @memberParams)
            }

            if ($members.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupKey -Action 'GetActiveDirectoryDistributionGroupMembers' -Status 'Completed' -Message 'Group has no direct members.' -Data ([ordered]@{
                                IdentityTypeRequested = $identityType
                                IdentityValueRequested = $identityValue
                                GroupDistinguishedName = Get-TrimmedValue -Value $group.DistinguishedName
                                GroupSamAccountName = Get-TrimmedValue -Value $group.SamAccountName
                                GroupName = Get-TrimmedValue -Value $group.Name
                                GroupCategory = Get-TrimmedValue -Value $group.GroupCategory
                                GroupScope = Get-TrimmedValue -Value $group.GroupScope
                                MemberDistinguishedName = ''
                                MemberObjectGuid = ''
                                MemberObjectClass = ''
                                MemberSamAccountName = ''
                                MemberUserPrincipalName = ''
                                MemberName = ''
                                MemberSid = ''
                            })))

                continue
            }

            foreach ($member in @($members | Sort-Object -Property objectClass, Name, DistinguishedName)) {
                $memberDetail = Get-MemberDetail -Member $member -Server $resolvedServer
                $memberDn = Get-TrimmedValue -Value $memberDetail.DistinguishedName
                $resultKey = if (-not [string]::IsNullOrWhiteSpace($memberDn)) { "$groupKey|$memberDn" } else { "$groupKey|$($memberDetail.Name)" }

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $resultKey -Action 'GetActiveDirectoryDistributionGroupMembers' -Status 'Completed' -Message 'Group member exported.' -Data ([ordered]@{
                                IdentityTypeRequested = $identityType
                                IdentityValueRequested = $identityValue
                                GroupDistinguishedName = Get-TrimmedValue -Value $group.DistinguishedName
                                GroupSamAccountName = Get-TrimmedValue -Value $group.SamAccountName
                                GroupName = Get-TrimmedValue -Value $group.Name
                                GroupCategory = Get-TrimmedValue -Value $group.GroupCategory
                                GroupScope = Get-TrimmedValue -Value $group.GroupScope
                                MemberDistinguishedName = $memberDn
                                MemberObjectGuid = Get-TrimmedValue -Value $memberDetail.ObjectGuid
                                MemberObjectClass = Get-TrimmedValue -Value $memberDetail.ObjectClass
                                MemberSamAccountName = Get-TrimmedValue -Value $memberDetail.SamAccountName
                                MemberUserPrincipalName = Get-TrimmedValue -Value $memberDetail.UserPrincipalName
                                MemberName = Get-TrimmedValue -Value $memberDetail.Name
                                MemberSid = Get-TrimmedValue -Value $memberDetail.SID
                            })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryDistributionGroupMembers' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                        IdentityTypeRequested = $identityType
                        IdentityValueRequested = $identityValue
                        GroupDistinguishedName = ''
                        GroupSamAccountName = ''
                        GroupName = ''
                        GroupCategory = ''
                        GroupScope = ''
                        MemberDistinguishedName = ''
                        MemberObjectGuid = ''
                        MemberObjectClass = ''
                        MemberSamAccountName = ''
                        MemberUserPrincipalName = ''
                        MemberName = ''
                        MemberSid = ''
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
Write-Status -Message 'Active Directory distribution group membership inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
