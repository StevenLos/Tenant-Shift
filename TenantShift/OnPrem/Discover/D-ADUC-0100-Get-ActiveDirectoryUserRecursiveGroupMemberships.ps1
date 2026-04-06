<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-141500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
ActiveDirectory

.MODULEVERSIONPOLICY
Best-effort PSGallery check (offline-safe)

.SYNOPSIS
    Gets ActiveDirectoryUserRecursiveGroupMemberships and exports results to CSV.

.DESCRIPTION
    Gets ActiveDirectoryUserRecursiveGroupMemberships from Active Directory and writes the results to a CSV file.
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
    .\SM-D0010-Get-ActiveDirectoryUserRecursiveGroupMemberships.ps1 -InputCsvPath .\0010.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D0010-Get-ActiveDirectoryUserRecursiveGroupMemberships.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_SM-D0010-Get-ActiveDirectoryUserRecursiveGroupMemberships_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
                Filter      = '*'
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

function Resolve-PrimaryGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$User,

        [Parameter(Mandatory)]
        [string[]]$GroupProperties,

        [AllowEmptyString()]
        [string]$SearchBase,

        [AllowEmptyString()]
        [string]$Server
    )

    $primaryGroupId = Get-TrimmedValue -Value $User.PrimaryGroupID
    $sidValue = Get-TrimmedValue -Value $User.SID

    if ([string]::IsNullOrWhiteSpace($primaryGroupId) -or [string]::IsNullOrWhiteSpace($sidValue)) {
        return $null
    }

    $lastDash = $sidValue.LastIndexOf('-')
    if ($lastDash -le 0) {
        return $null
    }

    $domainSid = $sidValue.Substring(0, $lastDash)
    $primaryGroupSid = "{0}-{1}" -f $domainSid, $primaryGroupId
    $escapedPrimarySid = Escape-AdFilterValue -Value $primaryGroupSid

    $params = @{
        Filter      = "SID -eq '$escapedPrimarySid'"
        Properties  = $GroupProperties
        ErrorAction = 'Stop'
    }

    if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
        $params['SearchBase'] = $SearchBase
    }
    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $params['Server'] = $Server
    }

    return Get-ADGroup @params | Select-Object -First 1
}

function Get-RecursiveGroupMembership {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$User,

        [Parameter(Mandatory)]
        [string[]]$GroupProperties,

        [AllowEmptyString()]
        [string]$SearchBase,

        [AllowEmptyString()]
        [string]$Server
    )

    $userDn = Get-TrimmedValue -Value $User.DistinguishedName
    if ([string]::IsNullOrWhiteSpace($userDn)) {
        return [PSCustomObject]@{
            Groups          = @()
            PrimaryGroupDn  = ''
        }
    }

    $escapedDn = Escape-AdFilterValue -Value $userDn

    $groupParams = @{
        Filter      = "member -RecursiveMatch '$escapedDn'"
        Properties  = $GroupProperties
        ErrorAction = 'Stop'
    }

    if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
        $groupParams['SearchBase'] = $SearchBase
    }
    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $groupParams['Server'] = $Server
    }

    $groups = @(Get-ADGroup @groupParams)
    $primaryGroup = Resolve-PrimaryGroup -User $User -GroupProperties $GroupProperties -SearchBase $SearchBase -Server $Server
    $primaryGroupDn = ''

    if ($primaryGroup) {
        $primaryGroupDn = Get-TrimmedValue -Value $primaryGroup.DistinguishedName
        if (-not [string]::IsNullOrWhiteSpace($primaryGroupDn)) {
            $alreadyIncluded = @($groups | Where-Object { (Get-TrimmedValue -Value $_.DistinguishedName) -eq $primaryGroupDn }).Count -gt 0
            if (-not $alreadyIncluded) {
                $groups += $primaryGroup
            }
        }
    }

    return [PSCustomObject]@{
        Groups         = @($groups)
        PrimaryGroupDn = $primaryGroupDn
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

$requiredHeaders = @(
    'IdentityType',
    'IdentityValue'
)

$userPropertyNames = @(
    'DistinguishedName',
    'ObjectGuid',
    'SID',
    'SamAccountName',
    'UserPrincipalName',
    'DisplayName',
    'Enabled',
    'PrimaryGroupID'
)

$groupPropertyNames = @(
    'DistinguishedName',
    'ObjectGuid',
    'SamAccountName',
    'Name',
    'GroupCategory',
    'GroupScope',
    'SID'
)

Write-Status -Message 'Starting Active Directory recursive group membership inventory script.'
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

    Write-Status -Message "DiscoverAll enabled for recursive group membership. SearchBase='$resolvedSearchBase'." -Level WARN
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
            Resolve-UsersByScope -IdentityType $identityType -IdentityValue $identityValue -PropertyNames $userPropertyNames -SearchBase $effectiveSearchBase -Server $resolvedServer
        }

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $users.Count -gt $MaxObjects) {
            $users = @($users | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($users.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryUserRecursiveGroupMembership' -Status 'NotFound' -Message 'No matching users were found.' -Data ([ordered]@{
                            IdentityTypeRequested      = $identityType
                            IdentityValueRequested     = $identityValue
                            UserDistinguishedName      = ''
                            UserObjectGuid             = ''
                            UserSid                    = ''
                            UserSamAccountName         = ''
                            UserUserPrincipalName      = ''
                            UserDisplayName            = ''
                            UserEnabled                = ''
                            UserPrimaryGroupId         = ''
                            GroupDistinguishedName     = ''
                            GroupObjectGuid            = ''
                            GroupSamAccountName        = ''
                            GroupName                  = ''
                            GroupCategory              = ''
                            GroupScope                 = ''
                            IsPrimaryGroupMembership   = ''
                        })))

            $rowNumber++
            continue
        }

        foreach ($user in @($users | Sort-Object -Property UserPrincipalName, SamAccountName, DistinguishedName)) {
            $userUpn = Get-TrimmedValue -Value $user.UserPrincipalName
            $userSam = Get-TrimmedValue -Value $user.SamAccountName
            $userDn = Get-TrimmedValue -Value $user.DistinguishedName
            $userKey = if (-not [string]::IsNullOrWhiteSpace($userUpn)) { $userUpn } elseif (-not [string]::IsNullOrWhiteSpace($userSam)) { $userSam } else { $userDn }

            $membership = Invoke-WithRetry -OperationName "Load recursive groups for user $userKey" -ScriptBlock {
                Get-RecursiveGroupMembership -User $user -GroupProperties $groupPropertyNames -SearchBase $effectiveSearchBase -Server $resolvedServer
            }

            $groups = @($membership.Groups)
            $primaryGroupDn = Get-TrimmedValue -Value $membership.PrimaryGroupDn

            if ($groups.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userKey -Action 'GetActiveDirectoryUserRecursiveGroupMembership' -Status 'Completed' -Message 'User has no recursive group memberships.' -Data ([ordered]@{
                                IdentityTypeRequested      = $identityType
                                IdentityValueRequested     = $identityValue
                                UserDistinguishedName      = $userDn
                                UserObjectGuid             = Get-TrimmedValue -Value $user.ObjectGuid
                                UserSid                    = Get-TrimmedValue -Value $user.SID
                                UserSamAccountName         = $userSam
                                UserUserPrincipalName      = $userUpn
                                UserDisplayName            = Get-TrimmedValue -Value $user.DisplayName
                                UserEnabled                = [string]$user.Enabled
                                UserPrimaryGroupId         = Get-TrimmedValue -Value $user.PrimaryGroupID
                                GroupDistinguishedName     = ''
                                GroupObjectGuid            = ''
                                GroupSamAccountName        = ''
                                GroupName                  = ''
                                GroupCategory              = ''
                                GroupScope                 = ''
                                IsPrimaryGroupMembership   = ''
                            })))

                continue
            }

            foreach ($group in @($groups | Sort-Object -Property SamAccountName, Name, DistinguishedName)) {
                $groupDn = Get-TrimmedValue -Value $group.DistinguishedName
                $resultKey = if (-not [string]::IsNullOrWhiteSpace($groupDn)) { "$userKey|$groupDn" } else { "$userKey|$($group.Name)" }

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $resultKey -Action 'GetActiveDirectoryUserRecursiveGroupMembership' -Status 'Completed' -Message 'Recursive group membership exported.' -Data ([ordered]@{
                                IdentityTypeRequested      = $identityType
                                IdentityValueRequested     = $identityValue
                                UserDistinguishedName      = $userDn
                                UserObjectGuid             = Get-TrimmedValue -Value $user.ObjectGuid
                                UserSid                    = Get-TrimmedValue -Value $user.SID
                                UserSamAccountName         = $userSam
                                UserUserPrincipalName      = $userUpn
                                UserDisplayName            = Get-TrimmedValue -Value $user.DisplayName
                                UserEnabled                = [string]$user.Enabled
                                UserPrimaryGroupId         = Get-TrimmedValue -Value $user.PrimaryGroupID
                                GroupDistinguishedName     = $groupDn
                                GroupObjectGuid            = Get-TrimmedValue -Value $group.ObjectGuid
                                GroupSamAccountName        = Get-TrimmedValue -Value $group.SamAccountName
                                GroupName                  = Get-TrimmedValue -Value $group.Name
                                GroupCategory              = Get-TrimmedValue -Value $group.GroupCategory
                                GroupScope                 = Get-TrimmedValue -Value $group.GroupScope
                                IsPrimaryGroupMembership   = [string]($groupDn -eq $primaryGroupDn)
                            })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryUserRecursiveGroupMembership' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                        IdentityTypeRequested      = $identityType
                        IdentityValueRequested     = $identityValue
                        UserDistinguishedName      = ''
                        UserObjectGuid             = ''
                        UserSid                    = ''
                        UserSamAccountName         = ''
                        UserUserPrincipalName      = ''
                        UserDisplayName            = ''
                        UserEnabled                = ''
                        UserPrimaryGroupId         = ''
                        GroupDistinguishedName     = ''
                        GroupObjectGuid            = ''
                        GroupSamAccountName        = ''
                        GroupName                  = ''
                        GroupCategory              = ''
                        GroupScope                 = ''
                        IsPrimaryGroupMembership   = ''
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
Write-Status -Message 'Active Directory recursive group membership inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
