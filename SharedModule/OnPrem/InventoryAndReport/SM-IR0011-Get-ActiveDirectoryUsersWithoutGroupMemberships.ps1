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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR0011-Get-ActiveDirectoryUsersWithoutGroupMemberships_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-RecursiveGroupMembershipCount {
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
        return 0
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

    if ($primaryGroup) {
        $primaryGroupDn = Get-TrimmedValue -Value $primaryGroup.DistinguishedName
        if (-not [string]::IsNullOrWhiteSpace($primaryGroupDn)) {
            $alreadyIncluded = @($groups | Where-Object { (Get-TrimmedValue -Value $_.DistinguishedName) -eq $primaryGroupDn }).Count -gt 0
            if (-not $alreadyIncluded) {
                $groups += $primaryGroup
            }
        }
    }

    return @($groups).Count
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

Write-Status -Message 'Starting Active Directory users-without-memberships inventory script.'
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

    Write-Status -Message "DiscoverAll enabled for users-without-memberships inventory. SearchBase='$resolvedSearchBase'." -Level WARN
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
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryUsersWithoutGroupMemberships' -Status 'NotFound' -Message 'No matching users were found.' -Data ([ordered]@{
                            IdentityTypeRequested            = $identityType
                            IdentityValueRequested           = $identityValue
                            UserDistinguishedName            = ''
                            UserObjectGuid                   = ''
                            UserSid                          = ''
                            UserSamAccountName               = ''
                            UserUserPrincipalName            = ''
                            UserDisplayName                  = ''
                            UserEnabled                      = ''
                            UserPrimaryGroupId               = ''
                            RecursiveGroupMembershipCount    = ''
                        })))

            $rowNumber++
            continue
        }

        $usersWithoutMembership = 0
        foreach ($user in @($users | Sort-Object -Property UserPrincipalName, SamAccountName, DistinguishedName)) {
            $userUpn = Get-TrimmedValue -Value $user.UserPrincipalName
            $userSam = Get-TrimmedValue -Value $user.SamAccountName
            $userDn = Get-TrimmedValue -Value $user.DistinguishedName
            $userKey = if (-not [string]::IsNullOrWhiteSpace($userUpn)) { $userUpn } elseif (-not [string]::IsNullOrWhiteSpace($userSam)) { $userSam } else { $userDn }

            $membershipCount = Invoke-WithRetry -OperationName "Load recursive group count for user $userKey" -ScriptBlock {
                Get-RecursiveGroupMembershipCount -User $user -GroupProperties $groupPropertyNames -SearchBase $effectiveSearchBase -Server $resolvedServer
            }

            if ($membershipCount -eq 0) {
                $usersWithoutMembership++
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userKey -Action 'GetActiveDirectoryUsersWithoutGroupMemberships' -Status 'Completed' -Message 'User has no recursive group memberships.' -Data ([ordered]@{
                                IdentityTypeRequested            = $identityType
                                IdentityValueRequested           = $identityValue
                                UserDistinguishedName            = $userDn
                                UserObjectGuid                   = Get-TrimmedValue -Value $user.ObjectGuid
                                UserSid                          = Get-TrimmedValue -Value $user.SID
                                UserSamAccountName               = $userSam
                                UserUserPrincipalName            = $userUpn
                                UserDisplayName                  = Get-TrimmedValue -Value $user.DisplayName
                                UserEnabled                      = [string]$user.Enabled
                                UserPrimaryGroupId               = Get-TrimmedValue -Value $user.PrimaryGroupID
                                RecursiveGroupMembershipCount    = [string]$membershipCount
                            })))
            }
        }

        if ($usersWithoutMembership -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryUsersWithoutGroupMemberships' -Status 'Completed' -Message 'No users without recursive group memberships were found for this scope.' -Data ([ordered]@{
                            IdentityTypeRequested            = $identityType
                            IdentityValueRequested           = $identityValue
                            UserDistinguishedName            = ''
                            UserObjectGuid                   = ''
                            UserSid                          = ''
                            UserSamAccountName               = ''
                            UserUserPrincipalName            = ''
                            UserDisplayName                  = ''
                            UserEnabled                      = ''
                            UserPrimaryGroupId               = ''
                            RecursiveGroupMembershipCount    = ''
                        })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryUsersWithoutGroupMemberships' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                        IdentityTypeRequested            = $identityType
                        IdentityValueRequested           = $identityValue
                        UserDistinguishedName            = ''
                        UserObjectGuid                   = ''
                        UserSid                          = ''
                        UserSamAccountName               = ''
                        UserUserPrincipalName            = ''
                        UserDisplayName                  = ''
                        UserEnabled                      = ''
                        UserPrimaryGroupId               = ''
                        RecursiveGroupMembershipCount    = ''
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
Write-Status -Message 'Active Directory users-without-memberships inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
