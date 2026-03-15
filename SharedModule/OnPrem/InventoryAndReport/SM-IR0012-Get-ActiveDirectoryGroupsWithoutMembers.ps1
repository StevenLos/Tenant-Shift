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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR0012-Get-ActiveDirectoryGroupsWithoutMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

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

            return @(Get-ADGroup @params)
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

            return @(Get-ADGroup @params)
        }
        'name' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            $params = @{
                Filter      = "Name -eq '$escaped'"
                Properties  = '*'
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
                Properties  = '*'
                ErrorAction = 'SilentlyContinue'
            }

            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            $group = Get-ADGroup @params
            if ($group) {
                return @($group)
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

            $group = Get-ADGroup @params
            if ($group) {
                return @($group)
            }

            return @()
        }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use All, SamAccountName, Name, DistinguishedName, or ObjectGuid."
        }
    }
}

function Get-DirectMemberCount {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Group
    )

    $count = 0
    foreach ($memberDn in @($Group.Member)) {
        if (-not [string]::IsNullOrWhiteSpace((Get-TrimmedValue -Value $memberDn))) {
            $count++
        }
    }

    return $count
}

function Get-PrimaryGroupMemberCount {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Group,

        [AllowEmptyString()]
        [string]$SearchBase,

        [AllowEmptyString()]
        [string]$Server
    )

    $sidValue = Get-TrimmedValue -Value $Group.SID
    if ([string]::IsNullOrWhiteSpace($sidValue)) {
        return 0
    }

    $ridText = $sidValue.Substring($sidValue.LastIndexOf('-') + 1)
    $rid = 0
    if (-not [int]::TryParse($ridText, [ref]$rid)) {
        return 0
    }

    $params = @{
        Filter      = "PrimaryGroupID -eq $rid"
        ResultSetSize = 1
        ErrorAction = 'Stop'
    }

    if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
        $params['SearchBase'] = $SearchBase
    }
    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $params['Server'] = $Server
    }

    $users = @(Get-ADUser @params)
    return $users.Count
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

$groupPropertyNames = @(
    'DistinguishedName',
    'ObjectGuid',
    'SamAccountName',
    'Name',
    'GroupCategory',
    'GroupScope',
    'Member',
    'SID'
)

Write-Status -Message 'Starting Active Directory groups-without-members inventory script.'
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

    Write-Status -Message "DiscoverAll enabled for groups-without-members inventory. SearchBase='$resolvedSearchBase'." -Level WARN
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
            Resolve-GroupsByScope -IdentityType $identityType -IdentityValue $identityValue -PropertyNames $groupPropertyNames -SearchBase $effectiveSearchBase -Server $resolvedServer
        }

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $groups.Count -gt $MaxObjects) {
            $groups = @($groups | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryGroupsWithoutMembers' -Status 'NotFound' -Message 'No matching groups were found.' -Data ([ordered]@{
                            IdentityTypeRequested    = $identityType
                            IdentityValueRequested   = $identityValue
                            GroupDistinguishedName   = ''
                            GroupObjectGuid          = ''
                            GroupSamAccountName      = ''
                            GroupName                = ''
                            GroupCategory            = ''
                            GroupScope               = ''
                            DirectMemberCount        = ''
                            PrimaryGroupMemberCount  = ''
                        })))

            $rowNumber++
            continue
        }

        $groupsWithoutMembers = 0
        foreach ($group in @($groups | Sort-Object -Property SamAccountName, Name, DistinguishedName)) {
            $groupDn = Get-TrimmedValue -Value $group.DistinguishedName
            $groupSam = Get-TrimmedValue -Value $group.SamAccountName
            $groupName = Get-TrimmedValue -Value $group.Name
            $groupKey = if (-not [string]::IsNullOrWhiteSpace($groupSam)) { $groupSam } elseif (-not [string]::IsNullOrWhiteSpace($groupName)) { $groupName } else { $groupDn }

            $directMemberCount = Get-DirectMemberCount -Group $group
            if ($directMemberCount -gt 0) {
                continue
            }

            $primaryGroupMemberCount = Invoke-WithRetry -OperationName "Check primary-group users for group $groupKey" -ScriptBlock {
                Get-PrimaryGroupMemberCount -Group $group -SearchBase $effectiveSearchBase -Server $resolvedServer
            }

            if ($primaryGroupMemberCount -gt 0) {
                continue
            }

            $groupsWithoutMembers++
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupKey -Action 'GetActiveDirectoryGroupsWithoutMembers' -Status 'Completed' -Message 'Group has no direct members and no primary-group users.' -Data ([ordered]@{
                            IdentityTypeRequested    = $identityType
                            IdentityValueRequested   = $identityValue
                            GroupDistinguishedName   = $groupDn
                            GroupObjectGuid          = Get-TrimmedValue -Value $group.ObjectGuid
                            GroupSamAccountName      = $groupSam
                            GroupName                = $groupName
                            GroupCategory            = Get-TrimmedValue -Value $group.GroupCategory
                            GroupScope               = Get-TrimmedValue -Value $group.GroupScope
                            DirectMemberCount        = '0'
                            PrimaryGroupMemberCount  = '0'
                        })))
        }

        if ($groupsWithoutMembers -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryGroupsWithoutMembers' -Status 'Completed' -Message 'No groups without members were found for this scope.' -Data ([ordered]@{
                            IdentityTypeRequested    = $identityType
                            IdentityValueRequested   = $identityValue
                            GroupDistinguishedName   = ''
                            GroupObjectGuid          = ''
                            GroupSamAccountName      = ''
                            GroupName                = ''
                            GroupCategory            = ''
                            GroupScope               = ''
                            DirectMemberCount        = ''
                            PrimaryGroupMemberCount  = ''
                        })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryGroupsWithoutMembers' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                        IdentityTypeRequested    = $identityType
                        IdentityValueRequested   = $identityValue
                        GroupDistinguishedName   = ''
                        GroupObjectGuid          = ''
                        GroupSamAccountName      = ''
                        GroupName                = ''
                        GroupCategory            = ''
                        GroupScope               = ''
                        DirectMemberCount        = ''
                        PrimaryGroupMemberCount  = ''
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
Write-Status -Message 'Active Directory groups-without-members inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
