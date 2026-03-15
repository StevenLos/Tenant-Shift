<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-120000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Groups

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR3006-Get-EntraDynamicUserSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-TrimmedValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    return ([string]$Value).Trim()
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

function Test-IsDynamicUserSecurityGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Group
    )

    $membershipRule = Get-TrimmedValue -Value $Group.MembershipRule
    $groupTypes = @($Group.GroupTypes)

    $isDynamic = ($groupTypes -contains 'DynamicMembership')
    $isSecurityEnabled = ($Group.SecurityEnabled -eq $true)
    $isMailDisabled = ($Group.MailEnabled -eq $false)
    $hasRule = -not [string]::IsNullOrWhiteSpace($membershipRule)
    $looksLikeUserRule = $membershipRule.ToLowerInvariant().Contains('user.')

    return ($isDynamic -and $isSecurityEnabled -and $isMailDisabled -and $hasRule -and $looksLikeUserRule)
}

$requiredHeaders = @(
    'GroupDisplayName'
)

Write-Status -Message 'Starting Entra ID dynamic user security group inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')
Ensure-GraphConnection -RequiredScopes @('Group.Read.All', 'Directory.Read.All')

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

$groupSelect = 'id,displayName,description,mail,mailNickname,proxyAddresses,securityEnabled,mailEnabled,membershipRule,membershipRuleProcessingState,groupTypes,visibility,classification,preferredDataLocation,isAssignableToRole,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesDomainName,onPremisesNetBiosName,onPremisesSamAccountName,onPremisesSecurityIdentifier,createdDateTime,renewedDateTime,deletedDateTime'
$allDynamicUserGroupsCache = $null

$rowNumber = 1
foreach ($row in $rows) {
    $groupDisplayName = ([string]$row.GroupDisplayName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($groupDisplayName)) {
            throw 'GroupDisplayName is required. Use * to inventory all dynamic user security groups.'
        }

        $groups = @()
        if ($groupDisplayName -eq '*') {
            if ($null -eq $allDynamicUserGroupsCache) {
                $allGroups = @(Invoke-WithRetry -OperationName 'Load all groups for dynamic user security group inventory' -ScriptBlock {
                    Get-MgGroup -All -Property $groupSelect -ErrorAction Stop
                })
                $allDynamicUserGroupsCache = @($allGroups | Where-Object { Test-IsDynamicUserSecurityGroup -Group $_ } | Sort-Object -Property DisplayName, Id)
            }

            $groups = @($allDynamicUserGroupsCache)
        }
        else {
            $escapedName = Escape-ODataString -Value $groupDisplayName
            $candidateGroups = @(Invoke-WithRetry -OperationName "Lookup group $groupDisplayName" -ScriptBlock {
                Get-MgGroup -Filter "displayName eq '$escapedName'" -ConsistencyLevel eventual -Property $groupSelect -ErrorAction Stop
            })

            $groups = @($candidateGroups | Where-Object { Test-IsDynamicUserSecurityGroup -Group $_ })
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'GetEntraDynamicUserSecurityGroup' -Status 'NotFound' -Message 'No matching dynamic user security groups were found.' -Data ([ordered]@{
                        GroupId                       = ''
                        GroupDisplayName              = $groupDisplayName
                        Description                   = ''
                        Mail                          = ''
                        MailNickname                  = ''
                        ProxyAddresses                = ''
                        SecurityEnabled               = ''
                        MailEnabled                   = ''
                        MembershipType                = ''
                        MembershipRule                = ''
                        MembershipRuleProcessingState = ''
                        GroupTypes                    = ''
                        Visibility                    = ''
                        Classification                = ''
                        PreferredDataLocation         = ''
                        IsAssignableToRole            = ''
                        OnPremisesSyncEnabled         = ''
                        OnPremisesLastSyncDateTime    = ''
                        OnPremisesDomainName          = ''
                        OnPremisesNetBiosName         = ''
                        OnPremisesSamAccountName      = ''
                        OnPremisesSecurityIdentifier  = ''
                        CreatedDateTime               = ''
                        RenewedDateTime               = ''
                        DeletedDateTime               = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Id)) {
            $groupId = Get-TrimmedValue -Value $group.Id
            $displayName = Get-TrimmedValue -Value $group.DisplayName
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$displayName|$groupId" -Action 'GetEntraDynamicUserSecurityGroup' -Status 'Completed' -Message 'Dynamic user security group exported.' -Data ([ordered]@{
                        GroupId                       = $groupId
                        GroupDisplayName              = $displayName
                        Description                   = Get-TrimmedValue -Value $group.Description
                        Mail                          = Get-TrimmedValue -Value $group.Mail
                        MailNickname                  = Get-TrimmedValue -Value $group.MailNickname
                        ProxyAddresses                = Convert-MultiValueToString -Value $group.ProxyAddresses
                        SecurityEnabled               = [string]$group.SecurityEnabled
                        MailEnabled                   = [string]$group.MailEnabled
                        MembershipType                = 'Dynamic'
                        MembershipRule                = Get-TrimmedValue -Value $group.MembershipRule
                        MembershipRuleProcessingState = Get-TrimmedValue -Value $group.MembershipRuleProcessingState
                        GroupTypes                    = Convert-MultiValueToString -Value $group.GroupTypes
                        Visibility                    = Get-TrimmedValue -Value $group.Visibility
                        Classification                = Get-TrimmedValue -Value $group.Classification
                        PreferredDataLocation         = Get-TrimmedValue -Value $group.PreferredDataLocation
                        IsAssignableToRole            = [string]$group.IsAssignableToRole
                        OnPremisesSyncEnabled         = [string]$group.OnPremisesSyncEnabled
                        OnPremisesLastSyncDateTime    = [string]$group.OnPremisesLastSyncDateTime
                        OnPremisesDomainName          = Get-TrimmedValue -Value $group.OnPremisesDomainName
                        OnPremisesNetBiosName         = Get-TrimmedValue -Value $group.OnPremisesNetBiosName
                        OnPremisesSamAccountName      = Get-TrimmedValue -Value $group.OnPremisesSamAccountName
                        OnPremisesSecurityIdentifier  = Get-TrimmedValue -Value $group.OnPremisesSecurityIdentifier
                        CreatedDateTime               = [string]$group.CreatedDateTime
                        RenewedDateTime               = [string]$group.RenewedDateTime
                        DeletedDateTime               = [string]$group.DeletedDateTime
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupDisplayName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'GetEntraDynamicUserSecurityGroup' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    GroupId                       = ''
                    GroupDisplayName              = $groupDisplayName
                    Description                   = ''
                    Mail                          = ''
                    MailNickname                  = ''
                    ProxyAddresses                = ''
                    SecurityEnabled               = ''
                    MailEnabled                   = ''
                    MembershipType                = ''
                    MembershipRule                = ''
                    MembershipRuleProcessingState = ''
                    GroupTypes                    = ''
                    Visibility                    = ''
                    Classification                = ''
                    PreferredDataLocation         = ''
                    IsAssignableToRole            = ''
                    OnPremisesSyncEnabled         = ''
                    OnPremisesLastSyncDateTime    = ''
                    OnPremisesDomainName          = ''
                    OnPremisesNetBiosName         = ''
                    OnPremisesSamAccountName      = ''
                    OnPremisesSecurityIdentifier  = ''
                    CreatedDateTime               = ''
                    RenewedDateTime               = ''
                    DeletedDateTime               = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID dynamic user security group inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





