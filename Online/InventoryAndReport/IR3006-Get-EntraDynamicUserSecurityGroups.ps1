<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-004416

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Groups

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3006-Get-EntraDynamicUserSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Test-IsDynamicUserSecurityGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Group
    )

    $membershipRule = ([string]$Group.MembershipRule).Trim()
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

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$groupSelect = 'id,displayName,description,mailNickname,securityEnabled,mailEnabled,membershipRule,membershipRuleProcessingState,groupTypes,visibility,createdDateTime'
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
                        MailNickname                  = ''
                        SecurityEnabled               = ''
                        MailEnabled                   = ''
                        MembershipType                = ''
                        MembershipRule                = ''
                        MembershipRuleProcessingState = ''
                        Visibility                    = ''
                        CreatedDateTime               = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Id)) {
            $groupId = ([string]$group.Id).Trim()
            $displayName = ([string]$group.DisplayName).Trim()
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$displayName|$groupId" -Action 'GetEntraDynamicUserSecurityGroup' -Status 'Completed' -Message 'Dynamic user security group exported.' -Data ([ordered]@{
                        GroupId                       = $groupId
                        GroupDisplayName              = $displayName
                        Description                   = ([string]$group.Description).Trim()
                        MailNickname                  = ([string]$group.MailNickname).Trim()
                        SecurityEnabled               = [string]$group.SecurityEnabled
                        MailEnabled                   = [string]$group.MailEnabled
                        MembershipType                = 'Dynamic'
                        MembershipRule                = ([string]$group.MembershipRule).Trim()
                        MembershipRuleProcessingState = ([string]$group.MembershipRuleProcessingState).Trim()
                        Visibility                    = ([string]$group.Visibility).Trim()
                        CreatedDateTime               = [string]$group.CreatedDateTime
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupDisplayName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'GetEntraDynamicUserSecurityGroup' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    GroupId                       = ''
                    GroupDisplayName              = $groupDisplayName
                    Description                   = ''
                    MailNickname                  = ''
                    SecurityEnabled               = ''
                    MailEnabled                   = ''
                    MembershipType                = ''
                    MembershipRule                = ''
                    MembershipRuleProcessingState = ''
                    Visibility                    = ''
                    CreatedDateTime               = ''
                })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID dynamic user security group inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







