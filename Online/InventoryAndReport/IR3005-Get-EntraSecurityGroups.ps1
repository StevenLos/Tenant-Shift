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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3005-Get-EntraSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Test-IsAssignedSecurityGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Group
    )

    $membershipRule = ([string]$Group.MembershipRule).Trim()
    $isAssigned = ($Group.SecurityEnabled -eq $true -and $Group.MailEnabled -eq $false -and [string]::IsNullOrWhiteSpace($membershipRule))
    return $isAssigned
}

$requiredHeaders = @(
    'GroupDisplayName'
)

Write-Status -Message 'Starting Entra ID assigned security group inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')
Ensure-GraphConnection -RequiredScopes @('Group.Read.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$groupSelect = 'id,displayName,description,mailNickname,securityEnabled,mailEnabled,membershipRule,membershipRuleProcessingState,visibility,createdDateTime'
$allAssignedGroupsCache = $null

$rowNumber = 1
foreach ($row in $rows) {
    $groupDisplayName = ([string]$row.GroupDisplayName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($groupDisplayName)) {
            throw 'GroupDisplayName is required. Use * to inventory all assigned security groups.'
        }

        $groups = @()
        if ($groupDisplayName -eq '*') {
            if ($null -eq $allAssignedGroupsCache) {
                $allGroups = @(Invoke-WithRetry -OperationName 'Load all groups for assigned security group inventory' -ScriptBlock {
                    Get-MgGroup -All -Property $groupSelect -ErrorAction Stop
                })
                $allAssignedGroupsCache = @($allGroups | Where-Object { Test-IsAssignedSecurityGroup -Group $_ } | Sort-Object -Property DisplayName, Id)
            }

            $groups = @($allAssignedGroupsCache)
        }
        else {
            $escapedName = Escape-ODataString -Value $groupDisplayName
            $candidateGroups = @(Invoke-WithRetry -OperationName "Lookup group $groupDisplayName" -ScriptBlock {
                Get-MgGroup -Filter "displayName eq '$escapedName'" -ConsistencyLevel eventual -Property $groupSelect -ErrorAction Stop
            })

            $groups = @($candidateGroups | Where-Object { Test-IsAssignedSecurityGroup -Group $_ })
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'GetEntraSecurityGroup' -Status 'NotFound' -Message 'No matching assigned security groups were found.' -Data ([ordered]@{
                        GroupId                       = ''
                        GroupDisplayName              = $groupDisplayName
                        Description                   = ''
                        MailNickname                  = ''
                        SecurityEnabled               = ''
                        MailEnabled                   = ''
                        MembershipType                = ''
                        MembershipRuleProcessingState = ''
                        Visibility                    = ''
                        CreatedDateTime               = ''
                    })))
            $rowNumber++
            continue
        }

        $sortedGroups = @($groups | Sort-Object -Property DisplayName, Id)
        foreach ($group in $sortedGroups) {
            $groupId = ([string]$group.Id).Trim()
            $displayName = ([string]$group.DisplayName).Trim()
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$displayName|$groupId" -Action 'GetEntraSecurityGroup' -Status 'Completed' -Message 'Assigned security group exported.' -Data ([ordered]@{
                        GroupId                       = $groupId
                        GroupDisplayName              = $displayName
                        Description                   = ([string]$group.Description).Trim()
                        MailNickname                  = ([string]$group.MailNickname).Trim()
                        SecurityEnabled               = [string]$group.SecurityEnabled
                        MailEnabled                   = [string]$group.MailEnabled
                        MembershipType                = 'Assigned'
                        MembershipRuleProcessingState = ([string]$group.MembershipRuleProcessingState).Trim()
                        Visibility                    = ([string]$group.Visibility).Trim()
                        CreatedDateTime               = [string]$group.CreatedDateTime
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupDisplayName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'GetEntraSecurityGroup' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    GroupId                       = ''
                    GroupDisplayName              = $groupDisplayName
                    Description                   = ''
                    MailNickname                  = ''
                    SecurityEnabled               = ''
                    MailEnabled                   = ''
                    MembershipType                = ''
                    MembershipRuleProcessingState = ''
                    Visibility                    = ''
                    CreatedDateTime               = ''
                })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID assigned security group inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







