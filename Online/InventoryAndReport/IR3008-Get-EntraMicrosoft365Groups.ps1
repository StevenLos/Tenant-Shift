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
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3008-Get-EntraMicrosoft365Groups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Test-IsMicrosoft365Group {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Group
    )

    $groupTypes = @($Group.GroupTypes)
    $isM365 = ($groupTypes -contains 'Unified') -and ($Group.MailEnabled -eq $true) -and ($Group.SecurityEnabled -eq $false)
    return $isM365
}

function Get-DirectoryObjectType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$DirectoryObject
    )

    $odataType = ''
    if ($DirectoryObject.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $additional = $DirectoryObject.AdditionalProperties
        if ($additional) {
            try {
                if ($additional.ContainsKey('@odata.type')) {
                    $odataType = ([string]$additional['@odata.type']).Trim()
                }
            }
            catch {
                # Best effort only.
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($odataType)) {
        return 'unknown'
    }

    $normalized = $odataType.TrimStart('#').ToLowerInvariant()
    if ($normalized.StartsWith('microsoft.graph.')) {
        $normalized = $normalized.Substring('microsoft.graph.'.Length)
    }

    return $normalized
}

$requiredHeaders = @(
    'GroupMailNickname'
)

Write-Status -Message 'Starting Entra ID Microsoft 365 group inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('Group.Read.All', 'User.Read.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$groupSelect = 'id,displayName,description,mailNickname,groupTypes,visibility,mailEnabled,securityEnabled,createdDateTime'
$allM365GroupsCache = $null
$userById = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

$rowNumber = 1
foreach ($row in $rows) {
    $groupMailNickname = ([string]$row.GroupMailNickname).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($groupMailNickname)) {
            throw 'GroupMailNickname is required. Use * to inventory all Microsoft 365 groups.'
        }

        $groups = @()
        if ($groupMailNickname -eq '*') {
            if ($null -eq $allM365GroupsCache) {
                $allGroups = @(Invoke-WithRetry -OperationName 'Load all groups for Microsoft 365 group inventory' -ScriptBlock {
                    Get-MgGroup -All -Property $groupSelect -ErrorAction Stop
                })
                $allM365GroupsCache = @($allGroups | Where-Object { Test-IsMicrosoft365Group -Group $_ } | Sort-Object -Property MailNickname, DisplayName, Id)
            }

            $groups = @($allM365GroupsCache)
        }
        else {
            $escapedAlias = Escape-ODataString -Value $groupMailNickname
            $candidateGroups = @(Invoke-WithRetry -OperationName "Lookup group alias $groupMailNickname" -ScriptBlock {
                Get-MgGroup -Filter "mailNickname eq '$escapedAlias'" -ConsistencyLevel eventual -Property $groupSelect -ErrorAction Stop
            })

            $groups = @($candidateGroups | Where-Object { Test-IsMicrosoft365Group -Group $_ })
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'GetEntraMicrosoft365Group' -Status 'NotFound' -Message 'No matching Microsoft 365 groups were found.' -Data ([ordered]@{
                        GroupId                  = ''
                        GroupDisplayName         = ''
                        GroupMailNickname        = $groupMailNickname
                        Description              = ''
                        Visibility               = ''
                        GroupTypes               = ''
                        MailEnabled              = ''
                        SecurityEnabled          = ''
                        OwnersUserPrincipalNames = ''
                        OwnersObjectIds          = ''
                        OwnersCount              = ''
                        CreatedDateTime          = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property MailNickname, DisplayName, Id)) {
            $groupId = ([string]$group.Id).Trim()
            $groupDisplayName = ([string]$group.DisplayName).Trim()
            $resolvedAlias = ([string]$group.MailNickname).Trim()

            $owners = @(Invoke-WithRetry -OperationName "Load owners for group $resolvedAlias" -ScriptBlock {
                Get-MgGroupOwner -GroupId $groupId -All -ErrorAction Stop
            })

            $ownerUpns = [System.Collections.Generic.List[string]]::new()
            $ownerObjectIds = [System.Collections.Generic.List[string]]::new()

            foreach ($owner in @($owners | Sort-Object -Property Id)) {
                $ownerId = ([string]$owner.Id).Trim()
                if ([string]::IsNullOrWhiteSpace($ownerId)) {
                    continue
                }

                if (-not $ownerObjectIds.Contains($ownerId)) {
                    $ownerObjectIds.Add($ownerId)
                }

                $ownerType = Get-DirectoryObjectType -DirectoryObject $owner
                if ($ownerType -ne 'user') {
                    continue
                }

                try {
                    $ownerUser = $null
                    if ($userById.ContainsKey($ownerId)) {
                        $ownerUser = $userById[$ownerId]
                    }
                    else {
                        $ownerUser = Invoke-WithRetry -OperationName "Load user owner details $ownerId" -ScriptBlock {
                            Get-MgUser -UserId $ownerId -Property 'id,userPrincipalName' -ErrorAction Stop
                        }
                        $userById[$ownerId] = $ownerUser
                    }

                    $ownerUpn = ([string]$ownerUser.UserPrincipalName).Trim()
                    if (-not [string]::IsNullOrWhiteSpace($ownerUpn) -and -not $ownerUpns.Contains($ownerUpn)) {
                        $ownerUpns.Add($ownerUpn)
                    }
                }
                catch {
                    Write-Status -Message "Owner detail lookup failed for owner ID '$ownerId' in group '$groupDisplayName': $($_.Exception.Message)" -Level WARN
                }
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$resolvedAlias|$groupId" -Action 'GetEntraMicrosoft365Group' -Status 'Completed' -Message 'Microsoft 365 group exported.' -Data ([ordered]@{
                        GroupId                  = $groupId
                        GroupDisplayName         = $groupDisplayName
                        GroupMailNickname        = $resolvedAlias
                        Description              = ([string]$group.Description).Trim()
                        Visibility               = ([string]$group.Visibility).Trim()
                        GroupTypes               = (@($group.GroupTypes) -join ';')
                        MailEnabled              = [string]$group.MailEnabled
                        SecurityEnabled          = [string]$group.SecurityEnabled
                        OwnersUserPrincipalNames = (@($ownerUpns | Sort-Object) -join ';')
                        OwnersObjectIds          = (@($ownerObjectIds | Sort-Object) -join ';')
                        OwnersCount              = [string]$ownerObjectIds.Count
                        CreatedDateTime          = [string]$group.CreatedDateTime
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupMailNickname) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'GetEntraMicrosoft365Group' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    GroupId                  = ''
                    GroupDisplayName         = ''
                    GroupMailNickname        = $groupMailNickname
                    Description              = ''
                    Visibility               = ''
                    GroupTypes               = ''
                    MailEnabled              = ''
                    SecurityEnabled          = ''
                    OwnersUserPrincipalNames = ''
                    OwnersObjectIds          = ''
                    OwnersCount              = ''
                    CreatedDateTime          = ''
                })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID Microsoft 365 group inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







