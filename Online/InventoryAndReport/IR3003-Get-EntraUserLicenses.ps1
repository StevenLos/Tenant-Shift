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
Microsoft.Graph.Users
Microsoft.Graph.Identity.DirectoryManagement

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3003-Get-EntraUserLicenses_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-CanonicalGuidString {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    $guidValue = [Guid]::Empty
    if (-not [Guid]::TryParse($text, [ref]$guidValue)) {
        return ''
    }

    return $guidValue.Guid.ToLowerInvariant()
}

$requiredHeaders = @(
    'UserPrincipalName'
)

Write-Status -Message 'Starting Entra ID user license inventory script.'
Assert-ModuleCurrent -ModuleNames @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Identity.DirectoryManagement'
)
Ensure-GraphConnection -RequiredScopes @('User.Read.All', 'Directory.Read.All', 'Organization.Read.All')

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

Write-Status -Message 'Loading subscribed SKU metadata.'
$subscribedSkus = @(Invoke-WithRetry -OperationName 'Load subscribed SKUs' -ScriptBlock {
    Get-MgSubscribedSku -All -ErrorAction Stop
})

$skuById = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$servicePlansBySku = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($sku in $subscribedSkus) {
    $skuId = Get-CanonicalGuidString -Value $sku.SkuId
    if ([string]::IsNullOrWhiteSpace($skuId)) {
        continue
    }

    $skuById[$skuId] = $sku

    $planMap = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($plan in @($sku.ServicePlans)) {
        $planId = Get-CanonicalGuidString -Value $plan.ServicePlanId
        if ([string]::IsNullOrWhiteSpace($planId)) {
            continue
        }

        $planMap[$planId] = ([string]$plan.ServicePlanName).Trim()
    }

    $servicePlansBySku[$skuId] = $planMap
}

$userSelect = 'id,userPrincipalName,assignedLicenses'

$rowNumber = 1
foreach ($row in $rows) {
    $userPrincipalName = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
            throw 'UserPrincipalName is required. Use * to inventory all users.'
        }

        $users = @()
        if ($userPrincipalName -eq '*') {
            $users = @(Invoke-WithRetry -OperationName 'Load all users with license assignments' -ScriptBlock {
                Get-MgUser -All -Property $userSelect -ErrorAction Stop
            })
        }
        else {
            $escapedUpn = Escape-ODataString -Value $userPrincipalName
            $users = @(Invoke-WithRetry -OperationName "Lookup user $userPrincipalName" -ScriptBlock {
                Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -Property $userSelect -ErrorAction Stop
            })
        }

        if ($users.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'GetEntraUserLicense' -Status 'NotFound' -Message 'No matching users were found.' -Data ([ordered]@{
                        UserId                 = ''
                        UserPrincipalName      = $userPrincipalName
                        SkuId                  = ''
                        SkuPartNumber          = ''
                        DisabledPlanIds        = ''
                        DisabledPlanNames      = ''
                        AssignedLicenseCount   = ''
                    })))
            $rowNumber++
            continue
        }

        $sortedUsers = @($users | Sort-Object -Property UserPrincipalName, Id)
        foreach ($user in $sortedUsers) {
            $upn = ([string]$user.UserPrincipalName).Trim()
            $userId = ([string]$user.Id).Trim()
            $assignedLicenses = @($user.AssignedLicenses)

            if ($assignedLicenses.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $upn -Action 'GetEntraUserLicense' -Status 'Completed' -Message 'User has no assigned licenses.' -Data ([ordered]@{
                            UserId               = $userId
                            UserPrincipalName    = $upn
                            SkuId                = ''
                            SkuPartNumber        = ''
                            DisabledPlanIds      = ''
                            DisabledPlanNames    = ''
                            AssignedLicenseCount = '0'
                        })))
                continue
            }

            foreach ($assignment in $assignedLicenses) {
                $skuId = Get-CanonicalGuidString -Value $assignment.SkuId
                $skuPartNumber = ''
                $disabledPlanNames = @()

                if (-not [string]::IsNullOrWhiteSpace($skuId) -and $skuById.ContainsKey($skuId)) {
                    $skuPartNumber = ([string]$skuById[$skuId].SkuPartNumber).Trim()
                }

                $disabledPlanIds = [System.Collections.Generic.List[string]]::new()
                foreach ($planIdRaw in @($assignment.DisabledPlans)) {
                    $planId = Get-CanonicalGuidString -Value $planIdRaw
                    if ([string]::IsNullOrWhiteSpace($planId)) {
                        continue
                    }

                    if (-not $disabledPlanIds.Contains($planId)) {
                        $disabledPlanIds.Add($planId)
                    }

                    if ($servicePlansBySku.ContainsKey($skuId)) {
                        $planMap = [System.Collections.Generic.Dictionary[string, string]]$servicePlansBySku[$skuId]
                        if ($planMap.ContainsKey($planId)) {
                            $planName = ([string]$planMap[$planId]).Trim()
                            if (-not [string]::IsNullOrWhiteSpace($planName)) {
                                $disabledPlanNames += $planName
                            }
                        }
                    }
                }

                $disabledPlanNameDistinct = @($disabledPlanNames | Sort-Object -Unique)
                $primaryKey = if ([string]::IsNullOrWhiteSpace($skuPartNumber)) { "$upn|$skuId" } else { "$upn|$skuPartNumber" }
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetEntraUserLicense' -Status 'Completed' -Message 'License assignment exported.' -Data ([ordered]@{
                            UserId               = $userId
                            UserPrincipalName    = $upn
                            SkuId                = $skuId
                            SkuPartNumber        = $skuPartNumber
                            DisabledPlanIds      = (@($disabledPlanIds | Sort-Object) -join ';')
                            DisabledPlanNames    = ($disabledPlanNameDistinct -join ';')
                            AssignedLicenseCount = [string]$assignedLicenses.Count
                        })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($userPrincipalName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'GetEntraUserLicense' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    UserId                 = ''
                    UserPrincipalName      = $userPrincipalName
                    SkuId                  = ''
                    SkuPartNumber          = ''
                    DisabledPlanIds        = ''
                    DisabledPlanNames      = ''
                    AssignedLicenseCount   = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID user license inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}












