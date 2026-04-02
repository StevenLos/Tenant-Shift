<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-160000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets EntraPrivilegedRoles and exports results to CSV.

.DESCRIPTION
    Gets EntraPrivilegedRoles from Microsoft 365 and writes the results to a CSV file.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all objects in scope (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all objects in scope rather than processing from an input CSV file. Uses the DiscoverAll parameter set.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-IR3004-Get-EntraPrivilegedRoles.ps1 -InputCsvPath .\3004.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3004-Get-EntraPrivilegedRoles.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication
    Required roles:   Global Administrator or appropriate workload-specific role
    Limitations:      None known.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-MEID-0060-Get-EntraPrivilegedRoles_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-GraphPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$Object,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    if ($null -eq $Object) {
        return $null
    }

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($PropertyName)) {
            return $Object[$PropertyName]
        }
    }

    if ($Object.PSObject.Properties.Name -contains $PropertyName) {
        return $Object.$PropertyName
    }

    if ($Object.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $additional = $Object.AdditionalProperties
        if ($additional -is [System.Collections.IDictionary] -and $additional.Contains($PropertyName)) {
            return $additional[$PropertyName]
        }
    }

    return $null
}

function Invoke-GraphPagedRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter(Mandatory)]
        [string]$OperationName
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $response = Invoke-WithRetry -OperationName $OperationName -ScriptBlock {
            Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject -ErrorAction Stop
        }

        $pageValues = Get-GraphPropertyValue -Object $response -PropertyName 'value'
        if ($null -ne $pageValues) {
            foreach ($item in @($pageValues)) {
                $results.Add($item)
            }
        }

        $nextLink = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $response -PropertyName '@odata.nextLink')
        if ([string]::IsNullOrWhiteSpace($nextLink)) {
            $nextLink = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $response -PropertyName 'odata.nextLink')
        }

        $nextUri = $nextLink
    }

    return $results.ToArray()
}

function Get-MemberObjectType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$Member
    )

    $odataType = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $Member -PropertyName '@odata.type')
    if ([string]::IsNullOrWhiteSpace($odataType)) {
        $odataType = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $Member -PropertyName 'odata.type')
    }

    if ([string]::IsNullOrWhiteSpace($odataType)) {
        return ''
    }

    if ($odataType.StartsWith('#microsoft.graph.')) {
        return $odataType.Substring('#microsoft.graph.'.Length)
    }

    return $odataType.TrimStart('#')
}

$requiredHeaders = @(
    'RoleDisplayName'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'RoleDisplayName',
    'RoleId',
    'RoleTemplateId',
    'RoleDescription',
    'MemberDisplayName',
    'MemberObjectType',
    'MemberObjectId',
    'MemberUserPrincipalName',
    'MemberMail',
    'MemberAppId'
)

Write-Status -Message 'Starting Entra privileged role inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication')
Ensure-GraphConnection -RequiredScopes @('Directory.Read.All', 'RoleManagement.Read.Directory')

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
$allRolesCache = $null

$rowNumber = 1
foreach ($row in $rows) {
    $roleDisplayName = Get-TrimmedValue -Value $row.RoleDisplayName

    try {
        if ([string]::IsNullOrWhiteSpace($roleDisplayName)) {
            throw 'RoleDisplayName is required. Use * to inventory all activated privileged roles.'
        }

        if ($null -eq $allRolesCache) {
            $allRolesCache = @(
                Invoke-GraphPagedRequest -Uri '/v1.0/directoryRoles?$select=id,displayName,description,roleTemplateId' -OperationName 'Load activated Entra directory roles'
            )
        }

        $roles = @()
        if ($roleDisplayName -eq '*') {
            $roles = @($allRolesCache)
        }
        else {
            $roles = @($allRolesCache | Where-Object { (Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $_ -PropertyName 'displayName')) -ieq $roleDisplayName })
        }

        if ($roles.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $roleDisplayName -Action 'GetEntraPrivilegedRoleMembership' -Status 'NotFound' -Message 'No matching activated Entra roles were found.' -Data ([ordered]@{
                        RoleId                   = ''
                        RoleDisplayName          = $roleDisplayName
                        RoleDescription          = ''
                        RoleTemplateId           = ''
                        MemberObjectId           = ''
                        MemberObjectType         = ''
                        MemberDisplayName        = ''
                        MemberUserPrincipalName  = ''
                        MemberMail               = ''
                        MemberAppId              = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($role in @($roles | Sort-Object -Property @{ Expression = { Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $_ -PropertyName 'displayName') } }, @{ Expression = { Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $_ -PropertyName 'id') } })) {
            $resolvedRoleId = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $role -PropertyName 'id')
            $resolvedRoleName = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $role -PropertyName 'displayName')

            $members = @(
                Invoke-GraphPagedRequest -Uri "/v1.0/directoryRoles/$resolvedRoleId/members?`$select=id,displayName,userPrincipalName,mail,appId" -OperationName "Load members for role $resolvedRoleName"
            )

            if ($members.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$resolvedRoleName|$resolvedRoleId" -Action 'GetEntraPrivilegedRoleMembership' -Status 'Completed' -Message 'Role has no active members.' -Data ([ordered]@{
                                RoleId                   = $resolvedRoleId
                                RoleDisplayName          = $resolvedRoleName
                                RoleDescription          = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $role -PropertyName 'description')
                                RoleTemplateId           = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $role -PropertyName 'roleTemplateId')
                                MemberObjectId           = ''
                                MemberObjectType         = ''
                                MemberDisplayName        = ''
                                MemberUserPrincipalName  = ''
                                MemberMail               = ''
                                MemberAppId              = ''
                            })))

                continue
            }

            foreach ($member in @($members | Sort-Object -Property @{ Expression = { Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $_ -PropertyName 'displayName') } }, @{ Expression = { Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $_ -PropertyName 'id') } })) {
                $memberId = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $member -PropertyName 'id')
                $memberDisplayName = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $member -PropertyName 'displayName')
                $resultKey = "$resolvedRoleName|$resolvedRoleId|$memberId"

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $resultKey -Action 'GetEntraPrivilegedRoleMembership' -Status 'Completed' -Message 'Privileged role member exported.' -Data ([ordered]@{
                                RoleId                   = $resolvedRoleId
                                RoleDisplayName          = $resolvedRoleName
                                RoleDescription          = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $role -PropertyName 'description')
                                RoleTemplateId           = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $role -PropertyName 'roleTemplateId')
                                MemberObjectId           = $memberId
                                MemberObjectType         = Get-MemberObjectType -Member $member
                                MemberDisplayName        = $memberDisplayName
                                MemberUserPrincipalName  = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $member -PropertyName 'userPrincipalName')
                                MemberMail               = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $member -PropertyName 'mail')
                                MemberAppId              = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $member -PropertyName 'appId')
                            })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($roleDisplayName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $roleDisplayName -Action 'GetEntraPrivilegedRoleMembership' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    RoleId                   = ''
                    RoleDisplayName          = $roleDisplayName
                    RoleDescription          = ''
                    RoleTemplateId           = ''
                    MemberObjectId           = ''
                    MemberObjectType         = ''
                    MemberDisplayName        = ''
                    MemberUserPrincipalName  = ''
                    MemberMail               = ''
                    MemberAppId              = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra privileged role inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
