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

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M3006-Update-EntraDynamicUserSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-TrimmedValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

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
    'Action',
    'Notes',
    'GroupMailNickname',
    'GroupDisplayName',
    'Description',
    'MailNickname',
    'MembershipRule',
    'MembershipRuleProcessingState',
    'ClearAttributes'
)

Write-Status -Message 'Starting Entra dynamic user security group update script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')
Ensure-GraphConnection -RequiredScopes @('Group.ReadWrite.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $groupMailNickname = Get-TrimmedValue -Value $row.GroupMailNickname

    try {
        if ([string]::IsNullOrWhiteSpace($groupMailNickname)) {
            throw 'GroupMailNickname is required.'
        }

        $escapedAlias = Escape-ODataString -Value $groupMailNickname
        $groups = @(Invoke-WithRetry -OperationName "Lookup group alias $groupMailNickname" -ScriptBlock {
            Get-MgGroup -Filter "mailNickname eq '$escapedAlias'" -ConsistencyLevel eventual -Property 'id,displayName,description,mailNickname,groupTypes,securityEnabled,mailEnabled,membershipRule,membershipRuleProcessingState' -ErrorAction Stop
        })

        if ($groups.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateDynamicUserSecurityGroup' -Status 'NotFound' -Message 'Group not found.'))
            $rowNumber++
            continue
        }

        if ($groups.Count -gt 1) {
            throw "Multiple groups found with mailNickname '$groupMailNickname'. Resolve duplicate aliases before retrying."
        }

        $group = $groups[0]
        if (-not (Test-IsDynamicUserSecurityGroup -Group $group)) {
            throw "Group '$groupMailNickname' is not a dynamic user security group."
        }

        $body = @{}

        $groupDisplayName = Get-TrimmedValue -Value $row.GroupDisplayName
        if (-not [string]::IsNullOrWhiteSpace($groupDisplayName)) {
            $body['displayName'] = $groupDisplayName
        }

        $description = Get-TrimmedValue -Value $row.Description
        if (-not [string]::IsNullOrWhiteSpace($description)) {
            $body['description'] = $description
        }

        $newAlias = Get-TrimmedValue -Value $row.MailNickname
        if (-not [string]::IsNullOrWhiteSpace($newAlias)) {
            $body['mailNickname'] = $newAlias
        }

        $membershipRule = Get-TrimmedValue -Value $row.MembershipRule
        if (-not [string]::IsNullOrWhiteSpace($membershipRule)) {
            $body['membershipRule'] = $membershipRule
        }

        $processingState = Get-TrimmedValue -Value $row.MembershipRuleProcessingState
        if (-not [string]::IsNullOrWhiteSpace($processingState)) {
            if ($processingState -notin @('On', 'Paused')) {
                throw "MembershipRuleProcessingState '$processingState' is invalid. Use On or Paused."
            }

            $body['membershipRuleProcessingState'] = $processingState
        }

        $clearRequested = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ClearAttributes)
        foreach ($clearName in $clearRequested) {
            if ($clearName -eq 'Description') {
                if ($body.ContainsKey('description')) {
                    throw "Description cannot be set and cleared in the same row."
                }

                $body['description'] = $null
                continue
            }

            throw "ClearAttributes value '$clearName' is not supported."
        }

        if ($body.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateDynamicUserSecurityGroup' -Status 'Skipped' -Message 'No updates were requested.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($groupMailNickname, 'Update dynamic user security group attributes')) {
            Invoke-WithRetry -OperationName "Update dynamic user security group $groupMailNickname" -ScriptBlock {
                Update-MgGroup -GroupId $group.Id -BodyParameter $body -ErrorAction Stop | Out-Null
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateDynamicUserSecurityGroup' -Status 'Updated' -Message 'Dynamic user security group updated successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateDynamicUserSecurityGroup' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupMailNickname) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateDynamicUserSecurityGroup' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra dynamic user security group update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
