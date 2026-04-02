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

.SYNOPSIS
    Provisions EntraDynamicUserSecurityGroups in Microsoft 365.

.DESCRIPTION
    Creates EntraDynamicUserSecurityGroups in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P3006-Create-EntraDynamicUserSecurityGroups.ps1 -InputCsvPath .\3006.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P3006-Create-EntraDynamicUserSecurityGroups.ps1 -InputCsvPath .\3006.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Groups
    Required roles:   Global Administrator or appropriate workload-specific role
    Limitations:      None known.

    CSV Fields:
    Column                         Type      Required  Description
    -----------------------------  ----      --------  -----------
    GroupDisplayName               String    Yes       <fill in description>
    MailNickname                   String    Yes       <fill in description>
    Description                    String    Yes       <fill in description>
    MembershipRule                 String    Yes       <fill in description>
    MembershipRuleProcessingState  String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P3006-Create-EntraDynamicUserSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {


$requiredHeaders = @(
    'GroupDisplayName',
    'MailNickname',
    'Description',
    'MembershipRule',
    'MembershipRuleProcessingState'
)

Write-Status -Message 'Starting Entra ID dynamic user security group creation script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')
Ensure-GraphConnection -RequiredScopes @('Group.ReadWrite.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $groupDisplayName = ([string]$row.GroupDisplayName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($groupDisplayName)) {
            throw 'GroupDisplayName is required.'
        }

        $mailNickname = ([string]$row.MailNickname).Trim()
        if ([string]::IsNullOrWhiteSpace($mailNickname)) {
            throw 'MailNickname is required.'
        }

        $membershipRule = ([string]$row.MembershipRule).Trim()
        if ([string]::IsNullOrWhiteSpace($membershipRule)) {
            throw 'MembershipRule is required.'
        }

        $membershipRuleProcessingState = ([string]$row.MembershipRuleProcessingState).Trim()
        if ([string]::IsNullOrWhiteSpace($membershipRuleProcessingState)) {
            $membershipRuleProcessingState = 'On'
        }

        if ($membershipRuleProcessingState -notin @('On', 'Paused')) {
            throw "MembershipRuleProcessingState '$membershipRuleProcessingState' is invalid. Use On or Paused."
        }

        $escapedMailNickname = Escape-ODataString -Value $mailNickname
        $existingGroupsByAlias = @(Invoke-WithRetry -OperationName "Lookup dynamic user security group alias $mailNickname" -ScriptBlock {
            Get-MgGroup -Filter "mailNickname eq '$escapedMailNickname'" -ConsistencyLevel eventual -ErrorAction Stop
        })

        if ($existingGroupsByAlias.Count -gt 1) {
            throw "Multiple groups found with mailNickname '$mailNickname'. Resolve duplicate aliases before running this script."
        }

        if ($existingGroupsByAlias.Count -eq 1) {
            $existingGroup = Invoke-WithRetry -OperationName "Load group details for alias $mailNickname" -ScriptBlock {
                Get-MgGroup -GroupId $existingGroupsByAlias[0].Id -Property 'id,displayName,mailNickname,groupTypes,securityEnabled,mailEnabled,membershipRule,membershipRuleProcessingState' -ErrorAction Stop
            }

            $existingGroupTypes = @($existingGroup.GroupTypes)
            $isDynamicGroup = $existingGroupTypes -contains 'DynamicMembership'
            $isSecurityEnabled = ($existingGroup.SecurityEnabled -eq $true)
            $isMailDisabled = ($existingGroup.MailEnabled -eq $false)

            if (-not ($isDynamicGroup -and $isSecurityEnabled -and $isMailDisabled)) {
                throw "A group with mailNickname '$mailNickname' already exists but is not a dynamic user security group."
            }

            $existingMembershipRule = ([string]$existingGroup.MembershipRule).Trim()
            $existingProcessingState = ([string]$existingGroup.MembershipRuleProcessingState).Trim()

            $ruleMatches = $existingMembershipRule.Equals($membershipRule, [System.StringComparison]::OrdinalIgnoreCase)
            $stateMatches = $existingProcessingState.Equals($membershipRuleProcessingState, [System.StringComparison]::OrdinalIgnoreCase)

            if ($ruleMatches -and $stateMatches) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'CreateDynamicUserSecurityGroup' -Status 'Skipped' -Message 'Dynamic user security group already exists with the requested rule and processing state.'))
                $rowNumber++
                continue
            }

            throw "A dynamic user security group with mailNickname '$mailNickname' already exists, but its membership rule or processing state differs from the CSV request."
        }

        $body = @{
            displayName                   = $groupDisplayName
            mailEnabled                   = $false
            mailNickname                  = $mailNickname
            securityEnabled               = $true
            groupTypes                    = @('DynamicMembership')
            membershipRule                = $membershipRule
            membershipRuleProcessingState = $membershipRuleProcessingState
        }

        $description = ([string]$row.Description).Trim()
        if (-not [string]::IsNullOrWhiteSpace($description)) {
            $body.description = $description
        }

        if ($PSCmdlet.ShouldProcess($groupDisplayName, 'Create Entra ID dynamic user security group')) {
            Invoke-WithRetry -OperationName "Create dynamic user security group $groupDisplayName" -ScriptBlock {
                New-MgGroup -BodyParameter $body -ErrorAction Stop | Out-Null
            }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'CreateDynamicUserSecurityGroup' -Status 'Created' -Message 'Dynamic user security group created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'CreateDynamicUserSecurityGroup' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupDisplayName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'CreateDynamicUserSecurityGroup' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID dynamic user security group creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







