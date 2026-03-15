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

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P3005-Create-EntraAssignedSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'Description'
)

Write-Status -Message 'Starting Entra ID assigned security group creation script.'
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

        $escapedMailNickname = Escape-ODataString -Value $mailNickname
        $existingGroupsByAlias = @(Invoke-WithRetry -OperationName "Lookup assigned security group alias $mailNickname" -ScriptBlock {
            Get-MgGroup -Filter "mailNickname eq '$escapedMailNickname'" -ConsistencyLevel eventual -ErrorAction Stop
        })

        if ($existingGroupsByAlias.Count -gt 1) {
            throw "Multiple groups found with mailNickname '$mailNickname'. Resolve duplicate aliases before running this script."
        }

        if ($existingGroupsByAlias.Count -eq 1) {
            $existingDisplayName = [string]$existingGroupsByAlias[0].DisplayName
            $message = if ([string]::IsNullOrWhiteSpace($existingDisplayName)) {
                "A security group with mailNickname '$mailNickname' already exists."
            }
            else {
                "A security group with mailNickname '$mailNickname' already exists (displayName '$existingDisplayName')."
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'CreateAssignedSecurityGroup' -Status 'Skipped' -Message $message))
            $rowNumber++
            continue
        }

        $body = @{
            displayName     = $groupDisplayName
            mailEnabled     = $false
            mailNickname    = $mailNickname
            securityEnabled = $true
            groupTypes      = @()
        }

        $description = ([string]$row.Description).Trim()
        if (-not [string]::IsNullOrWhiteSpace($description)) {
            $body.description = $description
        }

        if ($PSCmdlet.ShouldProcess($groupDisplayName, 'Create Entra ID assigned security group')) {
            Invoke-WithRetry -OperationName "Create assigned security group $groupDisplayName" -ScriptBlock {
                New-MgGroup -BodyParameter $body -ErrorAction Stop | Out-Null
            }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'CreateAssignedSecurityGroup' -Status 'Created' -Message 'Assigned security group created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'CreateAssignedSecurityGroup' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupDisplayName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'CreateAssignedSecurityGroup' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID assigned security group creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







