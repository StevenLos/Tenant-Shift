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

.SYNOPSIS
    Modifies EntraAssignedSecurityGroups in Microsoft 365.

.DESCRIPTION
    Updates EntraAssignedSecurityGroups in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M3005-Update-EntraAssignedSecurityGroups.ps1 -InputCsvPath .\3005.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M3005-Update-EntraAssignedSecurityGroups.ps1 -InputCsvPath .\3005.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Groups
    Required roles:   Global Administrator or appropriate workload-specific role
    Limitations:      None known.

    CSV Fields:
    Column                Type      Required  Description
    --------------------  ----      --------  -----------
    Action                String    Yes       <fill in description>
    Notes                 String    Yes       <fill in description>
    GroupMailNickname     String    Yes       <fill in description>
    GroupDisplayName      String    Yes       <fill in description>
    Description           String    Yes       <fill in description>
    MailNickname          String    Yes       <fill in description>
    ClearAttributes       String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3005-Update-EntraAssignedSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Test-IsAssignedSecurityGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Group
    )

    $membershipRule = Get-TrimmedValue -Value $Group.MembershipRule
    return ($Group.SecurityEnabled -eq $true -and $Group.MailEnabled -eq $false -and [string]::IsNullOrWhiteSpace($membershipRule))
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'GroupMailNickname',
    'GroupDisplayName',
    'Description',
    'MailNickname',
    'ClearAttributes'
)

Write-Status -Message 'Starting Entra assigned security group update script.'
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
            Get-MgGroup -Filter "mailNickname eq '$escapedAlias'" -ConsistencyLevel eventual -Property 'id,displayName,description,mailNickname,groupTypes,securityEnabled,mailEnabled,membershipRule' -ErrorAction Stop
        })

        if ($groups.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateAssignedSecurityGroup' -Status 'NotFound' -Message 'Group not found.'))
            $rowNumber++
            continue
        }

        if ($groups.Count -gt 1) {
            throw "Multiple groups found with mailNickname '$groupMailNickname'. Resolve duplicate aliases before retrying."
        }

        $group = $groups[0]
        if (-not (Test-IsAssignedSecurityGroup -Group $group)) {
            throw "Group '$groupMailNickname' is not an assigned security group."
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
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateAssignedSecurityGroup' -Status 'Skipped' -Message 'No updates were requested.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($groupMailNickname, 'Update assigned security group attributes')) {
            Invoke-WithRetry -OperationName "Update assigned security group $groupMailNickname" -ScriptBlock {
                Update-MgGroup -GroupId $group.Id -BodyParameter $body -ErrorAction Stop | Out-Null
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateAssignedSecurityGroup' -Status 'Updated' -Message 'Assigned security group updated successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateAssignedSecurityGroup' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupMailNickname) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateAssignedSecurityGroup' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra assigned security group update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
