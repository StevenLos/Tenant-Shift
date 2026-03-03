<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-000101

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Users
Microsoft.Graph.Groups

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M3007-Set-EntraSecurityGroupMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'UserPrincipalName',
    'MemberAction'
)

Write-Status -Message 'Starting Entra ID group membership script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Groups')
Ensure-GraphConnection -RequiredScopes @('Group.ReadWrite.All', 'Group.Read.All', 'User.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $groupDisplayName = ([string]$row.GroupDisplayName).Trim()
    $upn = ([string]$row.UserPrincipalName).Trim()
    $actionLabel = 'AddMember'

    try {
        if ([string]::IsNullOrWhiteSpace($groupDisplayName) -or [string]::IsNullOrWhiteSpace($upn)) {
            throw 'GroupDisplayName and UserPrincipalName are required.'
        }

        $memberActionRaw = 'Add'
        if ($row.PSObject.Properties.Name -contains 'MemberAction') {
            $memberActionCandidate = ([string]$row.MemberAction).Trim()
            if (-not [string]::IsNullOrWhiteSpace($memberActionCandidate)) {
                $memberActionRaw = $memberActionCandidate
            }
        }

        $memberAction = $memberActionRaw.ToLowerInvariant()
        if ($memberAction -notin @('add', 'remove')) {
            throw "MemberAction '$memberActionRaw' is invalid. Use Add or Remove."
        }

        $actionLabel = if ($memberAction -eq 'remove') { 'RemoveMember' } else { 'AddMember' }

        $escapedGroupName = Escape-ODataString -Value $groupDisplayName
        $groups = @(Invoke-WithRetry -OperationName "Lookup security group $groupDisplayName" -ScriptBlock {
            Get-MgGroup -Filter "displayName eq '$escapedGroupName'" -ConsistencyLevel eventual -ErrorAction Stop
        })
        if ($groups.Count -eq 0) {
            throw "Group '$groupDisplayName' was not found."
        }
        if ($groups.Count -gt 1) {
            throw "Multiple groups found with display name '$groupDisplayName'. Use unique names before running this script."
        }
        $group = $groups[0]

        $escapedUpn = Escape-ODataString -Value $upn
        $users = @(Invoke-WithRetry -OperationName "Lookup user $upn" -ScriptBlock {
            Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -ErrorAction Stop
        })
        if ($users.Count -eq 0) {
            throw "User '$upn' was not found."
        }

        $user = $users[0]

        $existingMembership = Invoke-WithRetry -OperationName "Check membership $groupDisplayName -> $upn" -ScriptBlock {
            Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop |
                Where-Object { $_.Id -eq $user.Id } |
                Select-Object -First 1
        }

        if ($memberAction -eq 'add') {
            if ($existingMembership) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$groupDisplayName|$upn" -Action $actionLabel -Status 'Skipped' -Message 'User is already a member of the group.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$groupDisplayName -> $upn", 'Add user to Entra ID security group')) {
                $memberRef = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)" }
                Invoke-WithRetry -OperationName "Add membership $groupDisplayName -> $upn" -ScriptBlock {
                    New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter $memberRef -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$groupDisplayName|$upn" -Action $actionLabel -Status 'Added' -Message 'User added to group successfully.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$groupDisplayName|$upn" -Action $actionLabel -Status 'WhatIf' -Message 'Membership update skipped due to WhatIf.'))
            }
        }
        else {
            if (-not $existingMembership) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$groupDisplayName|$upn" -Action $actionLabel -Status 'Skipped' -Message 'User is not currently a member of the group.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$groupDisplayName -> $upn", 'Remove user from Entra ID security group')) {
                Invoke-WithRetry -OperationName "Remove membership $groupDisplayName -> $upn" -ScriptBlock {
                    Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$groupDisplayName|$upn" -Action $actionLabel -Status 'Removed' -Message 'User removed from group successfully.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$groupDisplayName|$upn" -Action $actionLabel -Status 'WhatIf' -Message 'Membership update skipped due to WhatIf.'))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupDisplayName|$upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$groupDisplayName|$upn" -Action $actionLabel -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID group membership script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}




