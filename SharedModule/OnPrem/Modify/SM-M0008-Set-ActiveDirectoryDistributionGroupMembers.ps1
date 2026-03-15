<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260302-201500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
ActiveDirectory

.MODULEVERSIONPOLICY
Best-effort PSGallery check (offline-safe)
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M0008-Set-ActiveDirectoryDistributionGroupMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$expectedGroupCategory = 'Distribution'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Resolve-TargetAdGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IdentityType,

        [Parameter(Mandatory)]
        [string]$IdentityValue
    )

    $normalizedType = $IdentityType.Trim().ToLowerInvariant()
    switch ($normalizedType) {
        'samaccountname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADGroup -Filter "SamAccountName -eq '$escaped'" -Properties GroupCategory,DistinguishedName,Name,SamAccountName,ObjectGuid -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'name' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADGroup -Filter "Name -eq '$escaped'" -Properties GroupCategory,DistinguishedName,Name,SamAccountName,ObjectGuid -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'distinguishedname' {
            return Get-ADGroup -Identity $IdentityValue -Properties GroupCategory,DistinguishedName,Name,SamAccountName,ObjectGuid -ErrorAction SilentlyContinue
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            return Get-ADGroup -Identity $guid -Properties GroupCategory,DistinguishedName,Name,SamAccountName,ObjectGuid -ErrorAction SilentlyContinue
        }
        default {
            throw "GroupIdentityType '$IdentityType' is invalid. Use SamAccountName, Name, DistinguishedName, or ObjectGuid."
        }
    }
}

function Resolve-MemberObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IdentityType,

        [Parameter(Mandatory)]
        [string]$IdentityValue
    )

    $normalizedType = $IdentityType.Trim().ToLowerInvariant()
    switch ($normalizedType) {
        'samaccountname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue

            $user = Get-ADUser -Filter "SamAccountName -eq '$escaped'" -Properties DistinguishedName,ObjectGuid,SamAccountName,UserPrincipalName,Name -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($user) { return $user }

            $group = Get-ADGroup -Filter "SamAccountName -eq '$escaped'" -Properties DistinguishedName,ObjectGuid,SamAccountName,Name -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($group) { return $group }

            $computer = Get-ADComputer -Filter "SamAccountName -eq '$escaped'" -Properties DistinguishedName,ObjectGuid,SamAccountName,Name -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($computer) { return $computer }

            return $null
        }
        'userprincipalname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADUser -Filter "UserPrincipalName -eq '$escaped'" -Properties DistinguishedName,ObjectGuid,SamAccountName,UserPrincipalName,Name -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'distinguishedname' {
            return Get-ADObject -Identity $IdentityValue -Properties DistinguishedName,ObjectGuid,samAccountName,userPrincipalName,Name,ObjectClass -ErrorAction SilentlyContinue
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            return Get-ADObject -Identity $guid -Properties DistinguishedName,ObjectGuid,samAccountName,userPrincipalName,Name,ObjectClass -ErrorAction SilentlyContinue
        }
        default {
            throw "MemberIdentityType '$IdentityType' is invalid. Use SamAccountName, UserPrincipalName, DistinguishedName, or ObjectGuid."
        }
    }
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'GroupIdentityType',
    'GroupIdentityValue',
    'MemberIdentityType',
    'MemberIdentityValue',
    'MemberAction'
)

Write-Status -Message 'Starting Active Directory distribution group membership update script.'
Ensure-ActiveDirectoryConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $groupIdentityType = Get-TrimmedValue -Value $row.GroupIdentityType
    $groupIdentityValue = Get-TrimmedValue -Value $row.GroupIdentityValue
    $memberIdentityType = Get-TrimmedValue -Value $row.MemberIdentityType
    $memberIdentityValue = Get-TrimmedValue -Value $row.MemberIdentityValue
    $memberAction = (Get-TrimmedValue -Value $row.MemberAction).ToLowerInvariant()

    $primaryKey = "${groupIdentityType}:$groupIdentityValue|${memberIdentityType}:$memberIdentityValue"

    try {
        if ([string]::IsNullOrWhiteSpace($groupIdentityType) -or [string]::IsNullOrWhiteSpace($groupIdentityValue)) {
            throw 'GroupIdentityType and GroupIdentityValue are required.'
        }

        if ([string]::IsNullOrWhiteSpace($memberIdentityType) -or [string]::IsNullOrWhiteSpace($memberIdentityValue)) {
            throw 'MemberIdentityType and MemberIdentityValue are required.'
        }

        if ($memberAction -notin @('add', 'remove')) {
            throw "MemberAction '$($row.MemberAction)' is invalid. Use Add or Remove."
        }

        $targetGroup = Invoke-WithRetry -OperationName "Resolve AD group $groupIdentityValue" -ScriptBlock {
            Resolve-TargetAdGroup -IdentityType $groupIdentityType -IdentityValue $groupIdentityValue
        }

        if (-not $targetGroup) {
            throw 'Target group was not found.'
        }

        if ((Get-TrimmedValue -Value $targetGroup.GroupCategory) -ne $expectedGroupCategory) {
            throw "Target group category '$($targetGroup.GroupCategory)' does not match expected '$expectedGroupCategory'."
        }

        $memberObject = Invoke-WithRetry -OperationName "Resolve AD member $memberIdentityValue" -ScriptBlock {
            Resolve-MemberObject -IdentityType $memberIdentityType -IdentityValue $memberIdentityValue
        }

        if (-not $memberObject) {
            throw 'Member object was not found.'
        }

        $memberDn = Get-TrimmedValue -Value $memberObject.DistinguishedName
        if ([string]::IsNullOrWhiteSpace($memberDn)) {
            throw 'Resolved member object does not contain DistinguishedName.'
        }

        $groupState = Invoke-WithRetry -OperationName "Load group members for $($targetGroup.SamAccountName)" -ScriptBlock {
            Get-ADGroup -Identity $targetGroup.ObjectGuid -Properties Member -ErrorAction Stop
        }

        $existingMemberDns = @($groupState.Member)
        $memberExists = $existingMemberDns -contains $memberDn

        if ($memberAction -eq 'add') {
            if ($memberExists) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'Skipped' -Message 'Member already present in group.'))
            }
            elseif ($PSCmdlet.ShouldProcess($targetGroup.SamAccountName, "Add member $memberDn")) {
                Invoke-WithRetry -OperationName "Add member to AD distribution group $($targetGroup.SamAccountName)" -ScriptBlock {
                    Add-ADGroupMember -Identity $targetGroup.ObjectGuid -Members $memberDn -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'Completed' -Message 'Member added to group.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'WhatIf' -Message 'Add skipped due to WhatIf.'))
            }
        }
        else {
            if (-not $memberExists) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'Skipped' -Message 'Member is not currently in group.'))
            }
            elseif ($PSCmdlet.ShouldProcess($targetGroup.SamAccountName, "Remove member $memberDn")) {
                Invoke-WithRetry -OperationName "Remove member from AD distribution group $($targetGroup.SamAccountName)" -ScriptBlock {
                    Remove-ADGroupMember -Identity $targetGroup.ObjectGuid -Members $memberDn -Confirm:$false -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'Completed' -Message 'Member removed from group.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'WhatIf' -Message 'Remove skipped due to WhatIf.'))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory distribution group membership update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
