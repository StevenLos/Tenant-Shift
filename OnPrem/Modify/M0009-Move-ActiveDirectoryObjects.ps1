<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260302-191500

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M0009-Move-ActiveDirectoryObjects_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Escape-LdapFilterValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    $builder = [System.Text.StringBuilder]::new()
    foreach ($char in $Value.ToCharArray()) {
        switch ($char) {
            '\\' { [void]$builder.Append('\\5c') }
            '*' { [void]$builder.Append('\\2a') }
            '(' { [void]$builder.Append('\\28') }
            ')' { [void]$builder.Append('\\29') }
            ([char]0) { [void]$builder.Append('\\00') }
            default { [void]$builder.Append($char) }
        }
    }

    return $builder.ToString()
}

function Get-ObjectTypeLdapClause {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ObjectType
    )

    switch ($ObjectType.Trim().ToLowerInvariant()) {
        'any' { return '' }
        'user' { return '(&(objectClass=user)(!(objectClass=computer)))' }
        'group' { return '(objectClass=group)' }
        'contact' { return '(objectClass=contact)' }
        'organizationalunit' { return '(objectClass=organizationalUnit)' }
        'ou' { return '(objectClass=organizationalUnit)' }
        'computer' { return '(objectClass=computer)' }
        default {
            throw "ObjectType '$ObjectType' is invalid. Use Any, User, Group, Contact, OrganizationalUnit, OU, or Computer."
        }
    }
}

function Test-ObjectTypeMatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$AdObject,

        [Parameter(Mandatory)]
        [string]$ObjectType
    )

    $normalizedType = $ObjectType.Trim().ToLowerInvariant()
    if ($normalizedType -eq 'any') {
        return $true
    }

    $classes = @($AdObject.ObjectClass | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() })
    switch ($normalizedType) {
        'user' { return (($classes -contains 'user') -and (-not ($classes -contains 'computer'))) }
        'group' { return ($classes -contains 'group') }
        'contact' { return ($classes -contains 'contact') }
        'organizationalunit' { return ($classes -contains 'organizationalunit') }
        'ou' { return ($classes -contains 'organizationalunit') }
        'computer' { return ($classes -contains 'computer') }
    }

    return $false
}

function Resolve-TargetAdObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ObjectType,

        [Parameter(Mandatory)]
        [string]$IdentityType,

        [Parameter(Mandatory)]
        [string]$IdentityValue
    )

    $normalizedIdentityType = $IdentityType.Trim().ToLowerInvariant()

    if ($normalizedIdentityType -eq 'distinguishedname') {
        $candidate = Get-ADObject -Identity $IdentityValue -Properties * -ErrorAction SilentlyContinue
        if (-not $candidate) {
            return $null
        }

        if (-not (Test-ObjectTypeMatch -AdObject $candidate -ObjectType $ObjectType)) {
            throw "Resolved object does not match ObjectType '$ObjectType'."
        }

        return $candidate
    }

    if ($normalizedIdentityType -eq 'objectguid') {
        $guid = [guid]$IdentityValue
        $candidate = Get-ADObject -Identity $guid -Properties * -ErrorAction SilentlyContinue
        if (-not $candidate) {
            return $null
        }

        if (-not (Test-ObjectTypeMatch -AdObject $candidate -ObjectType $ObjectType)) {
            throw "Resolved object does not match ObjectType '$ObjectType'."
        }

        return $candidate
    }

    $escapedIdentityValue = Escape-LdapFilterValue -Value $IdentityValue
    $identityClause = switch ($normalizedIdentityType) {
        'samaccountname' { "(sAMAccountName=$escapedIdentityValue)" }
        'userprincipalname' { "(userPrincipalName=$escapedIdentityValue)" }
        'name' { "(name=$escapedIdentityValue)" }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use SamAccountName, UserPrincipalName, Name, DistinguishedName, or ObjectGuid."
        }
    }

    $typeClause = Get-ObjectTypeLdapClause -ObjectType $ObjectType
    $ldapFilter = if ([string]::IsNullOrWhiteSpace($typeClause)) { $identityClause } else { "(&$typeClause$identityClause)" }

    $matches = @(Get-ADObject -LDAPFilter $ldapFilter -Properties * -ErrorAction SilentlyContinue)
    if ($matches.Count -eq 0) {
        return $null
    }

    if ($matches.Count -gt 1) {
        throw "Identity '$IdentityType=$IdentityValue' resolved to $($matches.Count) objects. Use DistinguishedName or ObjectGuid for an unambiguous target."
    }

    return $matches[0]
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'ObjectType',
    'IdentityType',
    'IdentityValue',
    'TargetPath',
    'NewName',
    'ProtectionEnabled'
)

Write-Status -Message 'Starting Active Directory object move script.'
Ensure-ActiveDirectoryConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $objectType = Get-TrimmedValue -Value $row.ObjectType
    if ([string]::IsNullOrWhiteSpace($objectType)) {
        $objectType = 'Any'
    }

    $identityType = Get-TrimmedValue -Value $row.IdentityType
    $identityValue = Get-TrimmedValue -Value $row.IdentityValue
    $primaryKey = "$objectType|${identityType}:$identityValue"

    try {
        if ([string]::IsNullOrWhiteSpace($identityType) -or [string]::IsNullOrWhiteSpace($identityValue)) {
            throw 'IdentityType and IdentityValue are required.'
        }

        $targetObject = Invoke-WithRetry -OperationName "Resolve AD object $primaryKey" -ScriptBlock {
            Resolve-TargetAdObject -ObjectType $objectType -IdentityType $identityType -IdentityValue $identityValue
        }

        if (-not $targetObject) {
            throw 'Target object was not found.'
        }

        $resolvedKey = Get-TrimmedValue -Value $targetObject.DistinguishedName
        if ([string]::IsNullOrWhiteSpace($resolvedKey)) {
            $resolvedKey = "ObjectGuid:$($targetObject.ObjectGuid)"
        }

        $messages = [System.Collections.Generic.List[string]]::new()
        $changeCount = 0

        $newName = Get-TrimmedValue -Value $row.NewName
        if (-not [string]::IsNullOrWhiteSpace($newName) -and $newName -ne (Get-TrimmedValue -Value $targetObject.Name)) {
            $changeCount++

            if ($PSCmdlet.ShouldProcess($resolvedKey, "Rename AD object to '$newName'")) {
                Invoke-WithRetry -OperationName "Rename AD object $resolvedKey" -ScriptBlock {
                    Rename-ADObject -Identity $targetObject.ObjectGuid -NewName $newName -ErrorAction Stop
                }

                $messages.Add("Object renamed to '$newName'.")
            }
            else {
                $messages.Add('Rename skipped due to WhatIf.')
            }
        }

        $targetPath = Get-TrimmedValue -Value $row.TargetPath
        if (-not [string]::IsNullOrWhiteSpace($targetPath)) {
            $currentObject = Invoke-WithRetry -OperationName "Reload AD object $resolvedKey" -ScriptBlock {
                Get-ADObject -Identity $targetObject.ObjectGuid -Properties DistinguishedName -ErrorAction Stop
            }

            $currentParentPath = ($currentObject.DistinguishedName -split ',', 2)[1]
            if ($targetPath -ieq $currentParentPath) {
                $messages.Add('Object already in requested OU path.')
            }
            else {
                $changeCount++
                if ($PSCmdlet.ShouldProcess($resolvedKey, "Move AD object to '$targetPath'")) {
                    Invoke-WithRetry -OperationName "Move AD object $resolvedKey" -ScriptBlock {
                        Move-ADObject -Identity $targetObject.ObjectGuid -TargetPath $targetPath -ErrorAction Stop
                    }

                    $messages.Add("Object moved to '$targetPath'.")
                }
                else {
                    $messages.Add('Move skipped due to WhatIf.')
                }
            }
        }

        $protectionEnabled = Get-NullableBool -Value $row.ProtectionEnabled
        if ($null -ne $protectionEnabled) {
            $changeCount++
            if ($PSCmdlet.ShouldProcess($resolvedKey, "Set ProtectedFromAccidentalDeletion = $protectionEnabled")) {
                Invoke-WithRetry -OperationName "Set AD object protection $resolvedKey" -ScriptBlock {
                    Set-ADObject -Identity $targetObject.ObjectGuid -ProtectedFromAccidentalDeletion $protectionEnabled -ErrorAction Stop
                }

                $messages.Add("ProtectedFromAccidentalDeletion set to '$protectionEnabled'.")
            }
            else {
                $messages.Add('Protection change skipped due to WhatIf.')
            }
        }

        if ($changeCount -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'MoveActiveDirectoryObject' -Status 'Skipped' -Message 'No changes were requested for this row.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'MoveActiveDirectoryObject' -Status 'Completed' -Message ($messages -join ' ')))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'MoveActiveDirectoryObject' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory object move script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
