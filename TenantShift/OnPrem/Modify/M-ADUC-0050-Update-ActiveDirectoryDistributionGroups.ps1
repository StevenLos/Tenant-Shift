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

.SYNOPSIS
    Modifies ActiveDirectoryDistributionGroups in Active Directory.

.DESCRIPTION
    Updates ActiveDirectoryDistributionGroups in Active Directory based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M0006-Update-ActiveDirectoryDistributionGroups.ps1 -InputCsvPath .\0006.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M0006-Update-ActiveDirectoryDistributionGroups.ps1 -InputCsvPath .\0006.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: ActiveDirectory
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    Column                Type      Required  Description
    --------------------  ----      --------  -----------
    Action                String    Yes       <fill in description>
    Notes                 String    Yes       <fill in description>
    IdentityType          String    Yes       <fill in description>
    IdentityValue         String    Yes       <fill in description>
    ClearAttributes       String    Yes       <fill in description>
    Name                  String    Yes       <fill in description>
    SamAccountName        String    Yes       <fill in description>
    DisplayName           String    Yes       <fill in description>
    Description           String    Yes       <fill in description>
    GroupScope            String    Yes       <fill in description>
    ManagedBy             String    Yes       <fill in description>
    Mail                  String    Yes       <fill in description>
    MailNickname          String    Yes       <fill in description>
    ProxyAddresses        String    Yes       <fill in description>
    HideFromAddressLists  String    Yes       <fill in description>
    ExtensionAttribute1   String    Yes       <fill in description>
    ExtensionAttribute2   String    Yes       <fill in description>
    ExtensionAttribute3   String    Yes       <fill in description>
    ExtensionAttribute4   String    Yes       <fill in description>
    ExtensionAttribute5   String    Yes       <fill in description>
    ExtensionAttribute6   String    Yes       <fill in description>
    ExtensionAttribute7   String    Yes       <fill in description>
    ExtensionAttribute8   String    Yes       <fill in description>
    ExtensionAttribute9   String    Yes       <fill in description>
    ExtensionAttribute10  String    Yes       <fill in description>
    ExtensionAttribute11  String    Yes       <fill in description>
    ExtensionAttribute12  String    Yes       <fill in description>
    ExtensionAttribute13  String    Yes       <fill in description>
    ExtensionAttribute14  String    Yes       <fill in description>
    ExtensionAttribute15  String    Yes       <fill in description>
    Path                  String    Yes       <fill in description>
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M0006-Update-ActiveDirectoryDistributionGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
            return Get-ADGroup -Filter "SamAccountName -eq '$escaped'" -Properties * -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'name' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADGroup -Filter "Name -eq '$escaped'" -Properties * -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'distinguishedname' {
            return Get-ADGroup -Identity $IdentityValue -Properties * -ErrorAction SilentlyContinue
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            return Get-ADGroup -Identity $guid -Properties * -ErrorAction SilentlyContinue
        }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use SamAccountName, Name, DistinguishedName, or ObjectGuid."
        }
    }
}

function Add-SetGroupField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$SetParams,

        [Parameter(Mandatory)]
        [string]$ParameterName,

        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        $SetParams[$ParameterName] = $text
        return $true
    }

    return $false
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'IdentityType',
    'IdentityValue',
    'ClearAttributes',
    'Name',
    'SamAccountName',
    'DisplayName',
    'Description',
    'GroupScope',
    'ManagedBy',
    'Mail',
    'MailNickname',
    'ProxyAddresses',
    'HideFromAddressLists',
    'ExtensionAttribute1',
    'ExtensionAttribute2',
    'ExtensionAttribute3',
    'ExtensionAttribute4',
    'ExtensionAttribute5',
    'ExtensionAttribute6',
    'ExtensionAttribute7',
    'ExtensionAttribute8',
    'ExtensionAttribute9',
    'ExtensionAttribute10',
    'ExtensionAttribute11',
    'ExtensionAttribute12',
    'ExtensionAttribute13',
    'ExtensionAttribute14',
    'ExtensionAttribute15',
    'Path'
)

$clearAttributeMap = @{
    SamAccountName = 'sAMAccountName'
    DisplayName = 'displayName'
    Description = 'description'
    ManagedBy = 'managedBy'
    Mail = 'mail'
    MailNickname = 'mailNickname'
    ProxyAddresses = 'proxyAddresses'
    HideFromAddressLists = 'msExchHideFromAddressLists'
}

for ($i = 1; $i -le 15; $i++) {
    $clearAttributeMap["ExtensionAttribute$i"] = "extensionAttribute$i"
}

Write-Status -Message 'Starting Active Directory distribution group update script.'
Ensure-ActiveDirectoryConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $identityType = Get-TrimmedValue -Value $row.IdentityType
    $identityValue = Get-TrimmedValue -Value $row.IdentityValue
    $primaryKey = "${identityType}:$identityValue"

    try {
        if ([string]::IsNullOrWhiteSpace($identityType) -or [string]::IsNullOrWhiteSpace($identityValue)) {
            throw 'IdentityType and IdentityValue are required.'
        }

        $targetGroup = Invoke-WithRetry -OperationName "Resolve AD group $primaryKey" -ScriptBlock {
            Resolve-TargetAdGroup -IdentityType $identityType -IdentityValue $identityValue
        }

        if (-not $targetGroup) {
            throw 'Target group was not found.'
        }

        if ((Get-TrimmedValue -Value $targetGroup.GroupCategory) -ne $expectedGroupCategory) {
            throw "Target group category '$($targetGroup.GroupCategory)' does not match expected '$expectedGroupCategory'."
        }

        $resolvedKey = if (-not [string]::IsNullOrWhiteSpace((Get-TrimmedValue -Value $targetGroup.SamAccountName))) { (Get-TrimmedValue -Value $targetGroup.SamAccountName) } else { (Get-TrimmedValue -Value $targetGroup.Name) }

        $messages = [System.Collections.Generic.List[string]]::new()
        $changeCount = 0

        $setParams = @{
            Identity = $targetGroup.ObjectGuid
        }

        if (Add-SetGroupField -SetParams $setParams -ParameterName 'SamAccountName' -Value $row.SamAccountName) { $changeCount++ }
        if (Add-SetGroupField -SetParams $setParams -ParameterName 'DisplayName' -Value $row.DisplayName) { $changeCount++ }
        if (Add-SetGroupField -SetParams $setParams -ParameterName 'Description' -Value $row.Description) { $changeCount++ }
        if (Add-SetGroupField -SetParams $setParams -ParameterName 'ManagedBy' -Value $row.ManagedBy) { $changeCount++ }

        $groupScope = Get-TrimmedValue -Value $row.GroupScope
        if (-not [string]::IsNullOrWhiteSpace($groupScope)) {
            if ($groupScope -notin @('DomainLocal', 'Global', 'Universal')) {
                throw "GroupScope '$groupScope' is invalid. Use DomainLocal, Global, or Universal."
            }

            $setParams['GroupScope'] = $groupScope
            $changeCount++
        }

        $replaceAttributes = @{}
        Add-SetGroupField -SetParams $replaceAttributes -ParameterName 'mail' -Value $row.Mail | Out-Null
        Add-SetGroupField -SetParams $replaceAttributes -ParameterName 'mailNickname' -Value $row.MailNickname | Out-Null

        $proxyAddresses = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ProxyAddresses)
        if ($proxyAddresses.Count -gt 0) {
            $replaceAttributes['proxyAddresses'] = [string[]]$proxyAddresses
            $changeCount++
        }

        $hideFromAddressLists = Get-NullableBool -Value $row.HideFromAddressLists
        if ($null -ne $hideFromAddressLists) {
            $replaceAttributes['msExchHideFromAddressLists'] = $hideFromAddressLists
            $changeCount++
        }

        for ($i = 1; $i -le 15; $i++) {
            $columnName = "ExtensionAttribute$i"
            $attributeName = "extensionAttribute$i"
            $value = Get-TrimmedValue -Value $row.$columnName
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $replaceAttributes[$attributeName] = $value
                $changeCount++
            }
        }

        if (-not [string]::IsNullOrWhiteSpace((Get-TrimmedValue -Value $row.Mail))) { $changeCount++ }
        if (-not [string]::IsNullOrWhiteSpace((Get-TrimmedValue -Value $row.MailNickname))) { $changeCount++ }

        if ($replaceAttributes.Count -gt 0) {
            $setParams['Replace'] = $replaceAttributes
        }

        $clearAttributes = [System.Collections.Generic.List[string]]::new()
        $clearRequested = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ClearAttributes)
        foreach ($clearRequestedName in $clearRequested) {
            if ($clearAttributeMap.ContainsKey($clearRequestedName)) {
                $mapped = $clearAttributeMap[$clearRequestedName]
                if (-not $clearAttributes.Contains($mapped)) {
                    $clearAttributes.Add($mapped)
                }
            }
            else {
                if (-not $clearAttributes.Contains($clearRequestedName)) {
                    $clearAttributes.Add($clearRequestedName)
                }
            }
        }

        if ($clearAttributes.Count -gt 0) {
            $setParams['Clear'] = @($clearAttributes)
            $changeCount++
        }

        if ($setParams.Count -gt 1) {
            if ($PSCmdlet.ShouldProcess($resolvedKey, 'Update Active Directory distribution group attributes')) {
                Invoke-WithRetry -OperationName "Update AD distribution group attributes $resolvedKey" -ScriptBlock {
                    Set-ADGroup @setParams -ErrorAction Stop
                }

                $messages.Add('Attributes updated.')
            }
            else {
                $messages.Add('Attribute updates skipped due to WhatIf.')
            }
        }

        $newName = Get-TrimmedValue -Value $row.Name
        if (-not [string]::IsNullOrWhiteSpace($newName) -and $newName -ne (Get-TrimmedValue -Value $targetGroup.Name)) {
            $changeCount++

            if ($PSCmdlet.ShouldProcess($resolvedKey, "Rename AD group to '$newName'")) {
                Invoke-WithRetry -OperationName "Rename AD group $resolvedKey" -ScriptBlock {
                    Rename-ADObject -Identity $targetGroup.ObjectGuid -NewName $newName -ErrorAction Stop
                }

                $messages.Add("Group renamed to '$newName'.")
            }
            else {
                $messages.Add('Rename skipped due to WhatIf.')
            }
        }

        $targetPath = Get-TrimmedValue -Value $row.Path
        if (-not [string]::IsNullOrWhiteSpace($targetPath)) {
            $currentGroup = Invoke-WithRetry -OperationName "Reload AD group $resolvedKey" -ScriptBlock {
                Get-ADGroup -Identity $targetGroup.ObjectGuid -Properties DistinguishedName -ErrorAction Stop
            }

            $currentParentPath = ($currentGroup.DistinguishedName -split ',', 2)[1]
            if ($targetPath -ieq $currentParentPath) {
                $messages.Add('Group already in requested OU path.')
            }
            else {
                $changeCount++
                if ($PSCmdlet.ShouldProcess($resolvedKey, "Move AD group to '$targetPath'")) {
                    Invoke-WithRetry -OperationName "Move AD group $resolvedKey" -ScriptBlock {
                        Move-ADObject -Identity $targetGroup.ObjectGuid -TargetPath $targetPath -ErrorAction Stop
                    }

                    $messages.Add("Group moved to '$targetPath'.")
                }
                else {
                    $messages.Add('OU move skipped due to WhatIf.')
                }
            }
        }

        if ($changeCount -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'UpdateActiveDirectoryDistributionGroup' -Status 'Skipped' -Message 'No changes were requested for this row.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'UpdateActiveDirectoryDistributionGroup' -Status 'Completed' -Message ($messages -join ' ')))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'UpdateActiveDirectoryDistributionGroup' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory distribution group update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
