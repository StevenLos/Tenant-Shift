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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_P0006-Create-ActiveDirectoryDistributionGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Add-IfValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Hashtable,

        [Parameter(Mandatory)]
        [string]$Key,

        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        $Hashtable[$Key] = $text
    }
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'Name',
    'SamAccountName',
    'DisplayName',
    'Description',
    'GroupScope',
    'Path',
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
    'ExtensionAttribute15'
)

Write-Status -Message 'Starting Active Directory distribution group creation script.'
Ensure-ActiveDirectoryConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $name = Get-TrimmedValue -Value $row.Name
    $samAccountName = Get-TrimmedValue -Value $row.SamAccountName
    $path = Get-TrimmedValue -Value $row.Path
    $primaryKey = if (-not [string]::IsNullOrWhiteSpace($samAccountName)) { $samAccountName } else { $name }

    try {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Name is required.'
        }

        if ([string]::IsNullOrWhiteSpace($samAccountName)) {
            $samAccountName = $name
            $primaryKey = $samAccountName
        }

        if ([string]::IsNullOrWhiteSpace($path)) {
            throw 'Path (target OU distinguished name) is required.'
        }

        $groupScope = Get-TrimmedValue -Value $row.GroupScope
        if ([string]::IsNullOrWhiteSpace($groupScope)) {
            $groupScope = 'Global'
        }

        if ($groupScope -notin @('DomainLocal', 'Global', 'Universal')) {
            throw "GroupScope '$groupScope' is invalid. Use DomainLocal, Global, or Universal."
        }

        $escapedSam = Escape-AdFilterValue -Value $samAccountName
        $existingBySam = Invoke-WithRetry -OperationName "Lookup group by SamAccountName $samAccountName" -ScriptBlock {
            Get-ADGroup -Filter "SamAccountName -eq '$escapedSam'" -ErrorAction SilentlyContinue | Select-Object -First 1
        }

        if ($existingBySam) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryDistributionGroup' -Status 'Skipped' -Message "Group already exists as '$($existingBySam.DistinguishedName)'."))
            $rowNumber++
            continue
        }

        $newGroupParams = @{
            Name          = $name
            SamAccountName = $samAccountName
            GroupCategory = 'Distribution'
            GroupScope    = $groupScope
            Path          = $path
        }

        Add-IfValue -Hashtable $newGroupParams -Key 'DisplayName' -Value $row.DisplayName
        Add-IfValue -Hashtable $newGroupParams -Key 'Description' -Value $row.Description
        Add-IfValue -Hashtable $newGroupParams -Key 'ManagedBy' -Value $row.ManagedBy

        $replaceAttributes = @{}
        Add-IfValue -Hashtable $replaceAttributes -Key 'mail' -Value $row.Mail
        Add-IfValue -Hashtable $replaceAttributes -Key 'mailNickname' -Value $row.MailNickname

        $proxyAddresses = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ProxyAddresses)
        if ($proxyAddresses.Count -gt 0) {
            $replaceAttributes['proxyAddresses'] = $proxyAddresses
        }

        $hideFromAddressLists = Get-NullableBool -Value $row.HideFromAddressLists
        if ($null -ne $hideFromAddressLists) {
            $replaceAttributes['msExchHideFromAddressLists'] = $hideFromAddressLists
        }

        for ($i = 1; $i -le 15; $i++) {
            $columnName = "ExtensionAttribute$i"
            $attributeName = "extensionAttribute$i"
            Add-IfValue -Hashtable $replaceAttributes -Key $attributeName -Value $row.$columnName
        }

        if ($PSCmdlet.ShouldProcess($primaryKey, 'Create Active Directory distribution group')) {
            $createdGroup = Invoke-WithRetry -OperationName "Create AD security group $primaryKey" -ScriptBlock {
                New-ADGroup @newGroupParams -PassThru -ErrorAction Stop
            }

            if ($replaceAttributes.Count -gt 0) {
                Invoke-WithRetry -OperationName "Set AD security group attributes $primaryKey" -ScriptBlock {
                    Set-ADGroup -Identity $createdGroup.DistinguishedName -Replace $replaceAttributes -ErrorAction Stop
                }
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryDistributionGroup' -Status 'Created' -Message 'Group created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryDistributionGroup' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryDistributionGroup' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory distribution group creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
