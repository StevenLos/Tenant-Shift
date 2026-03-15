<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-141500

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3009-Set-EntraGroupCreatorsPolicy_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-NullableBool {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return (ConvertTo-Bool -Value $text)
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

function Convert-SettingValuesToMap {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$Values
    )

    $map = @{}

    foreach ($entry in @($Values)) {
        $name = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $entry -PropertyName 'name')
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $map[$name] = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $entry -PropertyName 'value')
    }

    return $map
}

function Convert-SettingMapToValues {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Map
    )

    $values = [System.Collections.Generic.List[object]]::new()

    foreach ($key in @($Map.Keys | Sort-Object)) {
        $values.Add([ordered]@{
                name  = [string]$key
                value = [string]$Map[$key]
            })
    }

    return $values.ToArray()
}

function Get-GroupUnifiedSetting {
    [CmdletBinding()]
    param()

    $response = Invoke-WithRetry -OperationName 'Load Group.Unified setting' -ScriptBlock {
        Invoke-MgGraphRequest -Method GET -Uri "/v1.0/groupSettings?`$filter=displayName eq 'Group.Unified'" -OutputType PSObject -ErrorAction Stop
    }

    $settings = @((Get-GraphPropertyValue -Object $response -PropertyName 'value'))
    if ($settings.Count -eq 0) {
        return $null
    }

    if ($settings.Count -gt 1) {
        Write-Status -Message 'Multiple Group.Unified settings found. The first setting object will be used.' -Level WARN
    }

    return $settings[0]
}

function Get-GroupUnifiedTemplate {
    [CmdletBinding()]
    param()

    $response = Invoke-WithRetry -OperationName 'Load Group.Unified template' -ScriptBlock {
        Invoke-MgGraphRequest -Method GET -Uri "/v1.0/groupSettingTemplates?`$filter=displayName eq 'Group.Unified'" -OutputType PSObject -ErrorAction Stop
    }

    $templates = @((Get-GraphPropertyValue -Object $response -PropertyName 'value'))
    if ($templates.Count -eq 0) {
        throw "Unable to locate the 'Group.Unified' setting template in Microsoft Graph."
    }

    return $templates[0]
}

function Resolve-AllowedGroup {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$IdentityType,

        [AllowEmptyString()]
        [string]$IdentityValue
    )

    $resolvedType = Get-TrimmedValue -Value $IdentityType
    $resolvedValue = Get-TrimmedValue -Value $IdentityValue

    if ([string]::IsNullOrWhiteSpace($resolvedType) -and [string]::IsNullOrWhiteSpace($resolvedValue)) {
        return [PSCustomObject]@{
            Id           = ''
            DisplayName  = ''
            MailNickname = ''
        }
    }

    if ([string]::IsNullOrWhiteSpace($resolvedType) -or [string]::IsNullOrWhiteSpace($resolvedValue)) {
        throw 'AllowedGroupIdentityType and AllowedGroupIdentityValue must both be set when resolving an allowed group.'
    }

    switch ($resolvedType.Trim().ToLowerInvariant()) {
        'groupid' {
            $group = Invoke-WithRetry -OperationName "Lookup allowed group by id $resolvedValue" -ScriptBlock {
                Get-MgGroup -GroupId $resolvedValue -Property 'id,displayName,mailNickname' -ErrorAction SilentlyContinue
            }

            if (-not $group) {
                throw "Allowed group '$resolvedValue' was not found."
            }

            return [PSCustomObject]@{
                Id           = Get-TrimmedValue -Value $group.Id
                DisplayName  = Get-TrimmedValue -Value $group.DisplayName
                MailNickname = Get-TrimmedValue -Value $group.MailNickname
            }
        }
        'displayname' {
            $escaped = Escape-ODataString -Value $resolvedValue
            $groups = @(Invoke-WithRetry -OperationName "Lookup allowed group by displayName $resolvedValue" -ScriptBlock {
                    Get-MgGroup -Filter "displayName eq '$escaped'" -ConsistencyLevel eventual -Property 'id,displayName,mailNickname' -ErrorAction Stop
                })

            if ($groups.Count -eq 0) {
                throw "Allowed group displayName '$resolvedValue' was not found."
            }

            if ($groups.Count -gt 1) {
                throw "Multiple groups matched displayName '$resolvedValue'. Use AllowedGroupIdentityType='GroupId' for a unique match."
            }

            $group = $groups[0]
            return [PSCustomObject]@{
                Id           = Get-TrimmedValue -Value $group.Id
                DisplayName  = Get-TrimmedValue -Value $group.DisplayName
                MailNickname = Get-TrimmedValue -Value $group.MailNickname
            }
        }
        'mailnickname' {
            $escaped = Escape-ODataString -Value $resolvedValue
            $groups = @(Invoke-WithRetry -OperationName "Lookup allowed group by mailNickname $resolvedValue" -ScriptBlock {
                    Get-MgGroup -Filter "mailNickname eq '$escaped'" -ConsistencyLevel eventual -Property 'id,displayName,mailNickname' -ErrorAction Stop
                })

            if ($groups.Count -eq 0) {
                throw "Allowed group mailNickname '$resolvedValue' was not found."
            }

            if ($groups.Count -gt 1) {
                throw "Multiple groups matched mailNickname '$resolvedValue'. Use AllowedGroupIdentityType='GroupId' for a unique match."
            }

            $group = $groups[0]
            return [PSCustomObject]@{
                Id           = Get-TrimmedValue -Value $group.Id
                DisplayName  = Get-TrimmedValue -Value $group.DisplayName
                MailNickname = Get-TrimmedValue -Value $group.MailNickname
            }
        }
        default {
            throw "AllowedGroupIdentityType '$resolvedType' is invalid. Use GroupId, DisplayName, or MailNickname."
        }
    }
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'AllowGroupCreation',
    'AllowedGroupIdentityType',
    'AllowedGroupIdentityValue',
    'ClearAllowedGroup'
)

Write-Status -Message 'Starting Entra group creators policy update script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')
Ensure-GraphConnection -RequiredScopes @('Directory.ReadWrite.All', 'Group.ReadWrite.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $policyKey = 'Group.Unified'

    try {
        $allowGroupCreation = Get-NullableBool -Value $row.AllowGroupCreation
        $allowedGroupIdentityType = Get-TrimmedValue -Value $row.AllowedGroupIdentityType
        $allowedGroupIdentityValue = Get-TrimmedValue -Value $row.AllowedGroupIdentityValue
        $clearAllowedGroup = Get-NullableBool -Value $row.ClearAllowedGroup

        if ($null -eq $clearAllowedGroup) {
            $clearAllowedGroup = $false
        }

        if ($clearAllowedGroup -and (-not [string]::IsNullOrWhiteSpace($allowedGroupIdentityType) -or -not [string]::IsNullOrWhiteSpace($allowedGroupIdentityValue))) {
            throw 'ClearAllowedGroup cannot be TRUE when an allowed group identity is also supplied.'
        }

        $resolvedAllowedGroup = Resolve-AllowedGroup -IdentityType $allowedGroupIdentityType -IdentityValue $allowedGroupIdentityValue
        $hasRequestedChange = ($null -ne $allowGroupCreation) -or $clearAllowedGroup -or (-not [string]::IsNullOrWhiteSpace($resolvedAllowedGroup.Id))

        if (-not $hasRequestedChange) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $policyKey -Action 'SetEntraGroupCreatorsPolicy' -Status 'Skipped' -Message 'No policy updates were requested.'))
            $rowNumber++
            continue
        }

        $setting = Get-GroupUnifiedSetting
        $settingId = ''
        $templateId = ''
        $valuesMap = @{}
        $settingWasCreated = $false

        if ($setting) {
            $settingId = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $setting -PropertyName 'id')
            $valuesMap = Convert-SettingValuesToMap -Values @((Get-GraphPropertyValue -Object $setting -PropertyName 'values'))
        }
        else {
            $template = Get-GroupUnifiedTemplate
            $templateId = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $template -PropertyName 'id')
            $valuesMap = Convert-SettingValuesToMap -Values @((Get-GraphPropertyValue -Object $template -PropertyName 'values'))
            $settingWasCreated = $true
        }

        if (-not $valuesMap.ContainsKey('EnableGroupCreation')) {
            $valuesMap['EnableGroupCreation'] = ''
        }
        if (-not $valuesMap.ContainsKey('GroupCreationAllowedGroupId')) {
            $valuesMap['GroupCreationAllowedGroupId'] = ''
        }

        $allowGroupCreationBefore = Get-TrimmedValue -Value $valuesMap['EnableGroupCreation']
        $allowedGroupIdBefore = Get-TrimmedValue -Value $valuesMap['GroupCreationAllowedGroupId']

        if ($null -ne $allowGroupCreation) {
            $valuesMap['EnableGroupCreation'] = if ($allowGroupCreation) { 'true' } else { 'false' }
        }

        if ($clearAllowedGroup) {
            $valuesMap['GroupCreationAllowedGroupId'] = ''
        }
        elseif (-not [string]::IsNullOrWhiteSpace($resolvedAllowedGroup.Id)) {
            $valuesMap['GroupCreationAllowedGroupId'] = $resolvedAllowedGroup.Id
        }

        $allowGroupCreationAfter = Get-TrimmedValue -Value $valuesMap['EnableGroupCreation']
        $allowedGroupIdAfter = Get-TrimmedValue -Value $valuesMap['GroupCreationAllowedGroupId']

        $isValueChange = ($allowGroupCreationBefore -ne $allowGroupCreationAfter) -or ($allowedGroupIdBefore -ne $allowedGroupIdAfter)
        if ((-not $isValueChange) -and (-not $settingWasCreated)) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $policyKey -Action 'SetEntraGroupCreatorsPolicy' -Status 'Skipped' -Message 'Group creators policy is already set to the requested values.'))
            $rowNumber++
            continue
        }

        $targetDescription = if ($settingWasCreated) { 'Create Group.Unified setting and apply group creators policy' } else { 'Update Group.Unified group creators policy' }
        if ($PSCmdlet.ShouldProcess($policyKey, $targetDescription)) {
            $bodyObject = @{
                values = Convert-SettingMapToValues -Map $valuesMap
            }

            if ($settingWasCreated) {
                $bodyObject['templateId'] = $templateId
                $createBody = $bodyObject | ConvertTo-Json -Depth 8 -Compress

                $createdSetting = Invoke-WithRetry -OperationName 'Create Group.Unified setting' -ScriptBlock {
                    Invoke-MgGraphRequest -Method POST -Uri '/v1.0/groupSettings' -Body $createBody -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
                }

                $settingId = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $createdSetting -PropertyName 'id')
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $policyKey -Action 'SetEntraGroupCreatorsPolicy' -Status 'Created' -Message 'Group creators policy setting created and updated.'))
            }
            else {
                $updateBody = $bodyObject | ConvertTo-Json -Depth 8 -Compress

                Invoke-WithRetry -OperationName 'Update Group.Unified setting' -ScriptBlock {
                    Invoke-MgGraphRequest -Method PATCH -Uri "/v1.0/groupSettings/$settingId" -Body $updateBody -ContentType 'application/json' -ErrorAction Stop | Out-Null
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $policyKey -Action 'SetEntraGroupCreatorsPolicy' -Status 'Updated' -Message 'Group creators policy updated successfully.'))
            }
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $policyKey -Action 'SetEntraGroupCreatorsPolicy' -Status 'WhatIf' -Message 'Policy update skipped due to WhatIf.'))
        }

        $resultIndex = $results.Count - 1
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'SettingId' -NotePropertyValue $settingId -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'AllowGroupCreationBefore' -NotePropertyValue $allowGroupCreationBefore -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'AllowGroupCreationAfter' -NotePropertyValue $allowGroupCreationAfter -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'AllowedGroupIdBefore' -NotePropertyValue $allowedGroupIdBefore -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'AllowedGroupIdAfter' -NotePropertyValue $allowedGroupIdAfter -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'AllowedGroupIdentityTypeRequested' -NotePropertyValue $allowedGroupIdentityType -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'AllowedGroupIdentityValueRequested' -NotePropertyValue $allowedGroupIdentityValue -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'ResolvedAllowedGroupDisplayName' -NotePropertyValue $resolvedAllowedGroup.DisplayName -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'ResolvedAllowedGroupMailNickname' -NotePropertyValue $resolvedAllowedGroup.MailNickname -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'SettingWasCreated' -NotePropertyValue ([string]$settingWasCreated) -Force
    }
    catch {
        Write-Status -Message "Row $rowNumber ($policyKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $policyKey -Action 'SetEntraGroupCreatorsPolicy' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

$extendedColumns = @(
    'SettingId',
    'AllowGroupCreationBefore',
    'AllowGroupCreationAfter',
    'AllowedGroupIdBefore',
    'AllowedGroupIdAfter',
    'AllowedGroupIdentityTypeRequested',
    'AllowedGroupIdentityValueRequested',
    'ResolvedAllowedGroupDisplayName',
    'ResolvedAllowedGroupMailNickname',
    'SettingWasCreated'
)

foreach ($result in $results) {
    foreach ($column in $extendedColumns) {
        if ($result.PSObject.Properties.Name -notcontains $column) {
            Add-Member -InputObject $result -NotePropertyName $column -NotePropertyValue '' -Force
        }
    }
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra group creators policy update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
