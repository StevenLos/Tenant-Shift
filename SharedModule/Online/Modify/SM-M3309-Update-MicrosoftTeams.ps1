<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-210000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Groups
Microsoft.Graph.Teams

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3309-Update-MicrosoftTeams_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Resolve-TeamByAlias {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TeamMailNickname,

        [Parameter(Mandatory)]
        [System.Collections.Generic.Dictionary[string, object]]$TeamByAlias
    )

    if ($TeamByAlias.ContainsKey($TeamMailNickname)) {
        return $TeamByAlias[$TeamMailNickname]
    }

    $escapedAlias = Escape-ODataString -Value $TeamMailNickname
    $groups = @(Invoke-WithRetry -OperationName "Lookup Team group alias $TeamMailNickname" -ScriptBlock {
        Get-MgGroup -Filter "mailNickname eq '$escapedAlias'" -ConsistencyLevel eventual -ErrorAction Stop
    })

    if ($groups.Count -eq 0) {
        throw "No group was found with mailNickname '$TeamMailNickname'."
    }

    if ($groups.Count -gt 1) {
        throw "Multiple groups were found with mailNickname '$TeamMailNickname'. Resolve duplicate aliases before running this script."
    }

    $group = Invoke-WithRetry -OperationName "Load Team group details for alias $TeamMailNickname" -ScriptBlock {
        Get-MgGroup -GroupId $groups[0].Id -Property 'id,displayName,description,mailNickname,visibility,groupTypes,securityEnabled,mailEnabled' -ErrorAction Stop
    }

    $groupTypes = @($group.GroupTypes)
    $isMicrosoft365Group = ($groupTypes -contains 'Unified') -and ($group.MailEnabled -eq $true) -and ($group.SecurityEnabled -eq $false)
    if (-not $isMicrosoft365Group) {
        throw "Group '$TeamMailNickname' exists but is not a Microsoft 365 group."
    }

    $team = Invoke-WithRetry -OperationName "Verify Team exists for alias $TeamMailNickname" -ScriptBlock {
        Get-MgGroupTeam -GroupId $group.Id -ErrorAction SilentlyContinue
    }
    if (-not $team) {
        throw "Microsoft 365 group '$TeamMailNickname' does not currently have a Team."
    }

    $TeamByAlias[$TeamMailNickname] = $group
    return $group
}

function Add-TeamBooleanSetting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Settings,

        [Parameter(Mandatory)]
        [string]$FieldValue,

        [Parameter(Mandatory)]
        [string]$BodyPropertyName
    )

    $trimmed = ([string]$FieldValue).Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return $false
    }

    $Settings[$BodyPropertyName] = ConvertTo-Bool -Value $trimmed
    return $true
}

$requiredHeaders = @(
    'TeamMailNickname',
    'TeamDisplayName',
    'Description',
    'Visibility',
    'AllowCreateUpdateChannels',
    'AllowDeleteChannels',
    'AllowAddRemoveApps',
    'AllowCreateUpdateRemoveTabs',
    'AllowCreateUpdateRemoveConnectors',
    'AllowUserEditMessages',
    'AllowUserDeleteMessages',
    'AllowOwnerDeleteMessages',
    'AllowTeamMentions',
    'AllowChannelMentions',
    'AllowGiphy',
    'GiphyContentRating',
    'AllowStickersAndMemes',
    'AllowCustomMemes',
    'ClearAttributes'
)

Write-Status -Message 'Starting Microsoft Teams update script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Teams')
Ensure-GraphConnection -RequiredScopes @('Group.ReadWrite.All', 'TeamSettings.ReadWrite.All', 'Team.ReadBasic.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$teamByAlias = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

$rowNumber = 1
foreach ($row in $rows) {
    $teamMailNickname = Get-TrimmedValue -Value $row.TeamMailNickname

    try {
        if ([string]::IsNullOrWhiteSpace($teamMailNickname)) {
            throw 'TeamMailNickname is required.'
        }

        $group = Resolve-TeamByAlias -TeamMailNickname $teamMailNickname -TeamByAlias $teamByAlias
        $groupId = ([string]$group.Id).Trim()

        $groupUpdateBody = @{}
        $groupChangeSummary = [System.Collections.Generic.List[string]]::new()

        $teamDisplayName = Get-TrimmedValue -Value $row.TeamDisplayName
        if (-not [string]::IsNullOrWhiteSpace($teamDisplayName)) {
            $groupUpdateBody['displayName'] = $teamDisplayName
            $groupChangeSummary.Add('TeamDisplayName')
        }

        $description = Get-TrimmedValue -Value $row.Description
        if (-not [string]::IsNullOrWhiteSpace($description)) {
            $groupUpdateBody['description'] = $description
            $groupChangeSummary.Add('Description')
        }

        $visibility = Get-TrimmedValue -Value $row.Visibility
        if (-not [string]::IsNullOrWhiteSpace($visibility)) {
            if ($visibility -notin @('Private', 'Public')) {
                throw "Visibility '$visibility' is invalid. Use Private or Public."
            }

            $groupUpdateBody['visibility'] = $visibility
            $groupChangeSummary.Add('Visibility')
        }

        $clearRequested = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ClearAttributes)
        foreach ($clearName in $clearRequested) {
            if ($clearName -eq 'Description') {
                if ($groupUpdateBody.ContainsKey('description')) {
                    throw 'Description cannot be set and cleared in the same row.'
                }

                $groupUpdateBody['description'] = $null
                $groupChangeSummary.Add('Description (Cleared)')
                continue
            }

            throw "ClearAttributes value '$clearName' is not supported."
        }

        $teamUpdateBody = @{}
        $teamChangeSummary = [System.Collections.Generic.List[string]]::new()

        $memberSettings = @{}
        $memberSettingsChanged = $false
        $memberSettingsChanged = (Add-TeamBooleanSetting -Settings $memberSettings -FieldValue $row.AllowCreateUpdateChannels -BodyPropertyName 'allowCreateUpdateChannels') -or $memberSettingsChanged
        $memberSettingsChanged = (Add-TeamBooleanSetting -Settings $memberSettings -FieldValue $row.AllowDeleteChannels -BodyPropertyName 'allowDeleteChannels') -or $memberSettingsChanged
        $memberSettingsChanged = (Add-TeamBooleanSetting -Settings $memberSettings -FieldValue $row.AllowAddRemoveApps -BodyPropertyName 'allowAddRemoveApps') -or $memberSettingsChanged
        $memberSettingsChanged = (Add-TeamBooleanSetting -Settings $memberSettings -FieldValue $row.AllowCreateUpdateRemoveTabs -BodyPropertyName 'allowCreateUpdateRemoveTabs') -or $memberSettingsChanged
        $memberSettingsChanged = (Add-TeamBooleanSetting -Settings $memberSettings -FieldValue $row.AllowCreateUpdateRemoveConnectors -BodyPropertyName 'allowCreateUpdateRemoveConnectors') -or $memberSettingsChanged
        if ($memberSettingsChanged) {
            $teamUpdateBody['memberSettings'] = $memberSettings
            $teamChangeSummary.Add('MemberSettings')
        }

        $messagingSettings = @{}
        $messagingSettingsChanged = $false
        $messagingSettingsChanged = (Add-TeamBooleanSetting -Settings $messagingSettings -FieldValue $row.AllowUserEditMessages -BodyPropertyName 'allowUserEditMessages') -or $messagingSettingsChanged
        $messagingSettingsChanged = (Add-TeamBooleanSetting -Settings $messagingSettings -FieldValue $row.AllowUserDeleteMessages -BodyPropertyName 'allowUserDeleteMessages') -or $messagingSettingsChanged
        $messagingSettingsChanged = (Add-TeamBooleanSetting -Settings $messagingSettings -FieldValue $row.AllowOwnerDeleteMessages -BodyPropertyName 'allowOwnerDeleteMessages') -or $messagingSettingsChanged
        $messagingSettingsChanged = (Add-TeamBooleanSetting -Settings $messagingSettings -FieldValue $row.AllowTeamMentions -BodyPropertyName 'allowTeamMentions') -or $messagingSettingsChanged
        $messagingSettingsChanged = (Add-TeamBooleanSetting -Settings $messagingSettings -FieldValue $row.AllowChannelMentions -BodyPropertyName 'allowChannelMentions') -or $messagingSettingsChanged
        if ($messagingSettingsChanged) {
            $teamUpdateBody['messagingSettings'] = $messagingSettings
            $teamChangeSummary.Add('MessagingSettings')
        }

        $funSettings = @{}
        $funSettingsChanged = $false
        $funSettingsChanged = (Add-TeamBooleanSetting -Settings $funSettings -FieldValue $row.AllowGiphy -BodyPropertyName 'allowGiphy') -or $funSettingsChanged
        $funSettingsChanged = (Add-TeamBooleanSetting -Settings $funSettings -FieldValue $row.AllowStickersAndMemes -BodyPropertyName 'allowStickersAndMemes') -or $funSettingsChanged
        $funSettingsChanged = (Add-TeamBooleanSetting -Settings $funSettings -FieldValue $row.AllowCustomMemes -BodyPropertyName 'allowCustomMemes') -or $funSettingsChanged

        $giphyContentRatingRaw = Get-TrimmedValue -Value $row.GiphyContentRating
        if (-not [string]::IsNullOrWhiteSpace($giphyContentRatingRaw)) {
            $giphyContentRating = switch -Regex ($giphyContentRatingRaw.ToLowerInvariant()) {
                '^strict$' { 'strict'; break }
                '^moderate$' { 'moderate'; break }
                default { throw "GiphyContentRating '$giphyContentRatingRaw' is invalid. Use Strict or Moderate." }
            }

            $funSettings['giphyContentRating'] = $giphyContentRating
            $funSettingsChanged = $true
        }

        if ($funSettingsChanged) {
            $teamUpdateBody['funSettings'] = $funSettings
            $teamChangeSummary.Add('FunSettings')
        }

        if ($groupUpdateBody.Count -eq 0 -and $teamUpdateBody.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $teamMailNickname -Action 'UpdateMicrosoftTeam' -Status 'Skipped' -Message 'No updates were requested.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($teamMailNickname, 'Update Microsoft Team settings')) {
            if ($groupUpdateBody.Count -gt 0) {
                Invoke-WithRetry -OperationName "Update Team group properties $teamMailNickname" -ScriptBlock {
                    Update-MgGroup -GroupId $groupId -BodyParameter $groupUpdateBody -ErrorAction Stop | Out-Null
                }
            }

            if ($teamUpdateBody.Count -gt 0) {
                Invoke-WithRetry -OperationName "Update Team settings $teamMailNickname" -ScriptBlock {
                    Update-MgTeam -TeamId $groupId -BodyParameter $teamUpdateBody -ErrorAction Stop | Out-Null
                }
            }

            $messageParts = [System.Collections.Generic.List[string]]::new()
            if ($groupChangeSummary.Count -gt 0) {
                $messageParts.Add(("Group fields: {0}." -f ($groupChangeSummary -join ', ')))
            }
            if ($teamChangeSummary.Count -gt 0) {
                $messageParts.Add(("Team settings families: {0}." -f ($teamChangeSummary -join ', ')))
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $teamMailNickname -Action 'UpdateMicrosoftTeam' -Status 'Updated' -Message ($messageParts -join ' ')))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $teamMailNickname -Action 'UpdateMicrosoftTeam' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($teamMailNickname) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $teamMailNickname -Action 'UpdateMicrosoftTeam' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Microsoft Teams update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
