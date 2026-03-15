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
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3008-Update-EntraMicrosoft365Groups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Test-IsMicrosoft365Group {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Group
    )

    $groupTypes = @($Group.GroupTypes)
    return (($groupTypes -contains 'Unified') -and ($Group.MailEnabled -eq $true) -and ($Group.SecurityEnabled -eq $false))
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'GroupMailNickname',
    'GroupDisplayName',
    'Description',
    'MailNickname',
    'Visibility',
    'ClearAttributes'
)

Write-Status -Message 'Starting Entra Microsoft 365 group update script.'
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
            Get-MgGroup -Filter "mailNickname eq '$escapedAlias'" -ConsistencyLevel eventual -Property 'id,displayName,description,mailNickname,visibility,groupTypes,securityEnabled,mailEnabled' -ErrorAction Stop
        })

        if ($groups.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateMicrosoft365Group' -Status 'NotFound' -Message 'Group not found.'))
            $rowNumber++
            continue
        }

        if ($groups.Count -gt 1) {
            throw "Multiple groups found with mailNickname '$groupMailNickname'. Resolve duplicate aliases before retrying."
        }

        $group = $groups[0]
        if (-not (Test-IsMicrosoft365Group -Group $group)) {
            throw "Group '$groupMailNickname' is not a Microsoft 365 group."
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

        $visibility = Get-TrimmedValue -Value $row.Visibility
        if (-not [string]::IsNullOrWhiteSpace($visibility)) {
            if ($visibility -notin @('Private', 'Public')) {
                throw "Visibility '$visibility' is invalid. Use Private or Public."
            }

            $body['visibility'] = $visibility
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
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateMicrosoft365Group' -Status 'Skipped' -Message 'No updates were requested.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($groupMailNickname, 'Update Microsoft 365 group attributes')) {
            Invoke-WithRetry -OperationName "Update Microsoft 365 group $groupMailNickname" -ScriptBlock {
                Update-MgGroup -GroupId $group.Id -BodyParameter $body -ErrorAction Stop | Out-Null
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateMicrosoft365Group' -Status 'Updated' -Message 'Microsoft 365 group updated successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateMicrosoft365Group' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupMailNickname) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'UpdateMicrosoft365Group' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra Microsoft 365 group update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
