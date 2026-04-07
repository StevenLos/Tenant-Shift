<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260406-000000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
PnP.PowerShell

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Configures versioning settings on SharePoint lists and document libraries.

.DESCRIPTION
    For each row in the input CSV, sets versioning options on the specified list or
    document library: enable/disable versioning, major version limit, enable/disable
    minor (draft) versions, and draft visibility. Useful for standardising versioning
    policies before or after migration.
    If ListName is omitted, the settings are applied to all non-hidden lists in the site.
    Reconnects to each site when SiteUrl changes between rows.
    Supports -WhatIf for dry-run validation before committing changes.
    All results are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Required fields: SiteUrl, EnableVersioning.
    See the companion .input.csv template for the full column list.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).
    Used for the initial PnP connection.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\M-SPOL-0100-Set-SharePointListVersioningSettings.ps1 -InputCsvPath .\M-SPOL-0100-Set-SharePointListVersioningSettings.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com -WhatIf

    Dry-run: shows what versioning changes would be applied.

.EXAMPLE
    .\M-SPOL-0100-Set-SharePointListVersioningSettings.ps1 -InputCsvPath .\M-SPOL-0100-Set-SharePointListVersioningSettings.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com

    Apply versioning settings.

.NOTES
    Version:          1.0
    Required modules: PnP.PowerShell
    Required roles:   SharePoint Administrator or Site Owner
    Limitations:      Minor versioning (drafts) applies to document libraries only;
                      setting EnableMinorVersions on a list has no effect and is logged as a warning.
                      MajorVersionLimit of 0 means unlimited.

    CSV Fields:
    Column                  Type      Required  Description
    ----------------------  --------  --------  -----------
    SiteUrl                 String    Yes       Full URL of the SharePoint site collection
    ListName                String    No        Title of the list/library. If blank, applies to all non-hidden lists.
    EnableVersioning        String    Yes       True to enable versioning; False to disable
    MajorVersionLimit       String    No        Maximum number of major versions to keep (0 = unlimited). Default: no change.
    EnableMinorVersions     String    No        True to enable minor (draft) versions (libraries only). Default: no change.
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [ValidatePattern('^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$')]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M-SPOL-0100-Set-SharePointListVersioningSettings_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function New-ModifyResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$RowNumber,
        [Parameter(Mandatory)][string]$PrimaryKey,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$Message,
        [Parameter(Mandatory)][hashtable]$Data
    )

    $base    = New-ResultObject -RowNumber $RowNumber -PrimaryKey $PrimaryKey -Action $Action -Status $Status -Message $Message
    $ordered = [ordered]@{}

    foreach ($prop in $base.PSObject.Properties.Name) {
        $ordered[$prop] = $base.$prop
    }

    foreach ($key in $Data.Keys) {
        $ordered[$key] = $Data[$key]
    }

    return [PSCustomObject]$ordered
}

function Set-ListVersioning {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$List,
        [Parameter(Mandatory)][bool]$EnableVersioning,
        [string]$MajorVersionLimitRaw,
        [string]$EnableMinorVersionsRaw,
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][int]$RowNumber,
        [Parameter(Mandatory)][string]$PrimaryKey
    )

    $listTitle = $list.Title
    $params    = @{ List = $listTitle; EnableVersioning = $EnableVersioning }

    if (-not [string]::IsNullOrWhiteSpace($MajorVersionLimitRaw)) {
        $limit = [int]$MajorVersionLimitRaw
        if ($limit -ge 0) { $params['MajorVersions'] = $limit }
    }

    if (-not [string]::IsNullOrWhiteSpace($EnableMinorVersionsRaw)) {
        $enableMinor = $EnableMinorVersionsRaw.Trim().ToLowerInvariant() -eq 'true'
        # Minor versioning only applies to libraries (BaseType 1).
        if ($list.BaseType -eq 1) {
            $params['EnableMinorVersions'] = $enableMinor
        } else {
            Write-Status -Message "  List '$listTitle' is not a library — EnableMinorVersions ignored." -Level WARN
        }
    }

    Invoke-WithRetry -OperationName "Set versioning on list '$listTitle' in $SiteUrl" -ScriptBlock {
        Set-PnPList @params -ErrorAction Stop | Out-Null
    }

    $appliedParams = ($params.Keys | Where-Object { $_ -ne 'List' } | ForEach-Object { "$_=$($params[$_])" }) -join '; '
    return $appliedParams
}

$requiredHeaders = @('SiteUrl', 'EnableVersioning')

Write-Status -Message 'Starting SharePoint list versioning settings script.'
Assert-ModuleCurrent -ModuleNames @('PnP.PowerShell')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders

$results        = [System.Collections.Generic.List[object]]::new()
$rowNumber      = 1
$currentSiteUrl = ''

foreach ($row in $rows) {
    $siteUrl              = Get-TrimmedValue -Value $row.SiteUrl
    $listName             = if ($row.PSObject.Properties['ListName']) { Get-TrimmedValue -Value $row.ListName } else { '' }
    $enableVersioningRaw  = Get-TrimmedValue -Value $row.EnableVersioning
    $majorVersionLimitRaw = if ($row.PSObject.Properties['MajorVersionLimit']) { Get-TrimmedValue -Value $row.MajorVersionLimit } else { '' }
    $enableMinorRaw       = if ($row.PSObject.Properties['EnableMinorVersions']) { Get-TrimmedValue -Value $row.EnableMinorVersions } else { '' }
    $primaryKey           = if ($listName) { "${siteUrl}|${listName}" } else { "${siteUrl}|*AllLists*" }

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl))            { throw 'SiteUrl is required.' }
        if ([string]::IsNullOrWhiteSpace($enableVersioningRaw)) { throw 'EnableVersioning is required.' }

        if ($enableVersioningRaw.ToLowerInvariant() -notin @('true', 'false')) {
            throw "EnableVersioning '$enableVersioningRaw' is invalid. Valid values: True, False."
        }
        $enableVersioning = $enableVersioningRaw.ToLowerInvariant() -eq 'true'

        $description = "Set versioning (EnableVersioning=$enableVersioning) on $(if ($listName) { "list '$listName'" } else { 'all lists' }) in '$siteUrl'"

        if (-not $PSCmdlet.ShouldProcess($siteUrl, $description)) {
            $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointListVersioning' -Status 'WhatIf' -Message "WhatIf: would $description." -Data ([ordered]@{
                SiteUrl = $siteUrl; ListName = $listName; EnableVersioning = $enableVersioningRaw
                MajorVersionLimit = $majorVersionLimitRaw; EnableMinorVersions = $enableMinorRaw; AppliedSettings = ''; Timestamp = ''
            })))
            $rowNumber++
            continue
        }

        if ($siteUrl -ne $currentSiteUrl) {
            Write-Status -Message "Connecting to site: $siteUrl"
            Connect-PnPOnline -Url $siteUrl -Interactive -ErrorAction Stop
            $currentSiteUrl = $siteUrl
        }

        $listsToProcess = if ($listName) {
            @(Invoke-WithRetry -OperationName "Get list '$listName' from $siteUrl" -ScriptBlock {
                Get-PnPList -Identity $listName -Includes BaseType -ErrorAction Stop
            })
        } else {
            Invoke-WithRetry -OperationName "Get all lists from $siteUrl" -ScriptBlock {
                Get-PnPList -Includes BaseType, Hidden -ErrorAction Stop | Where-Object { -not $_.Hidden }
            }
        }

        foreach ($list in $listsToProcess) {
            try {
                $appliedSettings = Set-ListVersioning -List $list -EnableVersioning $enableVersioning `
                    -MajorVersionLimitRaw $majorVersionLimitRaw -EnableMinorVersionsRaw $enableMinorRaw `
                    -SiteUrl $siteUrl -RowNumber $rowNumber -PrimaryKey $primaryKey

                $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointListVersioning' -Status 'Completed' -Message "Versioning settings applied to list '$($list.Title)' in '$siteUrl'." -Data ([ordered]@{
                    SiteUrl = $siteUrl; ListName = $list.Title; EnableVersioning = $enableVersioningRaw
                    MajorVersionLimit = $majorVersionLimitRaw; EnableMinorVersions = $enableMinorRaw
                    AppliedSettings = $appliedSettings; Timestamp = (Get-Date -Format 'o')
                })))
            }
            catch {
                Write-Status -Message "  List '$($list.Title)' in '$siteUrl' failed: $($_.Exception.Message)" -Level ERROR
                $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey "${siteUrl}|$($list.Title)" -Action 'SetSharePointListVersioning' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    SiteUrl = $siteUrl; ListName = $list.Title; EnableVersioning = $enableVersioningRaw
                    MajorVersionLimit = $majorVersionLimitRaw; EnableMinorVersions = $enableMinorRaw
                    AppliedSettings = ''; Timestamp = ''
                })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointListVersioning' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            SiteUrl = $siteUrl; ListName = $listName; EnableVersioning = $enableVersioningRaw
            MajorVersionLimit = $majorVersionLimitRaw; EnableMinorVersions = $enableMinorRaw
            AppliedSettings = ''; Timestamp = ''
        })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint list versioning settings script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
