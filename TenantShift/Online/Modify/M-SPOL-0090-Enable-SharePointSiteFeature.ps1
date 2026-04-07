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
    Activates or deactivates SharePoint site features by feature ID.

.DESCRIPTION
    For each row in the input CSV, activates or deactivates a specified SharePoint
    site feature at the site or web scope. Useful for enabling Classic list experience
    features, enabling publishing infrastructure, or activating site-level features
    that must be present before migration content is loaded.
    Action column controls direction: Enable or Disable.
    Reconnects to each site when SiteUrl changes between rows.
    Supports -WhatIf for dry-run validation before committing changes.
    All results are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Required fields: SiteUrl, FeatureId, FeatureScope, Action.
    See the companion .input.csv template for the full column list.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).
    Used for the initial PnP connection.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\M-SPOL-0090-Enable-SharePointSiteFeature.ps1 -InputCsvPath .\M-SPOL-0090-Enable-SharePointSiteFeature.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com -WhatIf

    Dry-run: shows what feature activations would be applied.

.EXAMPLE
    .\M-SPOL-0090-Enable-SharePointSiteFeature.ps1 -InputCsvPath .\M-SPOL-0090-Enable-SharePointSiteFeature.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com

    Activate or deactivate the listed features.

.NOTES
    Version:          1.0
    Required modules: PnP.PowerShell
    Required roles:   SharePoint Administrator
    Limitations:      Feature activation may fail if prerequisite features are not already active.
                      Some features cannot be deactivated if they have dependent features active.
                      Rows should be ordered so prerequisite features appear before dependent ones.

    CSV Fields:
    Column          Type      Required  Description
    --------------  --------  --------  -----------
    SiteUrl         String    Yes       Full URL of the SharePoint site collection
    FeatureId       String    Yes       GUID of the feature to activate or deactivate
    FeatureScope    String    Yes       Site (site collection scope) or Web (subweb scope)
    Action          String    Yes       Enable or Disable
    FeatureName     String    No        Human-readable label for logging; not used functionally
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [ValidatePattern('^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$')]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M-SPOL-0090-Enable-SharePointSiteFeature_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$requiredHeaders  = @('SiteUrl', 'FeatureId', 'FeatureScope', 'Action')
$validScopes      = @('Site', 'Web')
$validActions     = @('Enable', 'Disable')

Write-Status -Message 'Starting SharePoint site feature enable/disable script.'
Assert-ModuleCurrent -ModuleNames @('PnP.PowerShell')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders

$results        = [System.Collections.Generic.List[object]]::new()
$rowNumber      = 1
$currentSiteUrl = ''

foreach ($row in $rows) {
    $siteUrl     = Get-TrimmedValue -Value $row.SiteUrl
    $featureId   = Get-TrimmedValue -Value $row.FeatureId
    $featureScope = Get-TrimmedValue -Value $row.FeatureScope
    $action      = Get-TrimmedValue -Value $row.Action
    $featureName = if ($row.PSObject.Properties['FeatureName']) { Get-TrimmedValue -Value $row.FeatureName } else { '' }
    $primaryKey  = "${siteUrl}|${featureId}|${featureScope}"

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl))      { throw 'SiteUrl is required.' }
        if ([string]::IsNullOrWhiteSpace($featureId))    { throw 'FeatureId is required.' }
        if ([string]::IsNullOrWhiteSpace($featureScope)) { throw 'FeatureScope is required.' }
        if ([string]::IsNullOrWhiteSpace($action))       { throw 'Action is required.' }

        # Validate FeatureId is a GUID.
        [System.Guid]::Parse($featureId) | Out-Null

        $normalizedScope  = $validScopes  | Where-Object { $_ -ieq $featureScope }
        $normalizedAction = $validActions | Where-Object { $_ -ieq $action }
        if (-not $normalizedScope)  { throw "FeatureScope '$featureScope' is invalid. Valid values: Site, Web." }
        if (-not $normalizedAction) { throw "Action '$action' is invalid. Valid values: Enable, Disable." }
        $featureScope = $normalizedScope
        $action       = $normalizedAction

        $displayName  = if ($featureName) { "'$featureName' ($featureId)" } else { $featureId }
        $description  = "$action feature $displayName at $featureScope scope on '$siteUrl'"

        if (-not $PSCmdlet.ShouldProcess($siteUrl, $description)) {
            $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'EnableSharePointSiteFeature' -Status 'WhatIf' -Message "WhatIf: would $action feature $displayName ($featureScope scope) on '$siteUrl'." -Data ([ordered]@{
                SiteUrl = $siteUrl; FeatureId = $featureId; FeatureScope = $featureScope; Action = $action; FeatureName = $featureName; Timestamp = ''
            })))
            $rowNumber++
            continue
        }

        if ($siteUrl -ne $currentSiteUrl) {
            Write-Status -Message "Connecting to site: $siteUrl"
            Connect-PnPOnline -Url $siteUrl -Interactive -ErrorAction Stop
            $currentSiteUrl = $siteUrl
        }

        if ($action -eq 'Enable') {
            Invoke-WithRetry -OperationName "Enable feature $displayName on $siteUrl" -ScriptBlock {
                Enable-PnPFeature -Identity $featureId -Scope $featureScope -Force -ErrorAction Stop
            }
            $msg = "Feature $displayName enabled ($featureScope scope) on '$siteUrl'."
        } else {
            Invoke-WithRetry -OperationName "Disable feature $displayName on $siteUrl" -ScriptBlock {
                Disable-PnPFeature -Identity $featureId -Scope $featureScope -Force -ErrorAction Stop
            }
            $msg = "Feature $displayName disabled ($featureScope scope) on '$siteUrl'."
        }

        $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'EnableSharePointSiteFeature' -Status 'Completed' -Message $msg -Data ([ordered]@{
            SiteUrl = $siteUrl; FeatureId = $featureId; FeatureScope = $featureScope; Action = $action; FeatureName = $featureName; Timestamp = (Get-Date -Format 'o')
        })))
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'EnableSharePointSiteFeature' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            SiteUrl = $siteUrl; FeatureId = $featureId; FeatureScope = $featureScope; Action = $action; FeatureName = $featureName; Timestamp = ''
        })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint site feature enable/disable script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
