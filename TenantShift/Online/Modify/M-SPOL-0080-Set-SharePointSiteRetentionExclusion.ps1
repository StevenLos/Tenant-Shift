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
Microsoft.Graph.Authentication

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Excludes or re-includes SharePoint Online sites from tenant-level Microsoft Purview retention policies.

.DESCRIPTION
    Sets the adaptive scope exclusion flag for SharePoint Online site collections in
    Microsoft Purview compliance. For each site in the input CSV, either excludes
    the site from tenant-wide retention policies (ExcludeFromRetention = true) or
    re-includes it (ExcludeFromRetention = false) using the Microsoft Graph
    compliance API (/beta/compliance/ediscovery/... is not used; this script targets
    the /beta/security/dataRetention/exclusions endpoint for site-level exclusion).

    NOTE: As of April 2026 the Graph API surface for per-site retention exclusions
    is in beta and subject to change. The script uses the
    /beta/sites/{siteId}/settings endpoint to set the isRetentionExcluded flag.
    Verify the endpoint availability and required scopes with your Purview/compliance
    team before running in production.

    Supports -WhatIf for dry-run validation before committing changes.
    All results are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Required fields: SiteUrl, ExcludeFromRetention.
    See the companion .input.csv template for the full column list.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\M-SPOL-0080-Set-SharePointSiteRetentionExclusion.ps1 -InputCsvPath .\M-SPOL-0080-Set-SharePointSiteRetentionExclusion.input.csv -WhatIf

    Dry-run: shows what retention exclusion changes would be applied.

.EXAMPLE
    .\M-SPOL-0080-Set-SharePointSiteRetentionExclusion.ps1 -InputCsvPath .\M-SPOL-0080-Set-SharePointSiteRetentionExclusion.input.csv

    Apply retention exclusion settings for the listed sites.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication
    Required roles:   Compliance Administrator or Records Management (RecordsManagement.ReadWrite.All)
    Limitations:      Uses the Microsoft Graph beta endpoint — subject to change without deprecation notice.
                      The isRetentionExcluded site settings flag requires Purview compliance licensing.
                      Verify beta endpoint availability with your compliance team before production use.
                      Site IDs are resolved from site URLs via Graph (/sites?$filter=siteCollection/hostname...).

    CSV Fields:
    Column                  Type      Required  Description
    ----------------------  --------  --------  -----------
    SiteUrl                 String    Yes       Full URL of the SharePoint site collection
    ExcludeFromRetention    String    Yes       true to exclude from retention; false to re-include
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M-SPOL-0080-Set-SharePointSiteRetentionExclusion_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Resolve-SharePointSiteId {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SiteUrl)

    # Extract hostname and site path from URL.
    $uri      = [System.Uri]$SiteUrl
    $hostname = $uri.Host
    $sitePath = $uri.AbsolutePath.TrimStart('/')

    $filter   = "siteCollection/hostname eq '$hostname' and name eq '$($sitePath.Split('/')[-1])'"
    $graphUri = "https://graph.microsoft.com/v1.0/sites?`$filter=siteCollection/hostname+eq+'$hostname'&`$select=id,webUrl"

    $response = Invoke-MgGraphRequest -Method GET -Uri $graphUri -ErrorAction Stop
    $match    = $response.value | Where-Object { $_.webUrl -ieq $SiteUrl.TrimEnd('/') }

    if (-not $match) { throw "Could not resolve site ID for '$SiteUrl'. Verify the site URL and that the account has access." }
    return $match.id
}

$requiredHeaders = @('SiteUrl', 'ExcludeFromRetention')
$validValues     = @('true', 'false')

Write-Status -Message 'Starting SharePoint site retention exclusion script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication')
Ensure-GraphConnection -RequiredScopes @('RecordsManagement.ReadWrite.All', 'Sites.ReadWrite.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $siteUrl          = Get-TrimmedValue -Value $row.SiteUrl
    $excludeRaw       = Get-TrimmedValue -Value $row.ExcludeFromRetention
    $primaryKey       = $siteUrl

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl))    { throw 'SiteUrl is required.' }
        if ([string]::IsNullOrWhiteSpace($excludeRaw)) { throw 'ExcludeFromRetention is required.' }

        if ($excludeRaw.ToLowerInvariant() -notin $validValues) {
            throw "ExcludeFromRetention value '$excludeRaw' is invalid. Valid values: true, false."
        }

        $excludeBool = $excludeRaw.ToLowerInvariant() -eq 'true'
        $actionLabel = if ($excludeBool) { 'exclude from retention' } else { 're-include in retention' }
        $description = "Set isRetentionExcluded=$excludeBool on site '$siteUrl'"

        if (-not $PSCmdlet.ShouldProcess($siteUrl, $description)) {
            $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointSiteRetentionExclusion' -Status 'WhatIf' -Message "WhatIf: would $actionLabel for '$siteUrl'." -Data ([ordered]@{
                SiteUrl = $siteUrl; ExcludeFromRetention = $excludeRaw; SiteId = ''; Timestamp = ''
            })))
            $rowNumber++
            continue
        }

        # Resolve site ID from URL.
        $siteId = Invoke-WithRetry -OperationName "Resolve site ID for $siteUrl" -ScriptBlock {
            Resolve-SharePointSiteId -SiteUrl $siteUrl
        }

        # Apply the retention exclusion setting via Graph beta site settings.
        $settingsUri  = "https://graph.microsoft.com/beta/sites/$siteId/settings"
        $settingsBody = @{ isRetentionExcluded = $excludeBool }

        Invoke-WithRetry -OperationName "Set retention exclusion on $siteUrl" -ScriptBlock {
            Invoke-MgGraphRequest -Method PATCH -Uri $settingsUri -Body $settingsBody -ErrorAction Stop | Out-Null
        }

        $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointSiteRetentionExclusion' -Status 'Completed' -Message "Site retention exclusion set: isRetentionExcluded=$excludeBool on '$siteUrl'." -Data ([ordered]@{
            SiteUrl = $siteUrl; ExcludeFromRetention = $excludeRaw; SiteId = $siteId; Timestamp = (Get-Date -Format 'o')
        })))
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointSiteRetentionExclusion' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            SiteUrl = $siteUrl; ExcludeFromRetention = $excludeRaw; SiteId = ''; Timestamp = ''
        })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint site retention exclusion script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
