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
    Sets the list experience mode (Auto, Modern, or Classic) on SharePoint lists and libraries.

.DESCRIPTION
    For each row in the input CSV, sets the ListExperienceOptions property on the specified
    list or document library to Auto (tenant default), NewExperience (Modern), or
    ClassicExperience. If ListName is omitted, the setting is applied to all non-hidden
    lists in the site.
    Useful for standardising list experience modes before or after migration, or for
    reverting sites to Classic for compatibility with legacy workflows.
    Reconnects to each site when SiteUrl changes between rows.
    Supports -WhatIf for dry-run validation before committing changes.
    All results are written to the output CSV.
    Use D-SPOL-0050-Get-SharePointListExperience.ps1 first to discover current modes.

.PARAMETER InputCsvPath
    Path to the input CSV file. Required fields: SiteUrl, Experience.
    See the companion .input.csv template for the full column list.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).
    Used for the initial PnP connection.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\M-SPOL-0110-Set-SharePointListExperience.ps1 -InputCsvPath .\M-SPOL-0110-Set-SharePointListExperience.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com -WhatIf

    Dry-run: shows what list experience changes would be applied.

.EXAMPLE
    .\M-SPOL-0110-Set-SharePointListExperience.ps1 -InputCsvPath .\M-SPOL-0110-Set-SharePointListExperience.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com

    Apply list experience settings.

.NOTES
    Version:          1.0
    Required modules: PnP.PowerShell
    Required roles:   SharePoint Administrator or Site Owner
    Limitations:      Experience mode changes take effect immediately but may require
                      a browser refresh to be visible to end users.

    CSV Fields:
    Column      Type      Required  Description
    ----------  --------  --------  -----------
    SiteUrl     String    Yes       Full URL of the SharePoint site collection
    ListName    String    No        Title of the list/library. If blank, applies to all non-hidden lists.
    Experience  String    Yes       Auto, NewExperience (Modern), or ClassicExperience
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [ValidatePattern('^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$')]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M-SPOL-0110-Set-SharePointListExperience_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

# Map human-readable values to PnP ListExperienceOptions enum values.
$experienceMap = @{
    'auto'              = 'Auto'
    'newexperience'     = 'NewExperience'
    'modern'            = 'NewExperience'
    'classicexperience' = 'ClassicExperience'
    'classic'           = 'ClassicExperience'
}

$requiredHeaders = @('SiteUrl', 'Experience')

Write-Status -Message 'Starting SharePoint list experience mode script.'
Assert-ModuleCurrent -ModuleNames @('PnP.PowerShell')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders

$results        = [System.Collections.Generic.List[object]]::new()
$rowNumber      = 1
$currentSiteUrl = ''

foreach ($row in $rows) {
    $siteUrl    = Get-TrimmedValue -Value $row.SiteUrl
    $listName   = if ($row.PSObject.Properties['ListName']) { Get-TrimmedValue -Value $row.ListName } else { '' }
    $experience = Get-TrimmedValue -Value $row.Experience
    $primaryKey = if ($listName) { "${siteUrl}|${listName}" } else { "${siteUrl}|*AllLists*" }

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl))    { throw 'SiteUrl is required.' }
        if ([string]::IsNullOrWhiteSpace($experience)) { throw 'Experience is required.' }

        $normalizedExperience = $experienceMap[$experience.ToLowerInvariant()]
        if (-not $normalizedExperience) {
            throw "Experience '$experience' is invalid. Valid values: Auto, NewExperience (or Modern), ClassicExperience (or Classic)."
        }

        $targetLabel = if ($listName) { "list '$listName'" } else { 'all lists' }
        $description = "Set list experience to '$normalizedExperience' on $targetLabel in '$siteUrl'"

        if (-not $PSCmdlet.ShouldProcess($siteUrl, $description)) {
            $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointListExperience' -Status 'WhatIf' -Message "WhatIf: would $description." -Data ([ordered]@{
                SiteUrl = $siteUrl; ListName = $listName; Experience = $normalizedExperience; Timestamp = ''
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
                Get-PnPList -Identity $listName -ErrorAction Stop
            })
        } else {
            Invoke-WithRetry -OperationName "Get all lists from $siteUrl" -ScriptBlock {
                Get-PnPList -Includes Hidden -ErrorAction Stop | Where-Object { -not $_.Hidden }
            }
        }

        foreach ($list in $listsToProcess) {
            try {
                Invoke-WithRetry -OperationName "Set experience on list '$($list.Title)' in $siteUrl" -ScriptBlock {
                    Set-PnPList -Identity $list.Title -ListExperience $normalizedExperience -ErrorAction Stop | Out-Null
                }

                $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointListExperience' -Status 'Completed' -Message "List experience set to '$normalizedExperience' on '$($list.Title)' in '$siteUrl'." -Data ([ordered]@{
                    SiteUrl = $siteUrl; ListName = $list.Title; Experience = $normalizedExperience; Timestamp = (Get-Date -Format 'o')
                })))
            }
            catch {
                Write-Status -Message "  List '$($list.Title)' in '$siteUrl' failed: $($_.Exception.Message)" -Level ERROR
                $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey "${siteUrl}|$($list.Title)" -Action 'SetSharePointListExperience' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    SiteUrl = $siteUrl; ListName = $list.Title; Experience = $normalizedExperience; Timestamp = ''
                })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointListExperience' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            SiteUrl = $siteUrl; ListName = $listName; Experience = $experience; Timestamp = ''
        })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint list experience mode script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
