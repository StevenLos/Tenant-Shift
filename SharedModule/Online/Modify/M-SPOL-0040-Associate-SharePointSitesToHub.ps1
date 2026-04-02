<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-004416

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Online.SharePoint.PowerShell

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Modifies SharePointSitesToHub in Microsoft 365.

.DESCRIPTION
    Updates SharePointSitesToHub in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M3243-Associate-SharePointSitesToHub.ps1 -InputCsvPath .\3243.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M3243-Associate-SharePointSitesToHub.ps1 -InputCsvPath .\3243.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Online.SharePoint.PowerShell
    Required roles:   SharePoint Administrator
    Limitations:      None known.

    CSV Fields:
    Column                Type      Required  Description
    --------------------  ----      --------  -----------
    HubSiteUrl            String    Yes       <fill in description>
    HubSiteId             String    Yes       <fill in description>
    SiteUrl               String    Yes       <fill in description>
    EnforceSameTenant     String    Yes       <fill in description>
    AllowReassociation    String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3243-Associate-SharePointSitesToHub_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {


function Normalize-Id {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    return $text.Trim('{}').ToLowerInvariant()
}

function Get-HubIdFromObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$HubObject
    )

    foreach ($propName in @('ID', 'Id', 'HubSiteId', 'SiteId')) {
        if ($HubObject.PSObject.Properties.Name -contains $propName) {
            $value = Normalize-Id -Value $HubObject.$propName
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
        }
    }

    return ''
}

function Get-HubSiteUrlFromObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$HubObject
    )

    foreach ($propName in @('SiteUrl', 'SiteUrls', 'Url')) {
        if ($HubObject.PSObject.Properties.Name -contains $propName) {
            $value = ([string]$HubObject.$propName).Trim()
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
        }
    }

    return ''
}

function Get-CurrentHubIdFromSite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$SiteObject
    )

    foreach ($propName in @('HubSiteId', 'HubSiteID', 'HubSiteIdValue')) {
        if ($SiteObject.PSObject.Properties.Name -contains $propName) {
            $value = Normalize-Id -Value $SiteObject.$propName
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
        }
    }

    return ''
}

function Get-SpoSiteOrThrow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    try {
        return Invoke-WithRetry -OperationName "Lookup SharePoint site $SiteUrl" -ScriptBlock {
            Get-SPOSite -Identity $SiteUrl -Detailed -ErrorAction Stop
        }
    }
    catch {
        $message = ([string]$_.Exception.Message).ToLowerInvariant()
        if ($message -match 'cannot find|was not found|does not exist|not found') {
            throw "Site '$SiteUrl' was not found."
        }

        throw
    }
}

function Resolve-HubSite {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$HubSiteUrl,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$HubSiteId
    )

    $hubUrlTrimmed = ([string]$HubSiteUrl).Trim()
    $hubIdNormalized = Normalize-Id -Value $HubSiteId

    if (-not [string]::IsNullOrWhiteSpace($hubUrlTrimmed)) {
        try {
            $hub = Invoke-WithRetry -OperationName "Lookup hub site by URL $hubUrlTrimmed" -ScriptBlock {
                Get-SPOHubSite -Identity $hubUrlTrimmed -ErrorAction Stop
            }

            return $hub
        }
        catch {
            $message = ([string]$_.Exception.Message).ToLowerInvariant()
            if ($message -match 'cannot find|was not found|does not exist|not found') {
                throw "Hub site '$hubUrlTrimmed' was not found."
            }

            throw
        }
    }

    if ([string]::IsNullOrWhiteSpace($hubIdNormalized)) {
        throw 'Either HubSiteUrl or HubSiteId must be provided.'
    }

    $allHubs = @(Invoke-WithRetry -OperationName "Load all hub sites for ID lookup $hubIdNormalized" -ScriptBlock {
        Get-SPOHubSite -ErrorAction Stop
    })

    $matching = @()
    foreach ($hub in $allHubs) {
        $hubId = Get-HubIdFromObject -HubObject $hub
        if ([string]::IsNullOrWhiteSpace($hubId)) {
            continue
        }

        if ($hubId.Equals($hubIdNormalized, [System.StringComparison]::OrdinalIgnoreCase)) {
            $matching += $hub
        }
    }

    if ($matching.Count -eq 0) {
        throw "Hub site ID '$hubIdNormalized' was not found."
    }

    if ($matching.Count -gt 1) {
        throw "Multiple hub sites matched ID '$hubIdNormalized'."
    }

    return $matching[0]
}

$requiredHeaders = @(
    'HubSiteUrl',
    'HubSiteId',
    'SiteUrl',
    'EnforceSameTenant'
)

Write-Status -Message 'Starting SharePoint site-to-hub association script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Online.SharePoint.PowerShell')

if ([string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
    throw 'SharePointAdminUrl is required.'
}

$adminUrlTrimmed = $SharePointAdminUrl.Trim()
if ($adminUrlTrimmed -notmatch '^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$') {
    throw "SharePointAdminUrl '$adminUrlTrimmed' is invalid. Use: https://<tenant>-admin.sharepoint.com"
}

Ensure-SharePointConnection -AdminUrl $adminUrlTrimmed

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $siteUrl = ([string]$row.SiteUrl).Trim()
    $hubSiteUrl = ([string]$row.HubSiteUrl).Trim()
    $hubSiteId = ([string]$row.HubSiteId).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl)) {
            throw 'SiteUrl is required.'
        }

        $allowReassociation = $false
        if ($row.PSObject.Properties.Name -contains 'AllowReassociation') {
            $allowReassociation = ConvertTo-Bool -Value $row.AllowReassociation -Default $false
        }

        $enforceSameTenant = ConvertTo-Bool -Value $row.EnforceSameTenant -Default $true

        $site = Get-SpoSiteOrThrow -SiteUrl $siteUrl
        $hub = Resolve-HubSite -HubSiteUrl $hubSiteUrl -HubSiteId $hubSiteId

        $resolvedHubId = Get-HubIdFromObject -HubObject $hub
        $resolvedHubUrl = Get-HubSiteUrlFromObject -HubObject $hub
        if ([string]::IsNullOrWhiteSpace($resolvedHubUrl)) {
            $resolvedHubUrl = $hubSiteUrl
        }

        if ([string]::IsNullOrWhiteSpace($resolvedHubId) -or [string]::IsNullOrWhiteSpace($resolvedHubUrl)) {
            throw 'Unable to resolve hub ID and URL from provided HubSiteUrl/HubSiteId.'
        }

        if ($enforceSameTenant) {
            $siteHost = ([Uri]$siteUrl).Host
            $hubHost = ([Uri]$resolvedHubUrl).Host
            if (-not $siteHost.Equals($hubHost, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Site '$siteUrl' and hub '$resolvedHubUrl' are not in the same tenant host."
            }
        }

        $currentHubId = Get-CurrentHubIdFromSite -SiteObject $site
        if (-not [string]::IsNullOrWhiteSpace($currentHubId) -and $currentHubId.Equals($resolvedHubId, [System.StringComparison]::OrdinalIgnoreCase)) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$siteUrl|$resolvedHubUrl" -Action 'AssociateSPOSiteToHub' -Status 'Skipped' -Message 'Site is already associated to the requested hub.'))
            $rowNumber++
            continue
        }

        if (-not [string]::IsNullOrWhiteSpace($currentHubId) -and -not $currentHubId.Equals($resolvedHubId, [System.StringComparison]::OrdinalIgnoreCase) -and -not $allowReassociation) {
            throw "Site is already associated to a different hub ($currentHubId). Set AllowReassociation to TRUE to move it."
        }

        if ($PSCmdlet.ShouldProcess($siteUrl, "Associate to hub $resolvedHubUrl")) {
            if (-not [string]::IsNullOrWhiteSpace($currentHubId) -and -not $currentHubId.Equals($resolvedHubId, [System.StringComparison]::OrdinalIgnoreCase) -and $allowReassociation) {
                Invoke-WithRetry -OperationName "Remove existing hub association for $siteUrl" -ScriptBlock {
                    Remove-SPOHubSiteAssociation -Site $siteUrl -ErrorAction Stop
                }
            }

            Invoke-WithRetry -OperationName "Associate $siteUrl to hub $resolvedHubUrl" -ScriptBlock {
                Add-SPOHubSiteAssociation -Site $siteUrl -HubSite $resolvedHubUrl -ErrorAction Stop
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$siteUrl|$resolvedHubUrl" -Action 'AssociateSPOSiteToHub' -Status 'Completed' -Message 'Site associated to hub successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$siteUrl|$resolvedHubUrl" -Action 'AssociateSPOSiteToHub' -Status 'WhatIf' -Message 'Hub association skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($siteUrl) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'AssociateSPOSiteToHub' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint site-to-hub association script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







