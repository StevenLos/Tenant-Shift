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
Microsoft.Online.SharePoint.PowerShell

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Modifies OneDriveSharingSettings in Microsoft 365.

.DESCRIPTION
    Updates OneDriveSharingSettings in Microsoft 365 based on records provided in the input CSV file.
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
    .\SM-M3207-Set-OneDriveSharingSettings.ps1 -InputCsvPath .\3207.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M3207-Set-OneDriveSharingSettings.ps1 -InputCsvPath .\3207.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Online.SharePoint.PowerShell
    Required roles:   SharePoint Administrator
    Limitations:      None known.

    CSV Fields:
    Column                       Type      Required  Description
    ---------------------------  ----      --------  -----------
    UserPrincipalName            String    Yes       <fill in description>
    SharingCapability            String    Yes       <fill in description>
    DefaultSharingLinkType       String    Yes       <fill in description>
    DefaultLinkPermission        String    Yes       <fill in description>
    DefaultLinkToExistingAccess  String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3207-Set-OneDriveSharingSettings_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function ConvertTo-OneDriveUrlKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    return (($UserPrincipalName.Trim().ToLowerInvariant()) -replace '[^a-z0-9]', '_')
}

function Resolve-OneDriveSite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [hashtable]$SitesByOwner,

        [Parameter(Mandatory)]
        [hashtable]$SitesByUrlKey,

        [Parameter(Mandatory)]
        [string]$OneDriveHost
    )

    $ownerKey = $UserPrincipalName.Trim().ToLowerInvariant()
    $urlKey = ConvertTo-OneDriveUrlKey -UserPrincipalName $UserPrincipalName
    $expectedUrl = "https://$OneDriveHost/personal/$urlKey"

    $matches = [System.Collections.Generic.List[object]]::new()

    if ($SitesByOwner.ContainsKey($ownerKey)) {
        foreach ($site in $SitesByOwner[$ownerKey]) {
            $matches.Add($site)
        }
    }

    if ($SitesByUrlKey.ContainsKey($urlKey)) {
        foreach ($site in $SitesByUrlKey[$urlKey]) {
            $matches.Add($site)
        }
    }

    $uniqueSitesByUrl = @{}
    foreach ($site in $matches) {
        $siteUrlKey = ([string]$site.Url).Trim().ToLowerInvariant()
        if (-not [string]::IsNullOrWhiteSpace($siteUrlKey) -and -not $uniqueSitesByUrl.ContainsKey($siteUrlKey)) {
            $uniqueSitesByUrl[$siteUrlKey] = $site
        }
    }

    $resolvedSites = @($uniqueSitesByUrl.Values)
    if ($resolvedSites.Count -eq 0) {
        return [PSCustomObject]@{
            Status          = 'NotFound'
            Message         = 'No matching OneDrive personal site found for user.'
            Site            = $null
            ExpectedSiteUrl = $expectedUrl
        }
    }

    if ($resolvedSites.Count -eq 1) {
        return [PSCustomObject]@{
            Status          = 'Resolved'
            Message         = 'OneDrive personal site resolved.'
            Site            = $resolvedSites[0]
            ExpectedSiteUrl = $expectedUrl
        }
    }

    $expectedMatch = @($resolvedSites | Where-Object { ([string]$_.Url).Trim().Equals($expectedUrl, [System.StringComparison]::OrdinalIgnoreCase) })
    if ($expectedMatch.Count -eq 1) {
        return [PSCustomObject]@{
            Status          = 'Resolved'
            Message         = 'Multiple matches found; expected OneDrive URL selected.'
            Site            = $expectedMatch[0]
            ExpectedSiteUrl = $expectedUrl
        }
    }

    return [PSCustomObject]@{
        Status          = 'Ambiguous'
        Message         = 'Multiple OneDrive personal site matches found for user. Resolve ambiguity before applying updates.'
        Site            = $null
        ExpectedSiteUrl = $expectedUrl
    }
}

function ConvertTo-SetSpoParameterValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ParameterMetadata]$ParameterMetadata,

        [Parameter(Mandatory)]
        [string]$RawValue,

        [Parameter(Mandatory)]
        [string]$FieldName
    )

    $valueText = ([string]$RawValue).Trim()
    $parameterType = $ParameterMetadata.ParameterType

    if ($parameterType -eq [bool]) {
        return ConvertTo-Bool -Value $valueText
    }

    if ($parameterType.IsEnum) {
        try {
            return [System.Enum]::Parse($parameterType, $valueText, $true)
        }
        catch {
            $allowed = [System.Enum]::GetNames($parameterType) -join ', '
            throw "$FieldName '$valueText' is invalid. Valid values: $allowed."
        }
    }

    return $valueText
}

$requiredHeaders = @(
    'UserPrincipalName',
    'SharingCapability',
    'DefaultSharingLinkType',
    'DefaultLinkPermission',
    'DefaultLinkToExistingAccess'
)

Write-Status -Message 'Starting OneDrive sharing settings update script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Online.SharePoint.PowerShell')

if ([string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
    throw 'SharePointAdminUrl is required.'
}

$adminUrlTrimmed = $SharePointAdminUrl.Trim()
if ($adminUrlTrimmed -notmatch '^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$') {
    throw "SharePointAdminUrl '$adminUrlTrimmed' is invalid. Use format: https://<tenant>-admin.sharepoint.com"
}

Ensure-SharePointConnection -AdminUrl $adminUrlTrimmed

$setSpoSiteCommand = Get-Command -Name Set-SPOSite -ErrorAction Stop

$adminUri = [uri]$adminUrlTrimmed
$oneDriveHost = ($adminUri.Host -replace '-admin\.', '-my.')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

Write-Status -Message 'Loading personal sites to map OneDrive URLs by owner and URL key.'
$allSites = @(Invoke-WithRetry -OperationName 'Load personal sites' -ScriptBlock {
    Get-SPOSite -IncludePersonalSite $true -Limit All -Detailed -ErrorAction Stop
})

$personalSites = @($allSites | Where-Object {
    $url = ([string]$_.Url).Trim().ToLowerInvariant()
    $url.Contains('/personal/')
})

$sitesByOwner = @{}
$sitesByUrlKey = @{}

foreach ($site in $personalSites) {
    $siteUrl = ([string]$site.Url).Trim()
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        continue
    }

    $ownerKey = ([string]$site.Owner).Trim().ToLowerInvariant()
    if (-not [string]::IsNullOrWhiteSpace($ownerKey)) {
        if (-not $sitesByOwner.ContainsKey($ownerKey)) {
            $sitesByOwner[$ownerKey] = [System.Collections.Generic.List[object]]::new()
        }
        $sitesByOwner[$ownerKey].Add($site)
    }

    try {
        $uri = [uri]$siteUrl
        $parts = @($uri.AbsolutePath.Trim('/') -split '/')
        if ($parts.Count -ge 2 -and $parts[0].ToLowerInvariant() -eq 'personal') {
            $urlKey = $parts[1].ToLowerInvariant()
            if (-not [string]::IsNullOrWhiteSpace($urlKey)) {
                if (-not $sitesByUrlKey.ContainsKey($urlKey)) {
                    $sitesByUrlKey[$urlKey] = [System.Collections.Generic.List[object]]::new()
                }
                $sitesByUrlKey[$urlKey].Add($site)
            }
        }
    }
    catch {
        # Ignore malformed URLs in site results.
    }
}

$fieldMappings = @(
    @{ CsvColumn = 'SharingCapability'; ParameterName = 'SharingCapability' },
    @{ CsvColumn = 'DefaultSharingLinkType'; ParameterName = 'DefaultSharingLinkType' },
    @{ CsvColumn = 'DefaultLinkPermission'; ParameterName = 'DefaultLinkPermission' },
    @{ CsvColumn = 'DefaultLinkToExistingAccess'; ParameterName = 'DefaultLinkToExistingAccess' }
)

$rowNumber = 1
foreach ($row in $rows) {
    $userPrincipalName = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
            throw 'UserPrincipalName is required.'
        }

        if ($userPrincipalName -eq '*') {
            throw 'UserPrincipalName cannot be * for modify operations.'
        }

        $resolution = Resolve-OneDriveSite -UserPrincipalName $userPrincipalName -SitesByOwner $sitesByOwner -SitesByUrlKey $sitesByUrlKey -OneDriveHost $oneDriveHost
        if ($resolution.Status -eq 'NotFound') {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'SetOneDriveSharingSettings' -Status 'NotFound' -Message "$($resolution.Message) Expected URL: $($resolution.ExpectedSiteUrl)"))
            $rowNumber++
            continue
        }

        if ($resolution.Status -eq 'Ambiguous') {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'SetOneDriveSharingSettings' -Status 'Failed' -Message "$($resolution.Message) Expected URL: $($resolution.ExpectedSiteUrl)"))
            $rowNumber++
            continue
        }

        $site = $resolution.Site
        $siteUrl = ([string]$site.Url).Trim()

        $setParams = @{ Identity = $siteUrl }
        $changedFields = [System.Collections.Generic.List[string]]::new()

        foreach ($mapping in $fieldMappings) {
            $csvColumn = [string]$mapping.CsvColumn
            $parameterName = [string]$mapping.ParameterName
            $rawValue = ([string]$row.$csvColumn).Trim()

            if ([string]::IsNullOrWhiteSpace($rawValue)) {
                continue
            }

            if (-not $setSpoSiteCommand.Parameters.ContainsKey($parameterName)) {
                throw "Set-SPOSite parameter '-$parameterName' is not available in the installed Microsoft.Online.SharePoint.PowerShell module."
            }

            $convertedValue = ConvertTo-SetSpoParameterValue -ParameterMetadata $setSpoSiteCommand.Parameters[$parameterName] -RawValue $rawValue -FieldName $csvColumn
            $setParams[$parameterName] = $convertedValue
            $changedFields.Add($csvColumn)
        }

        if ($setParams.Count -eq 1) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'SetOneDriveSharingSettings' -Status 'Skipped' -Message 'No sharing settings were provided for update.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($siteUrl, 'Update OneDrive sharing settings')) {
            Invoke-WithRetry -OperationName "Update OneDrive sharing settings $siteUrl" -ScriptBlock {
                Set-SPOSite @setParams -ErrorAction Stop
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'SetOneDriveSharingSettings' -Status 'Completed' -Message ("Updated sharing fields: {0}." -f ($changedFields -join ', '))))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'SetOneDriveSharingSettings' -Status 'WhatIf' -Message 'Sharing settings update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($userPrincipalName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'SetOneDriveSharingSettings' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'OneDrive sharing settings update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
