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
    Provisions SharePointHubSites in Microsoft 365.

.DESCRIPTION
    Creates SharePointHubSites in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P3242-Create-SharePointHubSites.ps1 -InputCsvPath .\3242.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P3242-Create-SharePointHubSites.ps1 -InputCsvPath .\3242.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Online.SharePoint.PowerShell
    Required roles:   SharePoint Administrator
    Limitations:      None known.

    CSV Fields:
    Column                Type      Required  Description
    --------------------  ----      --------  -----------
    SiteUrl               String    Yes       <fill in description>
    HubDisplayName        String    Yes       <fill in description>
    HubDescription        String    Yes       <fill in description>
    HubLogoUrl            String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P3242-Create-SharePointHubSites_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {


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

function Get-SpoHubSiteIfExists {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    try {
        return Invoke-WithRetry -OperationName "Lookup hub site state for $SiteUrl" -ScriptBlock {
            Get-SPOHubSite -Identity $SiteUrl -ErrorAction Stop
        }
    }
    catch {
        $message = ([string]$_.Exception.Message).ToLowerInvariant()
        if ($message -match 'cannot find|was not found|does not exist|not found') {
            return $null
        }

        throw
    }
}

function Get-HubUpdateParameters {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$HubDisplayName,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$HubDescription,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$HubLogoUrl
    )

    $params = @{
        Identity    = $SiteUrl
        ErrorAction = 'Stop'
    }
    $notes = [System.Collections.Generic.List[string]]::new()

    $setHubCommand = Get-Command -Name Set-SPOHubSite -ErrorAction SilentlyContinue
    if (-not $setHubCommand) {
        return [PSCustomObject]@{
            Parameters = $params
            Notes      = @('Set-SPOHubSite command was not found; metadata update skipped.')
        }
    }

    $supportsTitle = $setHubCommand.Parameters.ContainsKey('Title')
    $supportsDescription = $setHubCommand.Parameters.ContainsKey('Description')
    $supportsLogoUrl = $setHubCommand.Parameters.ContainsKey('LogoUrl')

    if (-not [string]::IsNullOrWhiteSpace($HubDisplayName)) {
        if ($supportsTitle) {
            $params.Title = $HubDisplayName
        }
        else {
            $notes.Add('HubDisplayName provided but Set-SPOHubSite does not support -Title in this module version.')
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($HubDescription)) {
        if ($supportsDescription) {
            $params.Description = $HubDescription
        }
        else {
            $notes.Add('HubDescription provided but Set-SPOHubSite does not support -Description in this module version.')
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($HubLogoUrl)) {
        if ($supportsLogoUrl) {
            $params.LogoUrl = $HubLogoUrl
        }
        else {
            $notes.Add('HubLogoUrl provided but Set-SPOHubSite does not support -LogoUrl in this module version.')
        }
    }

    return [PSCustomObject]@{
        Parameters = $params
        Notes      = $notes.ToArray()
    }
}

$requiredHeaders = @(
    'SiteUrl',
    'HubDisplayName',
    'HubDescription',
    'HubLogoUrl'
)

Write-Status -Message 'Starting SharePoint hub site creation script.'
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
    $hubDisplayName = ([string]$row.HubDisplayName).Trim()
    $hubDescription = ([string]$row.HubDescription).Trim()
    $hubLogoUrl = ([string]$row.HubLogoUrl).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl)) {
            throw 'SiteUrl is required.'
        }

        if (-not [Uri]::IsWellFormedUriString($siteUrl, [UriKind]::Absolute)) {
            throw "SiteUrl '$siteUrl' is not a valid absolute URL."
        }

        Get-SpoSiteOrThrow -SiteUrl $siteUrl | Out-Null

        $existingHub = Get-SpoHubSiteIfExists -SiteUrl $siteUrl
        $wasRegistered = $false
        $metadataUpdated = $false
        $notes = [System.Collections.Generic.List[string]]::new()

        if (-not $existingHub) {
            if ($PSCmdlet.ShouldProcess($siteUrl, 'Register SharePoint hub site')) {
                Invoke-WithRetry -OperationName "Register hub site $siteUrl" -ScriptBlock {
                    Register-SPOHubSite -Site $siteUrl -ErrorAction Stop | Out-Null
                }
                $wasRegistered = $true
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'CreateSPOHubSite' -Status 'WhatIf' -Message 'Hub registration skipped due to WhatIf.'))
                $rowNumber++
                continue
            }
        }

        $hubUpdate = Get-HubUpdateParameters -SiteUrl $siteUrl -HubDisplayName $hubDisplayName -HubDescription $hubDescription -HubLogoUrl $hubLogoUrl
        foreach ($note in $hubUpdate.Notes) {
            $notes.Add($note)
        }

        if ($hubUpdate.Parameters.Count -gt 2) {
            if ($PSCmdlet.ShouldProcess($siteUrl, 'Update SharePoint hub metadata')) {
                $setHubParams = $hubUpdate.Parameters
                Invoke-WithRetry -OperationName "Update hub metadata $siteUrl" -ScriptBlock {
                    Set-SPOHubSite @setHubParams
                }
                $metadataUpdated = $true
            }
            else {
                $notes.Add('Hub metadata update skipped due to WhatIf.')
            }
        }

        if (-not $wasRegistered -and -not $metadataUpdated) {
            if ($notes.Count -eq 0) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'CreateSPOHubSite' -Status 'Skipped' -Message 'Site is already a hub and no metadata update was requested.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'CreateSPOHubSite' -Status 'Skipped' -Message ($notes -join ' ')))
            }
        }
        else {
            $messageParts = [System.Collections.Generic.List[string]]::new()
            if ($wasRegistered) {
                $messageParts.Add('Hub site registered.')
            }
            else {
                $messageParts.Add('Hub site already existed.')
            }

            if ($metadataUpdated) {
                $messageParts.Add('Hub metadata updated.')
            }

            if ($notes.Count -gt 0) {
                $messageParts.Add(($notes -join ' '))
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'CreateSPOHubSite' -Status 'Completed' -Message ($messageParts -join ' ')))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($siteUrl) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'CreateSPOHubSite' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint hub site creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







