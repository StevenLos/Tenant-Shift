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
    Creates document libraries in SharePoint Online site collections from a CSV input.

.DESCRIPTION
    For each row in the input CSV, creates a document library in the specified site
    collection with the given name, description, and optional versioning settings.
    Skips creation if a list with the same name already exists and logs as Skipped.
    Reconnects to each site when SiteUrl changes between rows.
    Supports -WhatIf for dry-run validation before committing changes.
    All results are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Required fields: SiteUrl, LibraryName.
    See the companion .input.csv template for the full column list.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).
    Used for the initial PnP connection.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\P-SPOL-0030-Create-SharePointDocumentLibraries.ps1 -InputCsvPath .\P-SPOL-0030-Create-SharePointDocumentLibraries.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com -WhatIf

    Dry-run: shows what document libraries would be created.

.EXAMPLE
    .\P-SPOL-0030-Create-SharePointDocumentLibraries.ps1 -InputCsvPath .\P-SPOL-0030-Create-SharePointDocumentLibraries.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com

    Create the listed document libraries.

.NOTES
    Version:          1.0
    Required modules: PnP.PowerShell
    Required roles:   SharePoint Administrator or Site Owner
    Limitations:      Does not configure column schemas or content types — use PnP provisioning templates for that.
                      If a library with the same name already exists the row is Skipped (not overwritten).

    CSV Fields:
    Column              Type      Required  Description
    ------------------  --------  --------  -----------
    SiteUrl             String    Yes       Full URL of the SharePoint site collection
    LibraryName         String    Yes       Display name for the new document library
    Description         String    No        Description for the document library
    EnableVersioning    String    No        True to enable versioning on creation. Default: True.
    MajorVersionLimit   String    No        Maximum major versions to keep (0 = unlimited). Default: 500.
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [ValidatePattern('^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$')]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_P-SPOL-0030-Create-SharePointDocumentLibraries_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function New-ProvisionResult {
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

$requiredHeaders = @('SiteUrl', 'LibraryName')

Write-Status -Message 'Starting SharePoint document library creation script.'
Assert-ModuleCurrent -ModuleNames @('PnP.PowerShell')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders

$results        = [System.Collections.Generic.List[object]]::new()
$rowNumber      = 1
$currentSiteUrl = ''

foreach ($row in $rows) {
    $siteUrl            = Get-TrimmedValue -Value $row.SiteUrl
    $libraryName        = Get-TrimmedValue -Value $row.LibraryName
    $description        = if ($row.PSObject.Properties['Description']) { Get-TrimmedValue -Value $row.Description } else { '' }
    $enableVersioningRaw = if ($row.PSObject.Properties['EnableVersioning']) { Get-TrimmedValue -Value $row.EnableVersioning } else { 'True' }
    $majorVersionLimit  = if ($row.PSObject.Properties['MajorVersionLimit']) { Get-TrimmedValue -Value $row.MajorVersionLimit } else { '500' }
    $primaryKey         = "${siteUrl}|${libraryName}"

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl))     { throw 'SiteUrl is required.' }
        if ([string]::IsNullOrWhiteSpace($libraryName)) { throw 'LibraryName is required.' }

        $enableVersioning = $enableVersioningRaw.ToLowerInvariant() -ne 'false'

        $actionDescription = "Create document library '$libraryName' in '$siteUrl'"

        if (-not $PSCmdlet.ShouldProcess($siteUrl, $actionDescription)) {
            $results.Add((New-ProvisionResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateSharePointDocumentLibrary' -Status 'WhatIf' -Message "WhatIf: would create document library '$libraryName' in '$siteUrl'." -Data ([ordered]@{
                SiteUrl = $siteUrl; LibraryName = $libraryName; Description = $description
                EnableVersioning = $enableVersioningRaw; MajorVersionLimit = $majorVersionLimit; LibraryUrl = ''; Timestamp = ''
            })))
            $rowNumber++
            continue
        }

        if ($siteUrl -ne $currentSiteUrl) {
            Write-Status -Message "Connecting to site: $siteUrl"
            Connect-PnPOnline -Url $siteUrl -Interactive -ErrorAction Stop
            $currentSiteUrl = $siteUrl
        }

        # Check if library already exists.
        $existing = Invoke-WithRetry -OperationName "Check existing list '$libraryName' in $siteUrl" -ScriptBlock {
            try {
                Get-PnPList -Identity $libraryName -ErrorAction Stop
            } catch { $null }
        }

        if ($existing) {
            $results.Add((New-ProvisionResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateSharePointDocumentLibrary' -Status 'Skipped' -Message "Library '$libraryName' already exists in '$siteUrl' — skipped." -Data ([ordered]@{
                SiteUrl = $siteUrl; LibraryName = $libraryName; Description = $description
                EnableVersioning = $enableVersioningRaw; MajorVersionLimit = $majorVersionLimit
                LibraryUrl = [string]$existing.DefaultViewUrl; Timestamp = ''
            })))
            $rowNumber++
            continue
        }

        $newList = Invoke-WithRetry -OperationName "Create document library '$libraryName' in $siteUrl" -ScriptBlock {
            New-PnPList -Title $libraryName -Template DocumentLibrary -EnableVersioning:$enableVersioning -ErrorAction Stop
        }

        # Apply version limit if versioning is enabled and a limit is specified.
        if ($enableVersioning -and -not [string]::IsNullOrWhiteSpace($majorVersionLimit) -and $majorVersionLimit -ne '0') {
            Invoke-WithRetry -OperationName "Set version limit on '$libraryName'" -ScriptBlock {
                Set-PnPList -Identity $libraryName -MajorVersions ([int]$majorVersionLimit) -ErrorAction Stop | Out-Null
            }
        }

        # Set description if provided.
        if (-not [string]::IsNullOrWhiteSpace($description)) {
            Invoke-WithRetry -OperationName "Set description on '$libraryName'" -ScriptBlock {
                Set-PnPList -Identity $libraryName -Description $description -ErrorAction Stop | Out-Null
            }
        }

        $libraryUrl = if ($newList.DefaultViewUrl) { [string]$newList.DefaultViewUrl } else { '' }

        $results.Add((New-ProvisionResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateSharePointDocumentLibrary' -Status 'Completed' -Message "Document library '$libraryName' created in '$siteUrl'." -Data ([ordered]@{
            SiteUrl = $siteUrl; LibraryName = $libraryName; Description = $description
            EnableVersioning = $enableVersioningRaw; MajorVersionLimit = $majorVersionLimit
            LibraryUrl = $libraryUrl; Timestamp = (Get-Date -Format 'o')
        })))
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ProvisionResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateSharePointDocumentLibrary' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            SiteUrl = $siteUrl; LibraryName = $libraryName; Description = $description
            EnableVersioning = $enableVersioningRaw; MajorVersionLimit = $majorVersionLimit
            LibraryUrl = ''; Timestamp = ''
        })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint document library creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
