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
    Creates migration artifact tracking entries in a designated SharePoint tracking list.

.DESCRIPTION
    Writes migration artifact records to a SharePoint list used as a living migration
    inventory. For each row in the input CSV, creates an entry in the designated
    tracking list recording the artifact source path, destination path, migration type,
    status, and optional notes.
    The tracking list must already exist. Use P-SPOL-0040 to bulk-create it if needed,
    or create it manually in the SharePoint site.
    Duplicate detection: if an item with the same SourcePath already exists in the list,
    the row is Skipped unless -OverwriteExisting is specified, in which case the existing
    item is updated.
    Reconnects to each site when SiteUrl changes between rows.
    Supports -WhatIf for dry-run validation before committing changes.
    All results are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Required fields: SiteUrl, ArtifactTrackerListName, SourcePath, MigrationType.
    See the companion .input.csv template for the full column list.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).
    Used for the initial PnP connection.

.PARAMETER OverwriteExisting
    When specified, updates existing tracker items whose SourcePath matches rather than skipping them.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\P-SPOL-0050-Create-SharePointArtifactTrackerItems.ps1 -InputCsvPath .\P-SPOL-0050-Create-SharePointArtifactTrackerItems.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com -WhatIf

    Dry-run: shows what tracker items would be created.

.EXAMPLE
    .\P-SPOL-0050-Create-SharePointArtifactTrackerItems.ps1 -InputCsvPath .\P-SPOL-0050-Create-SharePointArtifactTrackerItems.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com

    Create artifact tracker items.

.EXAMPLE
    .\P-SPOL-0050-Create-SharePointArtifactTrackerItems.ps1 -InputCsvPath .\P-SPOL-0050-Create-SharePointArtifactTrackerItems.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com -OverwriteExisting

    Create or update artifact tracker items — updates existing entries by SourcePath.

.NOTES
    Version:          1.0
    Required modules: PnP.PowerShell
    Required roles:   SharePoint Administrator or Site Owner with list Contribute permission
    Limitations:      The tracking list must already exist with the expected field schema.
                      Expected field internal names: Title, SourcePath, DestinationPath,
                      MigrationType, MigrationStatus, Notes.
                      Duplicate detection is by SourcePath field exact match (CAML query).

    CSV Fields:
    Column                      Type      Required  Description
    --------------------------  --------  --------  -----------
    SiteUrl                     String    Yes       Full URL of the SharePoint site collection hosting the tracker
    ArtifactTrackerListName     String    Yes       Internal or display name of the tracking list
    SourcePath                  String    Yes       Source path or URL of the artifact being tracked
    DestinationPath             String    No        Target path or URL after migration
    MigrationType               String    Yes       Type of migration (e.g. Mailbox, Site, FileShare, Team)
    MigrationStatus             String    No        Initial status (e.g. Pending, InProgress, Completed). Default: Pending.
    Notes                       String    No        Free-text notes for this artifact
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [ValidatePattern('^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$')]
    [string]$SharePointAdminUrl,

    [switch]$OverwriteExisting,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_P-SPOL-0050-Create-SharePointArtifactTrackerItems_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$requiredHeaders = @('SiteUrl', 'ArtifactTrackerListName', 'SourcePath', 'MigrationType')

Write-Status -Message 'Starting SharePoint artifact tracker item creation script.'
Assert-ModuleCurrent -ModuleNames @('PnP.PowerShell')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders

$results        = [System.Collections.Generic.List[object]]::new()
$rowNumber      = 1
$currentSiteUrl = ''

foreach ($row in $rows) {
    $siteUrl         = Get-TrimmedValue -Value $row.SiteUrl
    $listName        = Get-TrimmedValue -Value $row.ArtifactTrackerListName
    $sourcePath      = Get-TrimmedValue -Value $row.SourcePath
    $destPath        = if ($row.PSObject.Properties['DestinationPath']) { Get-TrimmedValue -Value $row.DestinationPath } else { '' }
    $migrationType   = Get-TrimmedValue -Value $row.MigrationType
    $migrationStatus = if ($row.PSObject.Properties['MigrationStatus'] -and -not [string]::IsNullOrWhiteSpace($row.MigrationStatus)) { Get-TrimmedValue -Value $row.MigrationStatus } else { 'Pending' }
    $notes           = if ($row.PSObject.Properties['Notes']) { Get-TrimmedValue -Value $row.Notes } else { '' }
    $primaryKey      = "${siteUrl}|${listName}|${sourcePath}"

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl))       { throw 'SiteUrl is required.' }
        if ([string]::IsNullOrWhiteSpace($listName))      { throw 'ArtifactTrackerListName is required.' }
        if ([string]::IsNullOrWhiteSpace($sourcePath))    { throw 'SourcePath is required.' }
        if ([string]::IsNullOrWhiteSpace($migrationType)) { throw 'MigrationType is required.' }

        $description = "Create tracker item for '$sourcePath' in list '$listName' at '$siteUrl'"

        if (-not $PSCmdlet.ShouldProcess($siteUrl, $description)) {
            $results.Add((New-ProvisionResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateArtifactTrackerItem' -Status 'WhatIf' -Message "WhatIf: would create tracker item for '$sourcePath' in '$listName' at '$siteUrl'." -Data ([ordered]@{
                SiteUrl = $siteUrl; ArtifactTrackerListName = $listName; SourcePath = $sourcePath
                DestinationPath = $destPath; MigrationType = $migrationType; MigrationStatus = $migrationStatus
                Notes = $notes; ItemId = ''; ItemAction = ''; Timestamp = ''
            })))
            $rowNumber++
            continue
        }

        if ($siteUrl -ne $currentSiteUrl) {
            Write-Status -Message "Connecting to site: $siteUrl"
            Connect-PnPOnline -Url $siteUrl -Interactive -ErrorAction Stop
            $currentSiteUrl = $siteUrl
        }

        # Check for existing item by SourcePath (CAML query).
        $caml = "<View><Query><Where><Eq><FieldRef Name='SourcePath'/><Value Type='Text'>$([System.Security.SecurityElement]::Escape($sourcePath))</Value></Eq></Where></Query></View>"
        $existing = Invoke-WithRetry -OperationName "Check existing tracker item for '$sourcePath'" -ScriptBlock {
            Get-PnPListItem -List $listName -Query $caml -ErrorAction Stop
        }

        $fieldValues = @{
            Title           = $sourcePath
            SourcePath      = $sourcePath
            DestinationPath = $destPath
            MigrationType   = $migrationType
            MigrationStatus = $migrationStatus
            Notes           = $notes
        }

        if ($existing -and @($existing).Count -gt 0) {
            if (-not $OverwriteExisting) {
                $results.Add((New-ProvisionResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateArtifactTrackerItem' -Status 'Skipped' -Message "Tracker item for '$sourcePath' already exists in '$listName' — skipped. Use -OverwriteExisting to update." -Data ([ordered]@{
                    SiteUrl = $siteUrl; ArtifactTrackerListName = $listName; SourcePath = $sourcePath
                    DestinationPath = $destPath; MigrationType = $migrationType; MigrationStatus = $migrationStatus
                    Notes = $notes; ItemId = [string]$existing[0].Id; ItemAction = 'Skipped'; Timestamp = ''
                })))
                $rowNumber++
                continue
            }

            # Update existing item.
            $existingId = $existing[0].Id
            Invoke-WithRetry -OperationName "Update tracker item $existingId for '$sourcePath'" -ScriptBlock {
                Set-PnPListItem -List $listName -Identity $existingId -Values $fieldValues -ErrorAction Stop | Out-Null
            }

            $results.Add((New-ProvisionResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateArtifactTrackerItem' -Status 'Completed' -Message "Tracker item updated for '$sourcePath' in '$listName' at '$siteUrl' (ID: $existingId)." -Data ([ordered]@{
                SiteUrl = $siteUrl; ArtifactTrackerListName = $listName; SourcePath = $sourcePath
                DestinationPath = $destPath; MigrationType = $migrationType; MigrationStatus = $migrationStatus
                Notes = $notes; ItemId = [string]$existingId; ItemAction = 'Updated'; Timestamp = (Get-Date -Format 'o')
            })))
        } else {
            # Create new item.
            $newItem = Invoke-WithRetry -OperationName "Create tracker item for '$sourcePath' in '$listName'" -ScriptBlock {
                Add-PnPListItem -List $listName -Values $fieldValues -ErrorAction Stop
            }

            $results.Add((New-ProvisionResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateArtifactTrackerItem' -Status 'Completed' -Message "Tracker item created for '$sourcePath' in '$listName' at '$siteUrl' (ID: $($newItem.Id))." -Data ([ordered]@{
                SiteUrl = $siteUrl; ArtifactTrackerListName = $listName; SourcePath = $sourcePath
                DestinationPath = $destPath; MigrationType = $migrationType; MigrationStatus = $migrationStatus
                Notes = $notes; ItemId = [string]$newItem.Id; ItemAction = 'Created'; Timestamp = (Get-Date -Format 'o')
            })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ProvisionResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateArtifactTrackerItem' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            SiteUrl = $siteUrl; ArtifactTrackerListName = $listName; SourcePath = $sourcePath
            DestinationPath = $destPath; MigrationType = $migrationType; MigrationStatus = $migrationStatus
            Notes = $notes; ItemId = ''; ItemAction = ''; Timestamp = ''
        })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint artifact tracker item creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
