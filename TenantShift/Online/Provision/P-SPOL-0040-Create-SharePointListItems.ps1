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
    Bulk-creates list items in a SharePoint Online list from a CSV input.

.DESCRIPTION
    For each row in the input CSV, creates a new list item in the specified SharePoint
    list. The CSV must include SiteUrl and ListName. All remaining columns are treated
    as field name-value pairs and mapped to list fields using their internal names.
    Field introspection is performed on first connection to each site/list combination
    to validate that referenced columns exist. Unknown columns are logged as warnings
    but do not halt processing.
    Reconnects to each site when SiteUrl changes between rows.
    Supports -WhatIf for dry-run validation before committing changes.
    All results are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Required fields: SiteUrl, ListName. All other columns
    are mapped to list fields.
    See the companion .input.csv template for an example schema.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).
    Used for the initial PnP connection.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\P-SPOL-0040-Create-SharePointListItems.ps1 -InputCsvPath .\P-SPOL-0040-Create-SharePointListItems.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com -WhatIf

    Dry-run: shows what list items would be created.

.EXAMPLE
    .\P-SPOL-0040-Create-SharePointListItems.ps1 -InputCsvPath .\P-SPOL-0040-Create-SharePointListItems.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com

    Create the listed items.

.NOTES
    Version:          1.0
    Required modules: PnP.PowerShell
    Required roles:   SharePoint Administrator or Site Owner with list Contribute permission
    Limitations:      Supports text, number, date, and boolean field types.
                      Person/Group and Lookup fields are not supported by this script — use
                      the PnP provisioning engine for complex field type mappings.
                      All values are written as strings and SharePoint performs implicit type coercion.
                      Does not update existing items — only creates new ones.

    CSV Fields:
    Column      Type      Required  Description
    ----------  --------  --------  -----------
    SiteUrl     String    Yes       Full URL of the SharePoint site collection
    ListName    String    Yes       Internal or display name of the target list
    Title       String    No        Value for the Title field (most lists require this)
    *           String    No        Any additional columns map to list field internal names
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [ValidatePattern('^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$')]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_P-SPOL-0040-Create-SharePointListItems_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

# Reserved columns — not passed to the list as field values.
$reservedColumns = @('SiteUrl', 'ListName')

$requiredHeaders = @('SiteUrl', 'ListName')

Write-Status -Message 'Starting SharePoint list item creation script.'
Assert-ModuleCurrent -ModuleNames @('PnP.PowerShell')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders

# Determine the non-reserved field columns from the CSV header.
$csvHeaders   = $rows[0].PSObject.Properties.Name
$fieldColumns = $csvHeaders | Where-Object { $_ -notin $reservedColumns }

Write-Status -Message "Field columns detected in CSV: $($fieldColumns -join ', ')"

$results            = [System.Collections.Generic.List[object]]::new()
$rowNumber          = 1
$currentSiteUrl     = ''
$currentListKey     = ''
$knownListFields    = @{}

foreach ($row in $rows) {
    $siteUrl    = Get-TrimmedValue -Value $row.SiteUrl
    $listName   = Get-TrimmedValue -Value $row.ListName
    $primaryKey = "${siteUrl}|${listName}|Row${rowNumber}"

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl))  { throw 'SiteUrl is required.' }
        if ([string]::IsNullOrWhiteSpace($listName)) { throw 'ListName is required.' }

        $listKey = "${siteUrl}|${listName}"

        if (-not $PSCmdlet.ShouldProcess("$listName in $siteUrl", "Create list item (Row $rowNumber)")) {
            $fieldSummary = ($fieldColumns | ForEach-Object { "$_=$($row.$_)" }) -join '; '
            $results.Add((New-ProvisionResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateSharePointListItem' -Status 'WhatIf' -Message "WhatIf: would create item in '$listName' at '$siteUrl' with fields: $fieldSummary." -Data ([ordered]@{
                SiteUrl = $siteUrl; ListName = $listName; ItemId = ''; FieldsSummary = $fieldSummary; Timestamp = ''
            })))
            $rowNumber++
            continue
        }

        if ($siteUrl -ne $currentSiteUrl) {
            Write-Status -Message "Connecting to site: $siteUrl"
            Connect-PnPOnline -Url $siteUrl -Interactive -ErrorAction Stop
            $currentSiteUrl = $siteUrl
        }

        # Introspect list fields on first access to this list.
        if ($listKey -ne $currentListKey) {
            Write-Status -Message "Introspecting fields for list '$listName' in '$siteUrl'."
            $listFields = Invoke-WithRetry -OperationName "Get fields for '$listName' in $siteUrl" -ScriptBlock {
                Get-PnPField -List $listName -ErrorAction Stop | Select-Object InternalName, Title
            }
            $knownListFields[$listKey] = @{}
            foreach ($field in $listFields) {
                $knownListFields[$listKey][$field.InternalName] = $field.Title
            }
            $currentListKey = $listKey
        }

        # Build field value hashtable — skip empty values, warn on unknown fields.
        $fieldValues = @{}
        foreach ($col in $fieldColumns) {
            $val = Get-TrimmedValue -Value $row.$col
            if ([string]::IsNullOrWhiteSpace($val)) { continue }

            if (-not $knownListFields[$listKey].ContainsKey($col)) {
                Write-Status -Message "  Column '$col' not found as a field in list '$listName' — skipped." -Level WARN
                continue
            }
            $fieldValues[$col] = $val
        }

        $newItem = Invoke-WithRetry -OperationName "Create item in '$listName' at $siteUrl" -ScriptBlock {
            Add-PnPListItem -List $listName -Values $fieldValues -ErrorAction Stop
        }

        $fieldSummary = ($fieldValues.Keys | ForEach-Object { "$_=$($fieldValues[$_])" }) -join '; '

        $results.Add((New-ProvisionResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateSharePointListItem' -Status 'Completed' -Message "List item created in '$listName' at '$siteUrl' (ID: $($newItem.Id))." -Data ([ordered]@{
            SiteUrl = $siteUrl; ListName = $listName; ItemId = [string]$newItem.Id; FieldsSummary = $fieldSummary; Timestamp = (Get-Date -Format 'o')
        })))
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ProvisionResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateSharePointListItem' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            SiteUrl = $siteUrl; ListName = $listName; ItemId = ''; FieldsSummary = ''; Timestamp = ''
        })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint list item creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
