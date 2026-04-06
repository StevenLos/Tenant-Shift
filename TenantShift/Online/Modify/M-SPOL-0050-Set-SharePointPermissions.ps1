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
    Grants, revokes, or resets permissions on SharePoint Online site collections.

.DESCRIPTION
    Modifies site-level permissions on SharePoint Online site collections based on records
    provided in the input CSV file. Supports three actions per row:
    - Grant: Adds the specified principal to a site group at the given permission level.
    - Revoke: Removes the specified principal from a site group or direct permission.
    - Reset: Clears all unique permissions on the site and restores inheritance.
             The Reset action is destructive and requires -Confirm.
    Supports -WhatIf for dry-run validation before committing changes.
    All results — including rows that could not be processed — are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Required fields: SiteUrl, Principal, PermissionLevel, Action.
    See the companion .input.csv template for the full column list.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\M-SPOL-0030-Set-SharePointPermissions.ps1 -InputCsvPath .\M-SPOL-0030-Set-SharePointPermissions.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com -WhatIf

    Dry-run: shows what permission changes would be made without applying them.

.EXAMPLE
    .\M-SPOL-0030-Set-SharePointPermissions.ps1 -InputCsvPath .\M-SPOL-0030-Set-SharePointPermissions.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com

    Apply permission changes from the input CSV.

.NOTES
    Version:          1.0
    Required modules: PnP.PowerShell
    Required roles:   SharePoint Administrator or Site Collection Administrator for each target site
    Limitations:      Reset action is destructive — it removes all unique permissions on the site
                      root web and restores permission inheritance. Requires -Confirm.
                      Grant/Revoke operates at the site-level group/direct-assignment layer only.

    CSV Fields:
    Column           Type      Required  Description
    ---------------  --------  --------  -----------
    SiteUrl          String    Yes       Absolute URL of the site collection
    Principal        String    Yes       Login name or email of the user or group
    PermissionLevel  String    Yes       SharePoint permission level (e.g. Read, Contribute, Full Control) — ignored for Revoke and Reset actions
    Action           String    Yes       Grant, Revoke, or Reset
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M-SPOL-0050-Set-SharePointPermissions_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Connect-PnPSite {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Url)

    Connect-PnPOnline -Url $Url -Interactive -ErrorAction Stop
}

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

$requiredHeaders = @('SiteUrl', 'Principal', 'PermissionLevel', 'Action')
$validActions    = @('Grant', 'Revoke', 'Reset')

Write-Status -Message 'Starting SharePoint permissions modify script.'
Assert-ModuleCurrent -ModuleNames @('PnP.PowerShell')

$adminUrlTrimmed = $SharePointAdminUrl.Trim().TrimEnd('/')
if ($adminUrlTrimmed -notmatch '^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com$') {
    throw "SharePointAdminUrl '$adminUrlTrimmed' is invalid. Use: https://<tenant>-admin.sharepoint.com"
}

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders

$results          = [System.Collections.Generic.List[object]]::new()
$rowNumber        = 1
$currentSiteUrl   = ''

foreach ($row in $rows) {
    $siteUrl        = ([string]$row.SiteUrl).Trim().TrimEnd('/')
    $principal      = Get-TrimmedValue -Value $row.Principal
    $permissionLevel = Get-TrimmedValue -Value $row.PermissionLevel
    $action         = Get-TrimmedValue -Value $row.Action
    $primaryKey     = "${siteUrl}|${principal}|${action}"

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl))    { throw 'SiteUrl is required.' }
        if ([string]::IsNullOrWhiteSpace($action))     { throw 'Action is required.' }
        if ($action -notin $validActions)               { throw "Action '$action' is invalid. Valid values: Grant, Revoke, Reset." }
        if ($action -ne 'Reset' -and [string]::IsNullOrWhiteSpace($principal)) {
            throw 'Principal is required for Grant and Revoke actions.'
        }
        if ($action -eq 'Grant' -and [string]::IsNullOrWhiteSpace($permissionLevel)) {
            throw 'PermissionLevel is required for Grant action.'
        }

        # Connect to the site if it has changed.
        if ($siteUrl -ne $currentSiteUrl) {
            Write-Status -Message "Connecting to site: $siteUrl"
            Connect-PnPSite -Url $siteUrl
            $currentSiteUrl = $siteUrl
        }

        switch ($action) {
            'Grant' {
                $description = "Grant '$permissionLevel' to '$principal' on $siteUrl"
                if ($PSCmdlet.ShouldProcess($siteUrl, $description)) {
                    Add-PnPRoleAssignment -LoginName $principal -RoleDefinitionName $permissionLevel -ErrorAction Stop
                    $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointPermission' -Status 'Completed' -Message "Granted '$permissionLevel' to '$principal'." -Data ([ordered]@{
                        SiteUrl         = $siteUrl
                        Principal       = $principal
                        PermissionLevel = $permissionLevel
                        ChangeAction    = $action
                        Timestamp       = (Get-Date -Format 'o')
                    })))
                } else {
                    $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointPermission' -Status 'WhatIf' -Message "WhatIf: would grant '$permissionLevel' to '$principal'." -Data ([ordered]@{
                        SiteUrl = $siteUrl; Principal = $principal; PermissionLevel = $permissionLevel; ChangeAction = $action; Timestamp = ''
                    })))
                }
            }
            'Revoke' {
                $description = "Revoke permissions for '$principal' on $siteUrl"
                if ($PSCmdlet.ShouldProcess($siteUrl, $description)) {
                    Remove-PnPRoleAssignment -LoginName $principal -ErrorAction Stop
                    $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointPermission' -Status 'Completed' -Message "Revoked permissions for '$principal'." -Data ([ordered]@{
                        SiteUrl         = $siteUrl
                        Principal       = $principal
                        PermissionLevel = $permissionLevel
                        ChangeAction    = $action
                        Timestamp       = (Get-Date -Format 'o')
                    })))
                } else {
                    $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointPermission' -Status 'WhatIf' -Message "WhatIf: would revoke permissions for '$principal'." -Data ([ordered]@{
                        SiteUrl = $siteUrl; Principal = $principal; PermissionLevel = $permissionLevel; ChangeAction = $action; Timestamp = ''
                    })))
                }
            }
            'Reset' {
                # Reset is destructive: clears all unique permissions and restores inheritance.
                $description = "DESTRUCTIVE: Reset all unique permissions on $siteUrl and restore inheritance"
                if (-not $PSCmdlet.ShouldProcess($siteUrl, $description)) {
                    $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointPermission' -Status 'WhatIf' -Message "WhatIf: would reset all unique permissions on $siteUrl." -Data ([ordered]@{
                        SiteUrl = $siteUrl; Principal = ''; PermissionLevel = ''; ChangeAction = $action; Timestamp = ''
                    })))
                } else {
                    if (-not $PSCmdlet.ShouldContinue("This will clear ALL unique permissions on $siteUrl and restore inheritance. This action cannot be undone. Continue?", 'Confirm Reset')) {
                        $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointPermission' -Status 'Skipped' -Message 'Reset cancelled by user.' -Data ([ordered]@{
                            SiteUrl = $siteUrl; Principal = ''; PermissionLevel = ''; ChangeAction = $action; Timestamp = ''
                        })))
                    } else {
                        Reset-PnPRoleInheritance -ErrorAction Stop
                        $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointPermission' -Status 'Completed' -Message 'Unique permissions cleared and inheritance restored.' -Data ([ordered]@{
                            SiteUrl         = $siteUrl
                            Principal       = ''
                            PermissionLevel = ''
                            ChangeAction    = $action
                            Timestamp       = (Get-Date -Format 'o')
                        })))
                    }
                }
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointPermission' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            SiteUrl         = $siteUrl
            Principal       = $principal
            PermissionLevel = $permissionLevel
            ChangeAction    = $action
            Timestamp       = ''
        })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint permissions modify script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
