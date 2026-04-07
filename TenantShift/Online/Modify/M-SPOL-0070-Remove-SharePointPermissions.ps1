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
    Removes SharePoint Online user or group permissions from site collections.

.DESCRIPTION
    Removes a specified user or group from a SharePoint Online site collection's
    permission assignments. For each row in the input CSV, removes the principal
    from any site groups they belong to, and removes any direct role assignments.
    Site collection admin rights are not removed by this script — use the SharePoint
    admin centre or M-SPOL-0030 to modify site collection admins.
    Supports -WhatIf for dry-run validation before committing changes.
    Reconnects to each site as the SiteUrl changes between rows.
    All results are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Required fields: SiteUrl, Principal.
    See the companion .input.csv template for the full column list.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).
    Used for the initial PnP tenant connection.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\M-SPOL-0070-Remove-SharePointPermissions.ps1 -InputCsvPath .\M-SPOL-0070-Remove-SharePointPermissions.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com -WhatIf

    Dry-run: shows what permission removals would be applied.

.EXAMPLE
    .\M-SPOL-0070-Remove-SharePointPermissions.ps1 -InputCsvPath .\M-SPOL-0070-Remove-SharePointPermissions.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com

    Remove the listed principals from the specified sites.

.NOTES
    Version:          1.0
    Required modules: PnP.PowerShell
    Required roles:   SharePoint Administrator
    Limitations:      Does not remove site collection admin rights.
                      Does not remove permissions granted via sharing links.
                      Reconnects to each site when SiteUrl changes — rows should be grouped by site for efficiency.

    CSV Fields:
    Column              Type      Required  Description
    ------------------  --------  --------  -----------
    SiteUrl             String    Yes       Full URL of the SharePoint site collection
    Principal           String    Yes       UPN or login name of the user/group to remove
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [ValidatePattern('^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$')]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M-SPOL-0070-Remove-SharePointPermissions_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$requiredHeaders = @('SiteUrl', 'Principal')

Write-Status -Message 'Starting SharePoint permission removal script.'
Assert-ModuleCurrent -ModuleNames @('PnP.PowerShell')

$adminUrlTrimmed  = $SharePointAdminUrl.TrimEnd('/')
$currentSiteUrl   = ''

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $siteUrl    = Get-TrimmedValue -Value $row.SiteUrl
    $principal  = Get-TrimmedValue -Value $row.Principal
    $primaryKey = "${siteUrl}|${principal}"

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl))   { throw 'SiteUrl is required.' }
        if ([string]::IsNullOrWhiteSpace($principal))  { throw 'Principal is required.' }

        # Reconnect when site changes.
        if ($siteUrl -ne $currentSiteUrl) {
            Write-Status -Message "Connecting to site: $siteUrl"
            Connect-PnPOnline -Url $siteUrl -Interactive -ErrorAction Stop
            $currentSiteUrl = $siteUrl
        }

        $description = "Remove principal '$principal' from all permission assignments on site '$siteUrl'"

        if (-not $PSCmdlet.ShouldProcess($siteUrl, $description)) {
            $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'RemoveSharePointPermission' -Status 'WhatIf' -Message "WhatIf: would remove '$principal' from '$siteUrl'." -Data ([ordered]@{
                SiteUrl = $siteUrl; Principal = $principal; GroupsRemoved = ''; RolesRemoved = ''; Timestamp = ''
            })))
            $rowNumber++
            continue
        }

        $groupsRemoved = [System.Collections.Generic.List[string]]::new()
        $rolesRemoved  = [System.Collections.Generic.List[string]]::new()

        # Remove from site groups.
        $groups = Invoke-WithRetry -OperationName "Get site groups for $siteUrl" -ScriptBlock {
            Get-PnPGroup -ErrorAction Stop
        }

        foreach ($group in $groups) {
            $members = Invoke-WithRetry -OperationName "Get members of group '$($group.Title)'" -ScriptBlock {
                Get-PnPGroupMember -Group $group.Title -ErrorAction Stop
            }

            $isMember = $members | Where-Object {
                ($_.Email -and $_.Email.Trim() -ieq $principal) -or
                ($_.LoginName -and $_.LoginName.Trim() -ieq $principal)
            }

            if ($isMember) {
                Invoke-WithRetry -OperationName "Remove '$principal' from group '$($group.Title)'" -ScriptBlock {
                    Remove-PnPGroupMember -Group $group.Title -LoginName $principal -ErrorAction Stop
                }
                $groupsRemoved.Add($group.Title)
            }
        }

        # Remove direct role assignments.
        try {
            Invoke-WithRetry -OperationName "Remove direct role assignment for '$principal' on $siteUrl" -ScriptBlock {
                Remove-PnPRoleAssignment -Principal $principal -ErrorAction Stop
            }
            $rolesRemoved.Add('DirectRoleAssignment')
        }
        catch {
            # No direct assignment to remove — not an error.
            if ($_.Exception.Message -inotmatch 'not found|does not exist|no assignment') {
                Write-Status -Message "  Note: direct role removal for '$principal' on '$siteUrl': $($_.Exception.Message)" -Level WARN
            }
        }

        $groupsStr = if ($groupsRemoved.Count -gt 0) { $groupsRemoved -join '; ' } else { '' }
        $rolesStr  = if ($rolesRemoved.Count -gt 0)  { $rolesRemoved  -join '; ' } else { '' }

        if ($groupsRemoved.Count -eq 0 -and $rolesRemoved.Count -eq 0) {
            $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'RemoveSharePointPermission' -Status 'Completed' -Message "No permissions found to remove for '$principal' on '$siteUrl'." -Data ([ordered]@{
                SiteUrl = $siteUrl; Principal = $principal; GroupsRemoved = ''; RolesRemoved = ''; Timestamp = (Get-Date -Format 'o')
            })))
        } else {
            $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'RemoveSharePointPermission' -Status 'Completed' -Message "Permissions removed for '$principal' on '$siteUrl'." -Data ([ordered]@{
                SiteUrl = $siteUrl; Principal = $principal; GroupsRemoved = $groupsStr; RolesRemoved = $rolesStr; Timestamp = (Get-Date -Format 'o')
            })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'RemoveSharePointPermission' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            SiteUrl = $siteUrl; Principal = $principal; GroupsRemoved = ''; RolesRemoved = ''; Timestamp = ''
        })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint permission removal script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
