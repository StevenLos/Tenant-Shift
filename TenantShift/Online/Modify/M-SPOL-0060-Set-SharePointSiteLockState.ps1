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
    Sets the lock state of SharePoint Online site collections.

.DESCRIPTION
    Sets the lock state of site collections to NoAccess, ReadOnly, or Unlock.
    Used at migration cutover to prevent writes to decommissioned sites, or to
    restore site access after migration is confirmed complete.
    Operationally significant — NoAccess immediately prevents all user access.
    All state changes are logged with timestamps in the output CSV.
    Supports -WhatIf for dry-run validation. NoAccess operations require -Confirm.

.PARAMETER InputCsvPath
    Path to the input CSV file. Required fields: SiteUrl, LockState.
    See the companion .input.csv template for the full column list.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\M-SPOL-0040-Set-SharePointSiteLockState.ps1 -InputCsvPath .\M-SPOL-0040-Set-SharePointSiteLockState.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com -WhatIf

    Dry-run: shows what lock state changes would be applied without making changes.

.EXAMPLE
    .\M-SPOL-0040-Set-SharePointSiteLockState.ps1 -InputCsvPath .\M-SPOL-0040-Set-SharePointSiteLockState.input.csv -SharePointAdminUrl https://los-admin.sharepoint.com

    Apply lock state changes. NoAccess rows prompt for -Confirm before applying.

.NOTES
    Version:          1.0
    Required modules: PnP.PowerShell
    Required roles:   SharePoint Administrator
    Limitations:      NoAccess immediately blocks all site access including site admins.
                      Always validate the target sites before applying NoAccess at cutover.
                      This script can run independently of D-SPOL-0020 and M-SPOL-0030.

    CSV Fields:
    Column      Type      Required  Description
    ----------  --------  --------  -----------
    SiteUrl     String    Yes       Absolute URL of the site collection
    LockState   String    Yes       Target lock state: NoAccess, ReadOnly, or Unlock
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M-SPOL-0060-Set-SharePointSiteLockState_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$requiredHeaders  = @('SiteUrl', 'LockState')
$validLockStates  = @('NoAccess', 'ReadOnly', 'Unlock')

Write-Status -Message 'Starting SharePoint site lock state script.'
Assert-ModuleCurrent -ModuleNames @('PnP.PowerShell')

$adminUrlTrimmed = $SharePointAdminUrl.Trim().TrimEnd('/')
if ($adminUrlTrimmed -notmatch '^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com$') {
    throw "SharePointAdminUrl '$adminUrlTrimmed' is invalid. Use: https://<tenant>-admin.sharepoint.com"
}

Write-Status -Message "Connecting to SharePoint admin center: $adminUrlTrimmed"
Connect-PnPOnline -Url $adminUrlTrimmed -Interactive -ErrorAction Stop

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $siteUrl   = ([string]$row.SiteUrl).Trim().TrimEnd('/')
    $lockState = Get-TrimmedValue -Value $row.LockState
    $primaryKey = $siteUrl

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl))    { throw 'SiteUrl is required.' }
        if ([string]::IsNullOrWhiteSpace($lockState))  { throw 'LockState is required.' }

        # Normalize lock state casing.
        $normalizedLockState = $validLockStates | Where-Object { $_ -ieq $lockState }
        if (-not $normalizedLockState) {
            throw "LockState '$lockState' is invalid. Valid values: NoAccess, ReadOnly, Unlock."
        }
        $lockState = $normalizedLockState

        # Get current lock state before applying change.
        $currentState = Invoke-WithRetry -OperationName "Get current lock state for $siteUrl" -ScriptBlock {
            $site = Get-PnPTenantSite -Url $siteUrl -ErrorAction Stop
            return [string]$site.LockState
        }

        if ($currentState -ieq $lockState) {
            Write-Status -Message "Row $rowNumber ($siteUrl): already in LockState '$lockState'. Skipping."
            $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointSiteLockState' -Status 'Skipped' -Message "Site is already in LockState '$lockState'." -Data ([ordered]@{
                SiteUrl          = $siteUrl
                RequestedLockState = $lockState
                PreviousLockState  = $currentState
                AppliedLockState   = ''
                Timestamp          = ''
            })))
            $rowNumber++
            continue
        }

        $description = "Set LockState to '$lockState' on $siteUrl (currently: $currentState)"

        # NoAccess is operationally significant — require explicit confirmation.
        if ($lockState -eq 'NoAccess') {
            if (-not $PSCmdlet.ShouldProcess($siteUrl, $description)) {
                $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointSiteLockState' -Status 'WhatIf' -Message "WhatIf: would set LockState to '$lockState'." -Data ([ordered]@{
                    SiteUrl = $siteUrl; RequestedLockState = $lockState; PreviousLockState = $currentState; AppliedLockState = ''; Timestamp = ''
                })))
                $rowNumber++
                continue
            }

            if (-not $PSCmdlet.ShouldContinue("Setting LockState to NoAccess on $siteUrl will IMMEDIATELY block all user access, including site administrators. Current state: $currentState. Continue?", 'Confirm NoAccess')) {
                $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointSiteLockState' -Status 'Skipped' -Message 'NoAccess operation cancelled by user.' -Data ([ordered]@{
                    SiteUrl = $siteUrl; RequestedLockState = $lockState; PreviousLockState = $currentState; AppliedLockState = ''; Timestamp = ''
                })))
                $rowNumber++
                continue
            }
        } else {
            if (-not $PSCmdlet.ShouldProcess($siteUrl, $description)) {
                $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointSiteLockState' -Status 'WhatIf' -Message "WhatIf: would set LockState to '$lockState'." -Data ([ordered]@{
                    SiteUrl = $siteUrl; RequestedLockState = $lockState; PreviousLockState = $currentState; AppliedLockState = ''; Timestamp = ''
                })))
                $rowNumber++
                continue
            }
        }

        $appliedTimestamp = Get-Date -Format 'o'
        Invoke-WithRetry -OperationName "Set lock state on $siteUrl" -ScriptBlock {
            Set-PnPTenantSite -Url $siteUrl -LockState $lockState -ErrorAction Stop
        }

        Write-Status -Message "Row $rowNumber ($siteUrl): LockState changed '$currentState' -> '$lockState'." -Level $(if ($lockState -eq 'NoAccess') { 'WARN' } else { 'INFO' })

        $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointSiteLockState' -Status 'Completed' -Message "LockState set to '$lockState'." -Data ([ordered]@{
            SiteUrl            = $siteUrl
            RequestedLockState = $lockState
            PreviousLockState  = $currentState
            AppliedLockState   = $lockState
            Timestamp          = $appliedTimestamp
        })))
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ModifyResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetSharePointSiteLockState' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            SiteUrl            = $siteUrl
            RequestedLockState = $lockState
            PreviousLockState  = ''
            AppliedLockState   = ''
            Timestamp          = ''
        })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint site lock state script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
