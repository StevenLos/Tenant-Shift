<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-000101

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Online.SharePoint.PowerShell

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Modifies OneDrive in Microsoft 365.

.DESCRIPTION
    Updates OneDrive in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).

.PARAMETER NoWait
    When $true (default), submits the request asynchronously without waiting for completion.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M3204-PreProvision-OneDrive.ps1 -InputCsvPath .\3204.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M3204-PreProvision-OneDrive.ps1 -InputCsvPath .\3204.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Online.SharePoint.PowerShell
    Required roles:   SharePoint Administrator
    Limitations:      None known.

    CSV Fields:
    Column                Type      Required  Description
    --------------------  ----      --------  -----------
    UserPrincipalName     String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [bool]$NoWait = $true,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3204-PreProvision-OneDrive_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {
$requiredHeaders = @(
    'UserPrincipalName'
)

Write-Status -Message 'Starting OneDrive pre-provisioning script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Online.SharePoint.PowerShell')

if ([string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
    throw 'SharePointAdminUrl is required.'
}

$adminUrlTrimmed = $SharePointAdminUrl.Trim()
if ($adminUrlTrimmed -notmatch '^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$') {
    throw "SharePointAdminUrl '$adminUrlTrimmed' is invalid. Use the tenant admin URL format: https://<tenant>-admin.sharepoint.com"
}

Ensure-SharePointConnection -AdminUrl $adminUrlTrimmed

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()
$processedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

$rowNumber = 1
foreach ($row in $rows) {
    $upn = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw 'UserPrincipalName is required.'
        }

        if (-not $processedUsers.Add($upn)) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'PreProvisionOneDrive' -Status 'Skipped' -Message 'Duplicate user in CSV (already processed earlier in this run).'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($upn, 'Request OneDrive pre-provisioning')) {
            try {
                Invoke-WithRetry -OperationName "Request OneDrive pre-provisioning for $upn" -ScriptBlock {
                    Request-SPOPersonalSite -UserEmails $upn -NoWait:$NoWait -ErrorAction Stop | Out-Null
                }

                $modeText = if ($NoWait) { 'Request queued with -NoWait.' } else { 'Request submitted.' }
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'PreProvisionOneDrive' -Status 'Requested' -Message "OneDrive pre-provisioning request succeeded. $modeText"))
            }
            catch {
                $innerMessage = $_.Exception.Message
                $lowerInnerMessage = $innerMessage.ToLowerInvariant()

                if ($lowerInnerMessage -match 'already.*personal site|already.*onedrive|personal site already exists|site .* already exists') {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'PreProvisionOneDrive' -Status 'Skipped' -Message 'OneDrive appears to already exist (or was previously requested).'))
                }
                else {
                    throw
                }
            }
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'PreProvisionOneDrive' -Status 'WhatIf' -Message 'Request skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'PreProvisionOneDrive' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'OneDrive pre-provisioning script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





