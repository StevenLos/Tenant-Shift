#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [bool]$NoWait = $true,

    [string]$OutputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath ("Results_B04-PreProvision-OneDrive_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'M365.Common.psm1') -Force -DisableNameChecking

function Test-SharePointConnection {
    [CmdletBinding()]
    param()

    try {
        Get-SPOTenant -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Ensure-SharePointConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AdminUrl
    )

    if (Test-SharePointConnection) {
        Write-Status -Message 'Already connected to SharePoint Online.' -Level SUCCESS
        return
    }

    Write-Status -Message 'No active SharePoint Online connection detected. Connecting now.' -Level WARN
    Connect-SPOService -Url $AdminUrl -ErrorAction Stop

    if (-not (Test-SharePointConnection)) {
        throw 'SharePoint Online connection failed. Unable to verify an active session.'
    }

    Write-Status -Message 'Connected to SharePoint Online.' -Level SUCCESS
}

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

