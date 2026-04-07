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
Microsoft.Graph.Authentication
Microsoft.Graph.Reports

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Exports Entra ID application sign-in activity to CSV.

.DESCRIPTION
    Exports service principal sign-in activity from the Microsoft Graph beta endpoint
    (/reports/servicePrincipalSignInActivities). For each service principal, outputs
    the last sign-in date, success/failure counts, and application display name.
    Display names are resolved from the optional D-MEID-0130 output CSV when provided
    via -AppRegistrationsCsvPath, avoiding additional Graph calls per app.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all service principal sign-in activity in the tenant (-DiscoverAll).
    All results — including apps with no recorded sign-in activity — are written to the output CSV.

    Sign-in activity data retention:
    - Free / Microsoft 365 tenants: approximately 30 days
    - Entra ID P1 / P2 tenants: up to 90 days
    Apps with no activity within the retention window will show empty LastSignInActivity fields.

.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include AppId.
    See the companion .input.csv template for the full column list.

.PARAMETER DiscoverAll
    Enumerate all service principal sign-in activity in the tenant.

.PARAMETER AppRegistrationsCsvPath
    Optional path to the output CSV from D-MEID-0130-Get-EntraApplicationRegistrations.ps1.
    When provided, application display names are resolved from this file instead of
    making additional Graph calls per app.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-MEID-0140-Get-EntraApplicationSignInActivity.ps1 -DiscoverAll -AppRegistrationsCsvPath .\Results_D-MEID-0130*.csv

    Export all sign-in activity with display names resolved from D-MEID-0130 output.

.EXAMPLE
    .\D-MEID-0140-Get-EntraApplicationSignInActivity.ps1 -InputCsvPath .\D-MEID-0140-Get-EntraApplicationSignInActivity.input.csv

    Export sign-in activity for specific AppIds from the input CSV.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Reports
    Required roles:   Reports Reader, Security Reader, or Global Reader (Reports.Read.All or AuditLog.Read.All)
    Limitations:      Uses the Microsoft Graph beta endpoint — subject to change without deprecation notice.
                      Sign-in activity retention: ~30 days (free), up to 90 days (P1/P2).
                      Apps that have never signed in or have no activity in the retention window
                      will return empty LastSignInActivity fields but are still exported as Completed rows.

    CSV Fields:
    Column    Type      Required  Description
    --------  --------  --------  -----------
    AppId     String    Yes       Application (client) ID to query sign-in activity for
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(ParameterSetName = 'FromCsv')]
    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$AppRegistrationsCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-MEID-0140-Get-EntraApplicationSignInActivity_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function New-InventoryResult {
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

function Get-AppDisplayNameMap {
    [CmdletBinding()]
    param([AllowEmptyString()][string]$CsvPath)

    $map = @{}
    if ([string]::IsNullOrWhiteSpace($CsvPath) -or -not (Test-Path -LiteralPath $CsvPath)) {
        return $map
    }

    Write-Status -Message "Loading app display names from: $CsvPath"
    foreach ($row in (Import-Csv -LiteralPath $CsvPath | Where-Object { [string]$_.Status -eq 'Completed' })) {
        $appId = ([string]$row.AppId).Trim().ToLowerInvariant()
        $name  = Get-TrimmedValue -Value $row.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($appId) -and -not $map.ContainsKey($appId)) {
            $map[$appId] = $name
        }
    }
    Write-Status -Message "Loaded $($map.Count) display name entries."
    return $map
}

function Invoke-GraphBetaRequest {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri)

    $response = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
    return $response
}

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'AppId',
    'DisplayName',
    'LastSignInActivity',
    'LastSignInRequestId',
    'SuccessfulSignInCount',
    'FailedSignInCount'
)

$requiredHeaders = @('AppId')

Write-Status -Message 'Starting Entra application sign-in activity export script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Reports')
Ensure-GraphConnection -RequiredScopes @('Reports.Read.All')

# Load app display name map if provided.
$resolvedCsvPath = if ($PSBoundParameters.ContainsKey('AppRegistrationsCsvPath')) { $AppRegistrationsCsvPath.Trim() } else { '' }
$displayNameMap  = Get-AppDisplayNameMap -CsvPath $resolvedCsvPath

$scopeMode = 'Csv'

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. Fetching all service principal sign-in activities.' -Level WARN

    $allActivity = Invoke-WithRetry -OperationName 'Get all service principal sign-in activities' -ScriptBlock {
        $uri      = 'https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities?$top=999'
        $allItems = [System.Collections.Generic.List[object]]::new()
        do {
            $response = Invoke-GraphBetaRequest -Uri $uri
            foreach ($item in $response.value) { $allItems.Add($item) }
            $uri = if ($response.'@odata.nextLink') { $response.'@odata.nextLink' } else { $null }
        } while ($uri)
        return $allItems
    }

    Write-Status -Message "Fetched $($allActivity.Count) service principal sign-in activity records."
    $rows = @($allActivity | ForEach-Object { [PSCustomObject]@{ AppId = [string]$_.AppId } })
} else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $appId      = Get-TrimmedValue -Value $row.AppId
    $primaryKey = $appId

    if ([string]::IsNullOrWhiteSpace($appId)) {
        Write-Status -Message "Row $rowNumber skipped: AppId is empty." -Level WARN
        $rowNumber++
        continue
    }

    try {
        $activity = Invoke-WithRetry -OperationName "Get sign-in activity for $appId" -ScriptBlock {
            $uri = "https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities/$appId"
            try {
                return Invoke-GraphBetaRequest -Uri $uri
            } catch {
                if ($_.Exception.Message -imatch '404|NotFound') { return $null }
                throw
            }
        }

        $displayName = if ($displayNameMap.ContainsKey($appId.ToLowerInvariant())) { $displayNameMap[$appId.ToLowerInvariant()] } else { '' }

        if (-not $activity) {
            # App exists in input but has no sign-in activity record in the retention window.
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetEntraAppSignInActivity' -Status 'Completed' -Message 'No sign-in activity record found within the retention window.' -Data ([ordered]@{
                AppId                  = $appId
                DisplayName            = $displayName
                LastSignInActivity     = ''
                LastSignInRequestId    = ''
                SuccessfulSignInCount  = ''
                FailedSignInCount      = ''
            })))
        } else {
            $lastSignIn       = if ($activity.lastSignInActivity) { Get-TrimmedValue -Value $activity.lastSignInActivity.lastSignInDateTime } else { '' }
            $lastRequestId    = if ($activity.lastSignInActivity) { Get-TrimmedValue -Value $activity.lastSignInActivity.requestId } else { '' }
            $successCount     = if ($null -ne $activity.successfulSignInCount) { [string]$activity.successfulSignInCount } else { '' }
            $failCount        = if ($null -ne $activity.failedSignInCount) { [string]$activity.failedSignInCount } else { '' }

            if ([string]::IsNullOrWhiteSpace($displayName) -and $activity.PSObject.Properties['displayName']) {
                $displayName = Get-TrimmedValue -Value $activity.displayName
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetEntraAppSignInActivity' -Status 'Completed' -Message 'Sign-in activity exported.' -Data ([ordered]@{
                AppId                  = $appId
                DisplayName            = $displayName
                LastSignInActivity     = $lastSignIn
                LastSignInRequestId    = $lastRequestId
                SuccessfulSignInCount  = $successCount
                FailedSignInCount      = $failCount
            })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetEntraAppSignInActivity' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            AppId = $appId; DisplayName = ''; LastSignInActivity = ''; LastSignInRequestId = ''; SuccessfulSignInCount = ''; FailedSignInCount = ''
        })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra application sign-in activity export script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
