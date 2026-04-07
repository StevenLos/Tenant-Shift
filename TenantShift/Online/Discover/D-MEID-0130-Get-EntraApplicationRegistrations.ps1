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
Microsoft.Graph.Applications

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Exports Entra ID app registrations and enterprise applications to CSV.

.DESCRIPTION
    Exports all app registrations and enterprise applications (service principals) from
    Entra ID via Microsoft Graph. For each app, outputs display name, application ID,
    object ID, creation date, publisher domain, sign-in audience, sign-in enabled flag,
    and assigned owners.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all applications in the tenant (-DiscoverAll parameter set).
    All results — including apps that could not be queried — are written to the output CSV.
    Output of this script can be provided to D-MEID-0140 via -AppRegistrationsCsvPath for
    display name resolution in the sign-in activity report.

.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include AppId.
    See the companion .input.csv template for the full column list.

.PARAMETER DiscoverAll
    Enumerate all app registrations in the tenant rather than processing from an input CSV file.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-MEID-0130-Get-EntraApplicationRegistrations.ps1 -InputCsvPath .\D-MEID-0130-Get-EntraApplicationRegistrations.input.csv

    Export app registration data for the AppIds listed in the input CSV.

.EXAMPLE
    .\D-MEID-0130-Get-EntraApplicationRegistrations.ps1 -DiscoverAll

    Export all app registrations in the tenant.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Applications
    Required roles:   Application Administrator, Cloud Application Administrator, or Global Reader
    Limitations:      Owner resolution requires Microsoft.Graph.Users access (read-only).
                      Pagination is handled automatically via the Graph SDK.

    CSV Fields:
    Column    Type      Required  Description
    --------  --------  --------  -----------
    AppId     String    Yes       Application (client) ID of the app registration to export
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-MEID-0130-Get-EntraApplicationRegistrations_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function New-EmptyAppData {
    [CmdletBinding()]
    param([string]$AppIdRequested = '')

    return [ordered]@{
        AppIdRequested    = $AppIdRequested
        DisplayName       = ''
        AppId             = ''
        ObjectId          = ''
        PublisherDomain   = ''
        SignInAudience    = ''
        CreatedDateTime   = ''
        Description       = ''
        Notes             = ''
        Owners            = ''
    }
}

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'AppIdRequested',
    'DisplayName',
    'AppId',
    'ObjectId',
    'PublisherDomain',
    'SignInAudience',
    'CreatedDateTime',
    'Description',
    'Notes',
    'Owners'
)

$requiredHeaders = @('AppId')

Write-Status -Message 'Starting Entra application registrations export script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications')
Ensure-GraphConnection -RequiredScopes @('Application.Read.All', 'User.Read.All')

$scopeMode = 'Csv'

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. Fetching all app registrations.' -Level WARN

    $allApps = Invoke-WithRetry -OperationName 'Get all app registrations' -ScriptBlock {
        Get-MgApplication -All -Property AppId, DisplayName, Id, PublisherDomain, SignInAudience, CreatedDateTime, Description, Notes -ErrorAction Stop
    }

    Write-Status -Message "Fetched $($allApps.Count) app registrations."
    $rows = @($allApps | ForEach-Object { [PSCustomObject]@{ AppId = [string]$_.AppId } })
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
        $app = Invoke-WithRetry -OperationName "Get app registration $appId" -ScriptBlock {
            Get-MgApplication -Filter "appId eq '$appId'" -Property AppId, DisplayName, Id, PublisherDomain, SignInAudience, CreatedDateTime, Description, Notes -ErrorAction Stop
        }

        if (-not $app -or @($app).Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetEntraApplicationRegistration' -Status 'NotFound' -Message 'No app registration found with the specified AppId.' -Data (New-EmptyAppData -AppIdRequested $appId)))
            $rowNumber++
            continue
        }

        $appObj = if ($app -is [array]) { $app[0] } else { $app }

        # Resolve owners.
        $owners = Invoke-WithRetry -OperationName "Get owners for app $appId" -ScriptBlock {
            Get-MgApplicationOwner -ApplicationId $appObj.Id -ErrorAction SilentlyContinue
        }

        $ownerList = if ($owners) {
            ($owners | ForEach-Object {
                $upn = if ($_.AdditionalProperties.ContainsKey('userPrincipalName')) { [string]$_.AdditionalProperties['userPrincipalName'] } else { [string]$_.Id }
                $upn.Trim()
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join '; '
        } else { '' }

        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $appId -Action 'GetEntraApplicationRegistration' -Status 'Completed' -Message 'App registration exported.' -Data ([ordered]@{
            AppIdRequested  = $appId
            DisplayName     = Get-TrimmedValue -Value $appObj.DisplayName
            AppId           = Get-TrimmedValue -Value $appObj.AppId
            ObjectId        = Get-TrimmedValue -Value $appObj.Id
            PublisherDomain = Get-TrimmedValue -Value $appObj.PublisherDomain
            SignInAudience  = Get-TrimmedValue -Value $appObj.SignInAudience
            CreatedDateTime = Get-TrimmedValue -Value $appObj.CreatedDateTime
            Description     = Get-TrimmedValue -Value $appObj.Description
            Notes           = Get-TrimmedValue -Value $appObj.Notes
            Owners          = $ownerList
        })))
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetEntraApplicationRegistration' -Status 'Failed' -Message $_.Exception.Message -Data (New-EmptyAppData -AppIdRequested $appId)))
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
Write-Status -Message 'Entra application registrations export script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
