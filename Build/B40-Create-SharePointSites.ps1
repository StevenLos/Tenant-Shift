#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [switch]$FailIfExists,

    [string]$OutputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath ("Results_B40-Create-SharePointSites_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'M365.Common.psm1') -Force -DisableNameChecking

function ConvertTo-NullableInt {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value,

        [Parameter(Mandatory)]
        [string]$FieldName
    )

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $parsed = 0
    if (-not [int]::TryParse($text, [ref]$parsed)) {
        throw "$FieldName '$text' is not a valid integer."
    }

    return $parsed
}

function Get-SpoSiteIfExists {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    try {
        return Invoke-WithRetry -OperationName "Lookup SharePoint site $SiteUrl" -ScriptBlock {
            Get-SPOSite -Identity $SiteUrl -Detailed -ErrorAction Stop
        }
    }
    catch {
        $message = ([string]$_.Exception.Message).ToLowerInvariant()
        if ($message -match 'cannot find|was not found|does not exist|not found') {
            return $null
        }

        throw
    }
}

$requiredHeaders = @(
    'SiteUrl',
    'Title',
    'Template',
    'PrimaryOwnerUPN',
    'SecondaryOwnersUPNs',
    'Language',
    'TimeZoneId',
    'StorageQuotaMB'
)

Write-Status -Message 'Starting SharePoint site creation script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Online.SharePoint.PowerShell')

if ([string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
    throw 'SharePointAdminUrl is required.'
}

$adminUrlTrimmed = $SharePointAdminUrl.Trim()
if ($adminUrlTrimmed -notmatch '^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$') {
    throw "SharePointAdminUrl '$adminUrlTrimmed' is invalid. Use: https://<tenant>-admin.sharepoint.com"
}

$adminUri = [Uri]$adminUrlTrimmed
$tenantHost = $adminUri.Host -replace '-admin\.sharepoint\.com$', '.sharepoint.com'

Ensure-SharePointConnection -AdminUrl $adminUrlTrimmed

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $siteUrl = ([string]$row.SiteUrl).Trim()
    $title = ([string]$row.Title).Trim()
    $template = ([string]$row.Template).Trim()
    $primaryOwnerUpn = ([string]$row.PrimaryOwnerUPN).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($siteUrl) -or [string]::IsNullOrWhiteSpace($title) -or [string]::IsNullOrWhiteSpace($template) -or [string]::IsNullOrWhiteSpace($primaryOwnerUpn)) {
            throw 'SiteUrl, Title, Template, and PrimaryOwnerUPN are required.'
        }

        if (-not [Uri]::IsWellFormedUriString($siteUrl, [UriKind]::Absolute)) {
            throw "SiteUrl '$siteUrl' is not a valid absolute URL."
        }

        $siteUri = [Uri]$siteUrl
        if ($siteUri.Scheme -ne 'https') {
            throw "SiteUrl '$siteUrl' must use HTTPS."
        }

        if (-not $siteUri.Host.EndsWith('.sharepoint.com', [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "SiteUrl '$siteUrl' must be a SharePoint Online URL."
        }

        if (-not $siteUri.Host.Equals($tenantHost, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "SiteUrl '$siteUrl' is outside tenant host '$tenantHost'."
        }

        $secondaryOwners = ConvertTo-Array -Value ([string]$row.SecondaryOwnersUPNs)
        $language = ConvertTo-NullableInt -Value $row.Language -FieldName 'Language'
        $timeZoneId = ConvertTo-NullableInt -Value $row.TimeZoneId -FieldName 'TimeZoneId'
        $storageQuotaMb = ConvertTo-NullableInt -Value $row.StorageQuotaMB -FieldName 'StorageQuotaMB'

        if ($storageQuotaMb -ne $null -and $storageQuotaMb -le 0) {
            throw 'StorageQuotaMB must be greater than zero when provided.'
        }

        $existingSite = Get-SpoSiteIfExists -SiteUrl $siteUrl
        if ($existingSite) {
            if ($FailIfExists) {
                throw "Site '$siteUrl' already exists and FailIfExists was specified."
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'CreateSPOSite' -Status 'Skipped' -Message 'Site already exists.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($siteUrl, 'Create SharePoint site')) {
            $createParams = @{
                Url         = $siteUrl
                Owner       = $primaryOwnerUpn
                Title       = $title
                Template    = $template
                ErrorAction = 'Stop'
            }

            if ($language -ne $null) {
                $createParams.LocaleId = $language
            }

            if ($timeZoneId -ne $null) {
                $createParams.TimeZoneId = $timeZoneId
            }

            if ($storageQuotaMb -ne $null) {
                $createParams.StorageQuota = $storageQuotaMb
            }

            Invoke-WithRetry -OperationName "Create SharePoint site $siteUrl" -ScriptBlock {
                New-SPOSite @createParams | Out-Null
            }

            $uniqueSecondaryOwners = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($secondaryOwner in $secondaryOwners) {
                $ownerUpn = ([string]$secondaryOwner).Trim()
                if ([string]::IsNullOrWhiteSpace($ownerUpn)) {
                    continue
                }

                if ($ownerUpn.Equals($primaryOwnerUpn, [System.StringComparison]::OrdinalIgnoreCase)) {
                    continue
                }

                if (-not $uniqueSecondaryOwners.Add($ownerUpn)) {
                    continue
                }

                Invoke-WithRetry -OperationName "Add site collection admin $ownerUpn on $siteUrl" -ScriptBlock {
                    Set-SPOUser -Site $siteUrl -LoginName $ownerUpn -IsSiteCollectionAdmin $true -ErrorAction Stop
                }
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'CreateSPOSite' -Status 'Created' -Message "Site created. Secondary owners added: $($uniqueSecondaryOwners.Count)."))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'CreateSPOSite' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($siteUrl) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $siteUrl -Action 'CreateSPOSite' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'SharePoint site creation script completed.' -Level SUCCESS
