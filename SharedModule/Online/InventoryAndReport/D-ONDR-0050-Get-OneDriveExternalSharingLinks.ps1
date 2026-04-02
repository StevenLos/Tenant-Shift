<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-191000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Online.SharePoint.PowerShell

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets OneDriveExternalSharingLinks and exports results to CSV.

.DESCRIPTION
    Gets OneDriveExternalSharingLinks from Microsoft 365 and writes the results to a CSV file.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all objects in scope (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all objects in scope rather than processing from an input CSV file. Uses the DiscoverAll parameter set.

.PARAMETER SharePointAdminUrl
    URL of the SharePoint Online admin centre (e.g. https://contoso-admin.sharepoint.com).

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-IR3208-Get-OneDriveExternalSharingLinks.ps1 -InputCsvPath .\3208.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3208-Get-OneDriveExternalSharingLinks.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Online.SharePoint.PowerShell
    Required roles:   SharePoint Administrator
    Limitations:      None known.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-ONDR-0050-Get-OneDriveExternalSharingLinks_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
        [Parameter(Mandatory)]
        [int]$RowNumber,

        [Parameter(Mandatory)]
        [string]$PrimaryKey,

        [Parameter(Mandatory)]
        [string]$Action,

        [Parameter(Mandatory)]
        [string]$Status,

        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory)]
        [hashtable]$Data
    )

    $base = New-ResultObject -RowNumber $RowNumber -PrimaryKey $PrimaryKey -Action $Action -Status $Status -Message $Message
    $ordered = [ordered]@{}

    foreach ($prop in $base.PSObject.Properties.Name) {
        $ordered[$prop] = $base.$prop
    }

    foreach ($key in $Data.Keys) {
        $ordered[$key] = $Data[$key]
    }

    return [PSCustomObject]$ordered
}

function ConvertTo-OneDriveUrlKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    return (($UserPrincipalName.Trim().ToLowerInvariant()) -replace '[^a-z0-9]', '_')
}

function Get-NormalizedLogin {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$LoginName
    )

    $value = ([string]$LoginName).Trim()
    if ([string]::IsNullOrWhiteSpace($value)) {
        return ''
    }

    if ($value.Contains('|')) {
        $parts = $value.Split('|')
        $value = $parts[$parts.Length - 1]
    }

    return $value.Trim().ToLowerInvariant()
}

function Get-NormalizedEmail {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Email
    )

    return ([string]$Email).Trim().ToLowerInvariant()
}

function Test-IsAccessDeniedMessage {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Message
    )

    $text = ([string]$Message).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $false
    }

    return ($text.Contains('access is denied') -or $text.Contains('e_accessdenied'))
}

function Get-ErrorMessageText {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$ErrorRecords
    )

    $messages = foreach ($record in @($ErrorRecords)) {
        if ($null -eq $record) {
            continue
        }

        $message = ''
        if ($record -is [System.Management.Automation.ErrorRecord]) {
            $message = ([string]$record.Exception.Message).Trim()
            if ([string]::IsNullOrWhiteSpace($message)) {
                $message = ([string]$record).Trim()
            }
        }
        else {
            $message = ([string]$record).Trim()
        }

        if (-not [string]::IsNullOrWhiteSpace($message)) {
            $message
        }
    }

    return ((@($messages | Select-Object -Unique) -join ' | ').Trim())
}

function Get-OneDriveSiteUsersQuiet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    $localErrors = @()
    $users = @()
    $previousErrorActionPreference = $ErrorActionPreference

    try {
        # The imported SPO proxy cmdlets can emit noisy terminating errors for
        # inaccessible personal sites when the session preference is Stop.
        $ErrorActionPreference = 'SilentlyContinue'
        $users = @(Get-SPOUser -Site $SiteUrl -Limit All -ErrorAction SilentlyContinue -ErrorVariable +localErrors 2>$null)
    }
    catch {
        $localErrors += $_
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }

    return [PSCustomObject]@{
        Users  = @($users)
        Errors = @($localErrors)
    }
}

function Get-SiteExternalPrincipalEntries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    $entries = [System.Collections.Generic.List[object]]::new()
    $siteUserLookup = Get-OneDriveSiteUsersQuiet -SiteUrl $SiteUrl
    $siteUsers = @($siteUserLookup.Users)
    $siteUserErrors = @($siteUserLookup.Errors)

    if ($siteUserErrors.Count -gt 0 -and $siteUsers.Count -eq 0) {
        $errorMessage = Get-ErrorMessageText -ErrorRecords $siteUserErrors
        if ([string]::IsNullOrWhiteSpace($errorMessage)) {
            $errorMessage = 'Unable to load site users for OneDrive site.'
        }

        $status = if (Test-IsAccessDeniedMessage -Message $errorMessage) {
            'AccessDenied'
        }
        else {
            'Failed'
        }

        return [PSCustomObject]@{
            Entries      = @()
            Status       = $status
            ErrorMessage = $errorMessage
        }
    }

    foreach ($siteUser in $siteUsers) {
        $loginName = ([string]$siteUser.LoginName).Trim()
        if ([string]::IsNullOrWhiteSpace($loginName)) {
            continue
        }

        $userType = ''
        if ($siteUser.PSObject.Properties.Name -contains 'UserType') {
            $userType = ([string]$siteUser.UserType).Trim()
        }

        $isExternal = $loginName.ToLowerInvariant().Contains('#ext#') -or $userType.Equals('Guest', [System.StringComparison]::OrdinalIgnoreCase)
        if (-not $isExternal) {
            continue
        }

        $emailValue = ''
        if ($siteUser.PSObject.Properties.Name -contains 'Email') {
            $emailValue = ([string]$siteUser.Email).Trim()
        }

        if ([string]::IsNullOrWhiteSpace($emailValue) -and $siteUser.PSObject.Properties.Name -contains 'UserPrincipalName') {
            $emailValue = ([string]$siteUser.UserPrincipalName).Trim()
        }

        $displayName = ''
        if ($siteUser.PSObject.Properties.Name -contains 'DisplayName') {
            $displayName = ([string]$siteUser.DisplayName).Trim()
        }

        $isSiteAdmin = $false
        if ($siteUser.PSObject.Properties.Name -contains 'IsSiteAdmin') {
            $isSiteAdmin = [bool]$siteUser.IsSiteAdmin
        }
        elseif ($siteUser.PSObject.Properties.Name -contains 'IsSiteCollectionAdmin') {
            $isSiteAdmin = [bool]$siteUser.IsSiteCollectionAdmin
        }

        $entries.Add([PSCustomObject]@{
                LoginName       = $loginName
                NormalizedLogin = Get-NormalizedLogin -LoginName $loginName
                Email           = $emailValue
                NormalizedEmail = Get-NormalizedEmail -Email $emailValue
                DisplayName     = $displayName
                UserType        = $userType
                IsSiteAdmin     = $isSiteAdmin
            })
    }

    return [PSCustomObject]@{
        Entries      = @($entries | Sort-Object -Property NormalizedEmail, NormalizedLogin)
        Status       = 'Completed'
        ErrorMessage = ''
    }
}

$requiredHeaders = @(
    'UserPrincipalName'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'UserPrincipalName',
    'OneDriveUrl',
    'ExpectedOneDriveUrl',
    'SiteOwner',
    'ExternalPrincipalLoginName',
    'ExternalPrincipalEmail',
    'ExternalPrincipalDisplayName',
    'ExternalPrincipalType',
    'IsSiteCollectionAdmin',
    'HasExternalAccess'
)

Write-Status -Message 'Starting OneDrive external sharing links inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Online.SharePoint.PowerShell')

if ([string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
    throw 'SharePointAdminUrl is required.'
}

$adminUrlTrimmed = $SharePointAdminUrl.Trim()
if ($adminUrlTrimmed -notmatch '^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$') {
    throw "SharePointAdminUrl '$adminUrlTrimmed' is invalid. Use format: https://<tenant>-admin.sharepoint.com"
}

Ensure-SharePointConnection -AdminUrl $adminUrlTrimmed

$adminUri = [uri]$adminUrlTrimmed
$oneDriveHost = ($adminUri.Host -replace '-admin\.', '-my.')

$scopeMode = 'Csv'
if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. CSV input is bypassed.' -Level WARN

    $discoverRow = [ordered]@{}
    foreach ($header in $requiredHeaders) {
        $discoverRow[$header] = '*'
    }

    $rows = @([PSCustomObject]$discoverRow)
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}
$results = [System.Collections.Generic.List[object]]::new()

Write-Status -Message 'Loading personal sites to map OneDrive URLs by owner and URL key.'
$allSites = @(Invoke-WithRetry -OperationName 'Load personal sites' -ScriptBlock {
    Get-SPOSite -IncludePersonalSite $true -Limit All -Detailed -ErrorAction Stop
})

$personalSites = @($allSites | Where-Object {
    $url = ([string]$_.Url).Trim().ToLowerInvariant()
    $url.Contains('/personal/')
})

$sitesByOwner = @{}
$sitesByUrlKey = @{}

foreach ($site in $personalSites) {
    $siteUrl = ([string]$site.Url).Trim()
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        continue
    }

    $ownerKey = ([string]$site.Owner).Trim().ToLowerInvariant()
    if (-not [string]::IsNullOrWhiteSpace($ownerKey)) {
        if (-not $sitesByOwner.ContainsKey($ownerKey)) {
            $sitesByOwner[$ownerKey] = [System.Collections.Generic.List[object]]::new()
        }
        $sitesByOwner[$ownerKey].Add($site)
    }

    try {
        $uri = [uri]$siteUrl
        $parts = @($uri.AbsolutePath.Trim('/') -split '/')
        if ($parts.Count -ge 2 -and $parts[0].ToLowerInvariant() -eq 'personal') {
            $urlKey = $parts[1].ToLowerInvariant()
            if (-not [string]::IsNullOrWhiteSpace($urlKey)) {
                if (-not $sitesByUrlKey.ContainsKey($urlKey)) {
                    $sitesByUrlKey[$urlKey] = [System.Collections.Generic.List[object]]::new()
                }
                $sitesByUrlKey[$urlKey].Add($site)
            }
        }
    }
    catch {
        # Ignore malformed URLs in site results.
    }
}

$rowNumber = 1
foreach ($row in $rows) {
    $userPrincipalName = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
            throw 'UserPrincipalName is required. Use * to export all discovered OneDrive personal sites.'
        }

        $resolvedSites = @()
        $expectedOneDriveUrl = ''

        if ($userPrincipalName -eq '*') {
            $resolvedSites = @($personalSites)
        }
        else {
            $ownerKey = $userPrincipalName.ToLowerInvariant()
            $urlKey = ConvertTo-OneDriveUrlKey -UserPrincipalName $userPrincipalName
            $expectedOneDriveUrl = "https://$oneDriveHost/personal/$urlKey"

            $matches = [System.Collections.Generic.List[object]]::new()
            if ($sitesByOwner.ContainsKey($ownerKey)) {
                foreach ($site in $sitesByOwner[$ownerKey]) {
                    $matches.Add($site)
                }
            }

            if ($sitesByUrlKey.ContainsKey($urlKey)) {
                foreach ($site in $sitesByUrlKey[$urlKey]) {
                    $matches.Add($site)
                }
            }

            $uniqueSitesByUrl = @{}
            foreach ($site in $matches) {
                $siteUrl = ([string]$site.Url).Trim().ToLowerInvariant()
                if (-not [string]::IsNullOrWhiteSpace($siteUrl) -and -not $uniqueSitesByUrl.ContainsKey($siteUrl)) {
                    $uniqueSitesByUrl[$siteUrl] = $site
                }
            }

            $resolvedSites = @($uniqueSitesByUrl.Values)
        }

        if ($resolvedSites.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'GetOneDriveExternalSharingLinks' -Status 'NotFound' -Message 'No matching OneDrive personal site found for user.' -Data ([ordered]@{
                        UserPrincipalName            = $userPrincipalName
                        OneDriveUrl                  = ''
                        ExpectedOneDriveUrl          = $expectedOneDriveUrl
                        SiteOwner                    = ''
                        ExternalPrincipalLoginName   = ''
                        ExternalPrincipalEmail       = ''
                        ExternalPrincipalDisplayName = ''
                        ExternalPrincipalType        = ''
                        IsSiteCollectionAdmin        = ''
                        HasExternalAccess            = 'False'
                    })))
            $rowNumber++
            continue
        }

        foreach ($site in @($resolvedSites | Sort-Object -Property Owner, Url)) {
            $siteUrl = ([string]$site.Url).Trim()
            $siteOwner = ([string]$site.Owner).Trim()
            $entryLookup = Get-SiteExternalPrincipalEntries -SiteUrl $siteUrl
            $entries = @($entryLookup.Entries)

            if ($entryLookup.Status -ne 'Completed') {
                $primaryKey = if ($userPrincipalName -eq '*') { $siteUrl } else { "$userPrincipalName|$siteUrl" }
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetOneDriveExternalSharingLinks' -Status $entryLookup.Status -Message "Unable to load external principals for OneDrive site. Error: $($entryLookup.ErrorMessage)" -Data ([ordered]@{
                            UserPrincipalName            = if ($userPrincipalName -eq '*') { $siteOwner } else { $userPrincipalName }
                            OneDriveUrl                  = $siteUrl
                            ExpectedOneDriveUrl          = if ($userPrincipalName -eq '*') { '' } else { $expectedOneDriveUrl }
                            SiteOwner                    = $siteOwner
                            ExternalPrincipalLoginName   = ''
                            ExternalPrincipalEmail       = ''
                            ExternalPrincipalDisplayName = ''
                            ExternalPrincipalType        = ''
                            IsSiteCollectionAdmin        = ''
                            HasExternalAccess            = ''
                        })))
                continue
            }

            if ($entries.Count -eq 0) {
                $primaryKey = if ($userPrincipalName -eq '*') { $siteUrl } else { $userPrincipalName }
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetOneDriveExternalSharingLinks' -Status 'Completed' -Message 'No external principals found for this OneDrive site.' -Data ([ordered]@{
                            UserPrincipalName            = if ($userPrincipalName -eq '*') { $siteOwner } else { $userPrincipalName }
                            OneDriveUrl                  = $siteUrl
                            ExpectedOneDriveUrl          = if ($userPrincipalName -eq '*') { '' } else { $expectedOneDriveUrl }
                            SiteOwner                    = $siteOwner
                            ExternalPrincipalLoginName   = ''
                            ExternalPrincipalEmail       = ''
                            ExternalPrincipalDisplayName = ''
                            ExternalPrincipalType        = ''
                            IsSiteCollectionAdmin        = ''
                            HasExternalAccess            = 'False'
                        })))
                continue
            }

            foreach ($entry in $entries) {
                $baseKey = if ($userPrincipalName -eq '*') { $siteUrl } else { $userPrincipalName }
                $identityKey = if (-not [string]::IsNullOrWhiteSpace($entry.NormalizedEmail)) { $entry.NormalizedEmail } else { $entry.NormalizedLogin }

                $message = if ($userPrincipalName -ne '*' -and $resolvedSites.Count -gt 1) {
                    'Multiple OneDrive site matches found; external principal row exported for each match.'
                }
                else {
                    'External sharing principal exported.'
                }

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$baseKey|$identityKey" -Action 'GetOneDriveExternalSharingLinks' -Status 'Completed' -Message $message -Data ([ordered]@{
                            UserPrincipalName            = if ($userPrincipalName -eq '*') { $siteOwner } else { $userPrincipalName }
                            OneDriveUrl                  = $siteUrl
                            ExpectedOneDriveUrl          = if ($userPrincipalName -eq '*') { '' } else { $expectedOneDriveUrl }
                            SiteOwner                    = $siteOwner
                            ExternalPrincipalLoginName   = $entry.LoginName
                            ExternalPrincipalEmail       = $entry.Email
                            ExternalPrincipalDisplayName = $entry.DisplayName
                            ExternalPrincipalType        = if ([string]::IsNullOrWhiteSpace($entry.UserType)) { 'Guest' } else { $entry.UserType }
                            IsSiteCollectionAdmin        = [string]$entry.IsSiteAdmin
                            HasExternalAccess            = 'True'
                        })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($userPrincipalName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrincipalName -Action 'GetOneDriveExternalSharingLinks' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    UserPrincipalName            = $userPrincipalName
                    OneDriveUrl                  = ''
                    ExpectedOneDriveUrl          = ''
                    SiteOwner                    = ''
                    ExternalPrincipalLoginName   = ''
                    ExternalPrincipalEmail       = ''
                    ExternalPrincipalDisplayName = ''
                    ExternalPrincipalType        = ''
                    IsSiteCollectionAdmin        = ''
                    HasExternalAccess            = ''
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
Write-Status -Message 'OneDrive external sharing links inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}




