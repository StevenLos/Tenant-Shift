<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260302-191500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
ActiveDirectory

.MODULEVERSIONPOLICY
Best-effort PSGallery check (offline-safe)
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\..\Standalone_OutputCsvPath') -ChildPath ("Results_SA-M0002-Update-ActiveDirectoryContacts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest


function Write-Status {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'INFO' { 'Cyan' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        'SUCCESS' { 'Green' }
    }

    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Start-RunTranscript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OutputCsvPath,

        [AllowNull()]
        [string]$ScriptPath
    )

    $directory = Split-Path -Path $OutputCsvPath -Parent
    if ([string]::IsNullOrWhiteSpace($directory) -and -not [string]::IsNullOrWhiteSpace($ScriptPath)) {
        $directory = Split-Path -Path $ScriptPath -Parent
    }

    if ([string]::IsNullOrWhiteSpace($directory)) {
        throw "Unable to determine transcript directory from OutputCsvPath '$OutputCsvPath'."
    }

    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $scriptName = 'Script'
    if (-not [string]::IsNullOrWhiteSpace($ScriptPath)) {
        $candidate = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            $scriptName = $candidate
        }
    }

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $transcriptPath = Join-Path -Path $directory -ChildPath ("Transcript_{0}_{1}.log" -f $scriptName, $timestamp)

    try {
        Start-Transcript -LiteralPath $transcriptPath -Force -ErrorAction Stop | Out-Null
    }
    catch {
        throw "Failed to start transcript at '$transcriptPath'. Error: $($_.Exception.Message)"
    }

    Write-Status -Message "Transcript started at '$transcriptPath'."
    return $transcriptPath
}

function Stop-RunTranscript {
    [CmdletBinding()]
    param()

    try {
        Stop-Transcript -ErrorAction Stop | Out-Null
    }
    catch {
        $message = ([string]$_.Exception.Message).ToLowerInvariant()
        if ($message -notmatch 'not currently transcribing') {
            throw
        }
    }
}

function ConvertTo-Bool {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value,

        [bool]$Default = $false
    )

    if ($null -eq $Value) {
        return $Default
    }

    $stringValue = [string]$Value
    if ([string]::IsNullOrWhiteSpace($stringValue)) {
        return $Default
    }

    switch -Regex ($stringValue.Trim().ToLowerInvariant()) {
        '^(1|true|t|yes|y)$' { return $true }
        '^(0|false|f|no|n)$' { return $false }
        default { throw "Invalid boolean value '$stringValue'. Use true/false or yes/no." }
    }
}

function ConvertTo-Array {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value,

        [string]$Delimiter = ';'
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    return @(
        $Value -split [Regex]::Escape($Delimiter) |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Get-TrimmedValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value).Trim()
}

function Get-NullableBool {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return (ConvertTo-Bool -Value $text)
}

function Assert-ModuleCurrent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ModuleNames,

        [switch]$FailOnOutdated,

        [switch]$FailOnGalleryLookupError
    )

    foreach ($moduleName in $ModuleNames) {
        Write-Status -Message "Checking module '$moduleName'."

        $installed = Get-Module -ListAvailable -Name $moduleName |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if (-not $installed) {
            throw "Required module '$moduleName' is not installed."
        }

        Write-Status -Message "Installed version for '$moduleName': $($installed.Version)."

        try {
            $gallery = Find-Module -Name $moduleName -Repository PSGallery -ErrorAction Stop
        }
        catch {
            if ($FailOnGalleryLookupError) {
                throw "Unable to verify the latest version for '$moduleName' from PSGallery. Error: $($_.Exception.Message)"
            }

            Write-Status -Message "PSGallery lookup unavailable for '$moduleName'. Continuing with installed version check only." -Level WARN
            continue
        }

        if ($installed.Version -lt $gallery.Version) {
            $message = "Module '$moduleName' is outdated. Installed: $($installed.Version), current: $($gallery.Version)."
            if ($FailOnOutdated) {
                throw "$message Update with: Update-Module $moduleName"
            }

            Write-Status -Message $message -Level WARN
        }
        else {
            Write-Status -Message "Module '$moduleName' is current ($($installed.Version))." -Level SUCCESS
        }
    }
}

function Import-ValidatedCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$InputCsvPath,

        [Parameter(Mandatory)]
        [string[]]$RequiredHeaders
    )

    if (-not (Test-Path -LiteralPath $InputCsvPath -PathType Leaf)) {
        throw "Input CSV file not found: $InputCsvPath"
    }

    $firstLine = Get-Content -LiteralPath $InputCsvPath -TotalCount 1
    if ([string]::IsNullOrWhiteSpace($firstLine)) {
        throw "CSV file '$InputCsvPath' is missing a header row."
    }

    $rawHeaders = @($firstLine -split ',')
    $headers = [System.Collections.Generic.List[string]]::new()
    foreach ($rawHeader in $rawHeaders) {
        $cleanHeader = ([string]$rawHeader).Trim().Trim('"').TrimStart([char]0xFEFF)
        $headers.Add($cleanHeader)
    }

    if ($headers.Count -eq 0) {
        throw "CSV file '$InputCsvPath' is missing a header row."
    }

    $duplicates = @($headers | Group-Object | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Name)
    if ($duplicates.Count -gt 0) {
        throw "CSV file '$InputCsvPath' contains duplicate headers: $($duplicates -join ', ')"
    }

    $missing = @($RequiredHeaders | Where-Object { $_ -notin $headers })
    if ($missing.Count -gt 0) {
        throw "CSV file '$InputCsvPath' is missing required headers: $($missing -join ', ')"
    }

    $rows = Import-Csv -LiteralPath $InputCsvPath
    if (-not $rows -or @($rows).Count -eq 0) {
        throw "CSV file '$InputCsvPath' has no data rows."
    }

    return @($rows)
}
function Ensure-ActiveDirectoryConnection {
    [CmdletBinding()]
    param()

    $isWindowsHost = $false
    $isWindowsVar = Get-Variable -Name IsWindows -ErrorAction SilentlyContinue
    if ($null -ne $isWindowsVar) {
        $isWindowsHost = [bool]$isWindowsVar.Value
    }
    else {
        $isWindowsHost = ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT)
    }

    if (-not $isWindowsHost) {
        throw 'ActiveDirectory scripts require Windows with RSAT/AD tooling available.'
    }

    Assert-ModuleCurrent -ModuleNames @('ActiveDirectory')

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch {
        throw "Unable to import ActiveDirectory module. Error: $($_.Exception.Message)"
    }

    try {
        Get-ADDomain -ErrorAction Stop | Out-Null
    }
    catch {
        throw "Unable to query Active Directory domain context. Ensure domain connectivity and permissions. Error: $($_.Exception.Message)"
    }

    Write-Status -Message 'Active Directory module loaded and domain context verified.' -Level SUCCESS
}

function Escape-AdFilterValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    return $Value.Replace("'", "''")
}

function ConvertTo-NullableDateTime {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    try {
        return [datetime]$text
    }
    catch {
        throw "Invalid datetime value '$text'. Use an ISA-like value (for example 2026-03-02 or 2026-03-02T10:30:00)."
    }
}

function Get-HttpStatusCodeFromException {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Exception]$Exception
    )

    foreach ($propertyName in @('ResponseStatusCode', 'StatusCode', 'HttpStatusCode')) {
        if ($Exception.PSObject.Properties.Name -contains $propertyName) {
            $rawValue = $Exception.$propertyName
            if ($null -eq $rawValue) {
                continue
            }

            try {
                return [int]$rawValue
            }
            catch {
                # Continue searching.
            }
        }
    }

    return $null
}

function Test-IsTransientException {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Exception]$Exception
    )

    $statusCode = Get-HttpStatusCodeFromException -Exception $Exception
    if ($null -ne $statusCode -and ($statusCode -eq 429 -or $statusCode -ge 500)) {
        return $true
    }

    $messageChain = [System.Collections.Generic.List[string]]::new()
    $cursor = $Exception
    while ($null -ne $cursor) {
        if (-not [string]::IsNullOrWhiteSpace($cursor.Message)) {
            $messageChain.Add($cursor.Message)
        }

        $cursor = $cursor.InnerException
    }

    $combinedMessage = ($messageChain -join ' ').ToLowerInvariant()
    return ($combinedMessage -match 'temporar|timeout|timed out|service unavailable|rate limit|try again|gateway|429|500|502|503|504|server is not operational')
}

function Get-RetryDelaySeconds {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Exception]$Exception,

        [Parameter(Mandatory)]
        [int]$AttemptNumber,

        [int]$BaseDelaySeconds = 2,
        [int]$MaxDelaySeconds = 60
    )

    $rawDelay = [Math]::Pow(2, [Math]::Min($AttemptNumber, 6)) * $BaseDelaySeconds
    $jitter = Get-Random -Minimum 0 -Maximum 3
    $delay = [int]([Math]::Min($rawDelay + $jitter, $MaxDelaySeconds))
    return [Math]::Max($delay, 1)
}

function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory)]
        [string]$OperationName,

        [ValidateRange(1, 15)]
        [int]$MaxAttempts = 5
    )

    $attempt = 1
    while ($attempt -le $MaxAttempts) {
        try {
            return & $ScriptBlock
        }
        catch {
            $exception = $_.Exception
            if ($attempt -ge $MaxAttempts -or -not (Test-IsTransientException -Exception $exception)) {
                throw
            }

            $delaySeconds = Get-RetryDelaySeconds -Exception $exception -AttemptNumber $attempt
            Write-Status -Level WARN -Message "Transient error during '$OperationName' (attempt $attempt/$MaxAttempts): $($exception.Message). Retrying in $delaySeconds second(s)."
            Start-Sleep -Seconds $delaySeconds
            $attempt++
        }
    }
}

function New-ResultObject {
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
        [string]$Message
    )

    return [PSCustomObject]@{
        TimestampUtc = (Get-Date).ToUniversalTime().ToString('o')
        RowNumber    = $RowNumber
        PrimaryKey   = $PrimaryKey
        Action       = $Action
        Status       = $Status
        Message      = $Message
    }
}

function Export-ResultsCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Results,

        [Parameter(Mandatory)]
        [string]$OutputCsvPath
    )

    $directory = Split-Path -Path $OutputCsvPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $Results | Export-Csv -LiteralPath $OutputCsvPath -NoTypeInformation -Encoding UTF8
    Write-Status -Message "Results exported to '$OutputCsvPath'." -Level SUCCESS
}


$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Test-IsAdContact {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$AdObject
    )

    if ($null -eq $AdObject) {
        return $false
    }

    $classes = @($AdObject.ObjectClass | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() })
    return $classes -contains 'contact'
}

function Resolve-TargetAdContact {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IdentityType,

        [Parameter(Mandatory)]
        [string]$IdentityValue
    )

    $normalizedType = $IdentityType.Trim().ToLowerInvariant()
    switch ($normalizedType) {
        'name' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADObject -Filter "ObjectClass -eq 'contact' -and Name -eq '$escaped'" -Properties * -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'mail' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADObject -Filter "ObjectClass -eq 'contact' -and mail -eq '$escaped'" -Properties * -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'proxyaddress' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADObject -Filter "ObjectClass -eq 'contact' -and proxyAddresses -eq '$escaped'" -Properties * -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'distinguishedname' {
            $candidate = Get-ADObject -Identity $IdentityValue -Properties * -ErrorAction SilentlyContinue
            if (Test-IsAdContact -AdObject $candidate) {
                return $candidate
            }

            return $null
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            $candidate = Get-ADObject -Identity $guid -Properties * -ErrorAction SilentlyContinue
            if (Test-IsAdContact -AdObject $candidate) {
                return $candidate
            }

            return $null
        }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use Name, Mail, ProxyAddress, DistinguishedName, or ObjectGuid."
        }
    }
}

function Add-ReplaceField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$ReplaceAttributes,

        [Parameter(Mandatory)]
        [string]$AttributeName,

        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        $ReplaceAttributes[$AttributeName] = $text
        return $true
    }

    return $false
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'IdentityType',
    'IdentityValue',
    'ClearAttributes',
    'Name',
    'DisplayName',
    'GivenName',
    'Initials',
    'Surname',
    'Description',
    'Company',
    'Department',
    'Division',
    'Title',
    'Office',
    'Manager',
    'OfficePhone',
    'MobilePhone',
    'HomePhone',
    'IpPhone',
    'Fax',
    'Pager',
    'Mail',
    'MailNickname',
    'ProxyAddresses',
    'TargetAddress',
    'HideFromAddressLists',
    'StreetAddress',
    'PostOfficeBox',
    'City',
    'StateOrProvince',
    'PostalCode',
    'CountryCode',
    'CountryName',
    'CountryNumericCode',
    'WebPage',
    'ExtensionAttribute1',
    'ExtensionAttribute2',
    'ExtensionAttribute3',
    'ExtensionAttribute4',
    'ExtensionAttribute5',
    'ExtensionAttribute6',
    'ExtensionAttribute7',
    'ExtensionAttribute8',
    'ExtensionAttribute9',
    'ExtensionAttribute10',
    'ExtensionAttribute11',
    'ExtensionAttribute12',
    'ExtensionAttribute13',
    'ExtensionAttribute14',
    'ExtensionAttribute15',
    'Path'
)

$clearAttributeMap = @{
    DisplayName = 'displayName'
    GivenName = 'givenName'
    Initials = 'initials'
    Surname = 'sn'
    Description = 'description'
    Company = 'company'
    Department = 'department'
    Division = 'division'
    Title = 'title'
    Office = 'physicalDeliveryOfficeName'
    Manager = 'manager'
    OfficePhone = 'telephoneNumber'
    MobilePhone = 'mobile'
    HomePhone = 'homePhone'
    IpPhone = 'ipPhone'
    Fax = 'facsimileTelephoneNumber'
    Pager = 'pager'
    Mail = 'mail'
    MailNickname = 'mailNickname'
    ProxyAddresses = 'proxyAddresses'
    TargetAddress = 'targetAddress'
    HideFromAddressLists = 'msExchHideFromAddressLists'
    StreetAddress = 'streetAddress'
    PostOfficeBox = 'postOfficeBox'
    City = 'l'
    StateOrProvince = 'st'
    PostalCode = 'postalCode'
    CountryCode = 'c'
    CountryName = 'co'
    CountryNumericCode = 'countryCode'
    WebPage = 'wWWHomePage'
}

for ($i = 1; $i -le 15; $i++) {
    $clearAttributeMap["ExtensionAttribute$i"] = "extensionAttribute$i"
}

Write-Status -Message 'Starting Active Directory contact update script.'
Ensure-ActiveDirectoryConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $identityType = Get-TrimmedValue -Value $row.IdentityType
    $identityValue = Get-TrimmedValue -Value $row.IdentityValue
    $primaryKey = "${identityType}:$identityValue"

    try {
        if ([string]::IsNullOrWhiteSpace($identityType) -or [string]::IsNullOrWhiteSpace($identityValue)) {
            throw 'IdentityType and IdentityValue are required.'
        }

        $targetContact = Invoke-WithRetry -OperationName "Resolve AD contact $primaryKey" -ScriptBlock {
            Resolve-TargetAdContact -IdentityType $identityType -IdentityValue $identityValue
        }

        if (-not $targetContact) {
            throw 'Target contact was not found.'
        }

        $resolvedKey = if (-not [string]::IsNullOrWhiteSpace((Get-TrimmedValue -Value $targetContact.mail))) { (Get-TrimmedValue -Value $targetContact.mail) } else { (Get-TrimmedValue -Value $targetContact.Name) }
        $messages = [System.Collections.Generic.List[string]]::new()
        $changeCount = 0

        $setParams = @{
            Identity = $targetContact.ObjectGuid
        }

        $replaceAttributes = @{}
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'displayName' -Value $row.DisplayName) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'givenName' -Value $row.GivenName) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'initials' -Value $row.Initials) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'sn' -Value $row.Surname) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'description' -Value $row.Description) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'company' -Value $row.Company) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'department' -Value $row.Department) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'division' -Value $row.Division) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'title' -Value $row.Title) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'physicalDeliveryOfficeName' -Value $row.Office) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'manager' -Value $row.Manager) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'telephoneNumber' -Value $row.OfficePhone) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'mobile' -Value $row.MobilePhone) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'homePhone' -Value $row.HomePhone) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'ipPhone' -Value $row.IpPhone) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'facsimileTelephoneNumber' -Value $row.Fax) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'pager' -Value $row.Pager) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'mail' -Value $row.Mail) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'mailNickname' -Value $row.MailNickname) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'targetAddress' -Value $row.TargetAddress) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'streetAddress' -Value $row.StreetAddress) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'postOfficeBox' -Value $row.PostOfficeBox) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'l' -Value $row.City) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'st' -Value $row.StateOrProvince) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'postalCode' -Value $row.PostalCode) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'c' -Value $row.CountryCode) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'co' -Value $row.CountryName) { $changeCount++ }
        if (Add-ReplaceField -ReplaceAttributes $replaceAttributes -AttributeName 'wWWHomePage' -Value $row.WebPage) { $changeCount++ }

        $proxyAddresses = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ProxyAddresses)
        if ($proxyAddresses.Count -gt 0) {
            $replaceAttributes['proxyAddresses'] = [string[]]$proxyAddresses
            $changeCount++
        }

        $hideFromAddressLists = Get-NullableBool -Value $row.HideFromAddressLists
        if ($null -ne $hideFromAddressLists) {
            $replaceAttributes['msExchHideFromAddressLists'] = $hideFromAddressLists
            $changeCount++
        }

        $countryNumericCode = Get-TrimmedValue -Value $row.CountryNumericCode
        if (-not [string]::IsNullOrWhiteSpace($countryNumericCode)) {
            try {
                $replaceAttributes['countryCode'] = [int]$countryNumericCode
                $changeCount++
            }
            catch {
                throw "CountryNumericCode '$countryNumericCode' must be an integer value."
            }
        }

        for ($i = 1; $i -le 15; $i++) {
            $columnName = "ExtensionAttribute$i"
            $attributeName = "extensionAttribute$i"
            $value = Get-TrimmedValue -Value $row.$columnName
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $replaceAttributes[$attributeName] = $value
                $changeCount++
            }
        }

        if ($replaceAttributes.Count -gt 0) {
            $setParams['Replace'] = $replaceAttributes
        }

        $clearAttributes = [System.Collections.Generic.List[string]]::new()
        $clearRequested = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ClearAttributes)
        foreach ($clearRequestedName in $clearRequested) {
            if ($clearAttributeMap.ContainsKey($clearRequestedName)) {
                $mapped = $clearAttributeMap[$clearRequestedName]
                if (-not $clearAttributes.Contains($mapped)) {
                    $clearAttributes.Add($mapped)
                }
            }
            else {
                if (-not $clearAttributes.Contains($clearRequestedName)) {
                    $clearAttributes.Add($clearRequestedName)
                }
            }
        }

        if ($clearAttributes.Count -gt 0) {
            $setParams['Clear'] = @($clearAttributes)
            $changeCount++
        }

        if ($setParams.Count -gt 1) {
            if ($PSCmdlet.ShouldProcess($resolvedKey, 'Update Active Directory contact attributes')) {
                Invoke-WithRetry -OperationName "Update AD contact attributes $resolvedKey" -ScriptBlock {
                    Set-ADObject @setParams -ErrorAction Stop
                }

                $messages.Add('Attributes updated.')
            }
            else {
                $messages.Add('Attribute updates skipped due to WhatIf.')
            }
        }

        $newName = Get-TrimmedValue -Value $row.Name
        if (-not [string]::IsNullOrWhiteSpace($newName) -and $newName -ne (Get-TrimmedValue -Value $targetContact.Name)) {
            $changeCount++

            if ($PSCmdlet.ShouldProcess($resolvedKey, "Rename AD contact to '$newName'")) {
                Invoke-WithRetry -OperationName "Rename AD contact $resolvedKey" -ScriptBlock {
                    Rename-ADObject -Identity $targetContact.ObjectGuid -NewName $newName -ErrorAction Stop
                }

                $messages.Add("Contact renamed to '$newName'.")
            }
            else {
                $messages.Add('Rename skipped due to WhatIf.')
            }
        }

        $targetPath = Get-TrimmedValue -Value $row.Path
        if (-not [string]::IsNullOrWhiteSpace($targetPath)) {
            $currentContact = Invoke-WithRetry -OperationName "Reload AD contact $resolvedKey" -ScriptBlock {
                Get-ADObject -Identity $targetContact.ObjectGuid -Properties DistinguishedName -ErrorAction Stop
            }

            $currentParentPath = ($currentContact.DistinguishedName -split ',', 2)[1]
            if ($targetPath -ieq $currentParentPath) {
                $messages.Add('Contact already in requested OU path.')
            }
            else {
                $changeCount++
                if ($PSCmdlet.ShouldProcess($resolvedKey, "Move AD contact to '$targetPath'")) {
                    Invoke-WithRetry -OperationName "Move AD contact $resolvedKey" -ScriptBlock {
                        Move-ADObject -Identity $targetContact.ObjectGuid -TargetPath $targetPath -ErrorAction Stop
                    }

                    $messages.Add("Contact moved to '$targetPath'.")
                }
                else {
                    $messages.Add('OU move skipped due to WhatIf.')
                }
            }
        }

        if ($changeCount -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'UpdateActiveDirectoryContact' -Status 'Skipped' -Message 'No changes were requested for this row.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'UpdateActiveDirectoryContact' -Status 'Completed' -Message ($messages -join ' ')))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'UpdateActiveDirectoryContact' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory contact update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}


