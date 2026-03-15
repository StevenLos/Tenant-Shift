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

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$SearchBase,

    [Parameter(ParameterSetName = 'FromCsv')]
    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$Server,

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [ValidateRange(0, 10000000)]
    [int]$MaxObjects = 0,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\..\Standalone_OutputCsvPath') -ChildPath ("Results_SA-IR0002-Get-ActiveDirectoryContacts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Escape-LdapFilterValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    $builder = [System.Text.StringBuilder]::new()
    foreach ($char in $Value.ToCharArray()) {
        switch ($char) {
            '\\' { [void]$builder.Append('\\5c') }
            '*' { [void]$builder.Append('\\2a') }
            '(' { [void]$builder.Append('\\28') }
            ')' { [void]$builder.Append('\\29') }
            ([char]0) { [void]$builder.Append('\\00') }
            default { [void]$builder.Append($char) }
        }
    }

    return $builder.ToString()
}

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

function Resolve-ContactsByScope {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IdentityType,

        [Parameter(Mandatory)]
        [string]$IdentityValue,

        [Parameter(Mandatory)]
        [string[]]$PropertyNames,

        [AllowEmptyString()]
        [string]$SearchBase,

        [AllowEmptyString()]
        [string]$Server
    )

    $normalizedType = $IdentityType.Trim().ToLowerInvariant()

    switch ($normalizedType) {
        'all' {
            if ($IdentityValue -ne '*') {
                throw "IdentityValue must be '*' when IdentityType is 'All'."
            }

            $params = @{
                LDAPFilter = '(objectClass=contact)'
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADObject @params)
        }
        'name' {
            $escaped = Escape-LdapFilterValue -Value $IdentityValue
            $params = @{
                LDAPFilter = "(&(objectClass=contact)(name=$escaped))"
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADObject @params)
        }
        'mail' {
            $escaped = Escape-LdapFilterValue -Value $IdentityValue
            $params = @{
                LDAPFilter = "(&(objectClass=contact)(mail=$escaped))"
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADObject @params)
        }
        'distinguishedname' {
            $params = @{
                Identity = $IdentityValue
                Properties  = '*'
                ErrorAction = 'SilentlyContinue'
            }

            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            $contact = Get-ADObject @params
            if ($contact -and (Test-IsAdContact -AdObject $contact)) {
                return @($contact)
            }

            return @()
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            $params = @{
                Identity = $guid
                Properties  = '*'
                ErrorAction = 'SilentlyContinue'
            }

            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            $contact = Get-ADObject @params
            if ($contact -and (Test-IsAdContact -AdObject $contact)) {
                return @($contact)
            }

            return @()
        }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use All, Name, Mail, DistinguishedName, or ObjectGuid."
        }
    }
}

function Convert-MultiValueToString {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [string]) {
        return ([string]$Value).Trim()
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        $items = [System.Collections.Generic.List[string]]::new()
        foreach ($item in $Value) {
            $text = ([string]$item).Trim()
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $items.Add($text)
            }
        }

        return (@($items | Sort-Object -Unique) -join ';')
    }

    return ([string]$Value).Trim()
}

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

$requiredHeaders = @(
    'IdentityType',
    'IdentityValue'
)

$propertyNames = @(
    'DistinguishedName',
    'ObjectGuid',
    'objectClass',
    'Name',
    'displayName',
    'givenName',
    'initials',
    'sn',
    'description',
    'company',
    'department',
    'division',
    'title',
    'physicalDeliveryOfficeName',
    'manager',
    'telephoneNumber',
    'mobile',
    'homePhone',
    'ipPhone',
    'facsimileTelephoneNumber',
    'pager',
    'mail',
    'mailNickname',
    'proxyAddresses',
    'targetAddress',
    'msExchHideFromAddressLists',
    'streetAddress',
    'postOfficeBox',
    'l',
    'st',
    'postalCode',
    'c',
    'co',
    'countryCode',
    'wWWHomePage',
    'extensionAttribute1',
    'extensionAttribute2',
    'extensionAttribute3',
    'extensionAttribute4',
    'extensionAttribute5',
    'extensionAttribute6',
    'extensionAttribute7',
    'extensionAttribute8',
    'extensionAttribute9',
    'extensionAttribute10',
    'extensionAttribute11',
    'extensionAttribute12',
    'extensionAttribute13',
    'extensionAttribute14',
    'extensionAttribute15',
    'CanonicalName',
    'whenCreated',
    'whenChanged',
    'showInAddressBook',
    'legacyExchangeDN',
    'memberOf'
)

function New-EmptyContactData {
    return [ordered]@{
        IdentityTypeRequested = ''
        IdentityValueRequested = ''
        DistinguishedName = ''
        ObjectGuid = ''
        ObjectClass = ''
        Name = ''
        DisplayName = ''
        GivenName = ''
        Initials = ''
        Surname = ''
        Description = ''
        Company = ''
        Department = ''
        Division = ''
        Title = ''
        Office = ''
        Manager = ''
        OfficePhone = ''
        MobilePhone = ''
        HomePhone = ''
        IpPhone = ''
        Fax = ''
        Pager = ''
        Mail = ''
        MailNickname = ''
        ProxyAddresses = ''
        TargetAddress = ''
        HideFromAddressLists = ''
        StreetAddress = ''
        PostOfficeBox = ''
        City = ''
        StateOrProvince = ''
        PostalCode = ''
        CountryCode = ''
        CountryName = ''
        CountryNumericCode = ''
        WebPage = ''
        ExtensionAttribute1 = ''
        ExtensionAttribute2 = ''
        ExtensionAttribute3 = ''
        ExtensionAttribute4 = ''
        ExtensionAttribute5 = ''
        ExtensionAttribute6 = ''
        ExtensionAttribute7 = ''
        ExtensionAttribute8 = ''
        ExtensionAttribute9 = ''
        ExtensionAttribute10 = ''
        ExtensionAttribute11 = ''
        ExtensionAttribute12 = ''
        ExtensionAttribute13 = ''
        ExtensionAttribute14 = ''
        ExtensionAttribute15 = ''
        CanonicalName = ''
        WhenCreated = ''
        WhenChanged = ''
        ShowInAddressBook = ''
        LegacyExchangeDn = ''
        MemberOf = ''
    }
}

Write-Status -Message 'Starting Active Directory contact inventory script.'
Ensure-ActiveDirectoryConnection

$scopeMode = 'Csv'
$resolvedServer = Get-TrimmedValue -Value $Server
$resolvedSearchBase = ''
$runWasTruncated = $false

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    $resolvedSearchBase = Get-TrimmedValue -Value $SearchBase
    if ([string]::IsNullOrWhiteSpace($resolvedSearchBase)) {
        $domainParams = @{
            ErrorAction = 'Stop'
        }

        if (-not [string]::IsNullOrWhiteSpace($resolvedServer)) {
            $domainParams['Server'] = $resolvedServer
        }

        $resolvedSearchBase = (Get-ADDomain @domainParams).DistinguishedName
    }

    if ([string]::IsNullOrWhiteSpace($resolvedSearchBase)) {
        throw 'Unable to determine SearchBase for DiscoverAll mode.'
    }

    Write-Status -Message "DiscoverAll enabled for Active Directory contacts. SearchBase='$resolvedSearchBase'." -Level WARN
    $rows = @([PSCustomObject]@{
            IdentityType = 'All'
            IdentityValue = '*'
        })
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

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

        $effectiveSearchBase = if ($scopeMode -eq 'DiscoverAll') { $resolvedSearchBase } else { '' }
        $contacts = Invoke-WithRetry -OperationName "Load contacts for $primaryKey" -ScriptBlock {
            Resolve-ContactsByScope -IdentityType $identityType -IdentityValue $identityValue -PropertyNames $propertyNames -SearchBase $effectiveSearchBase -Server $resolvedServer
        }

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $contacts.Count -gt $MaxObjects) {
            $contacts = @($contacts | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($contacts.Count -eq 0) {
            $data = New-EmptyContactData
            $data['IdentityTypeRequested'] = $identityType
            $data['IdentityValueRequested'] = $identityValue
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryContact' -Status 'NotFound' -Message 'No matching contacts were found.' -Data $data))

            $rowNumber++
            continue
        }

        foreach ($contact in @($contacts | Sort-Object -Property mail, Name, DistinguishedName)) {
            $contactPrimaryKey = if (-not [string]::IsNullOrWhiteSpace((Get-TrimmedValue -Value $contact.mail))) { (Get-TrimmedValue -Value $contact.mail) } else { (Get-TrimmedValue -Value $contact.Name) }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $contactPrimaryKey -Action 'GetActiveDirectoryContact' -Status 'Completed' -Message 'Contact exported.' -Data ([ordered]@{
                            IdentityTypeRequested = $identityType
                            IdentityValueRequested = $identityValue
                            DistinguishedName = Get-TrimmedValue -Value $contact.DistinguishedName
                            ObjectGuid = Get-TrimmedValue -Value $contact.ObjectGuid
                            ObjectClass = Convert-MultiValueToString -Value $contact.objectClass
                            Name = Get-TrimmedValue -Value $contact.Name
                            DisplayName = Get-TrimmedValue -Value $contact.displayName
                            GivenName = Get-TrimmedValue -Value $contact.givenName
                            Initials = Get-TrimmedValue -Value $contact.initials
                            Surname = Get-TrimmedValue -Value $contact.sn
                            Description = Get-TrimmedValue -Value $contact.description
                            Company = Get-TrimmedValue -Value $contact.company
                            Department = Get-TrimmedValue -Value $contact.department
                            Division = Get-TrimmedValue -Value $contact.division
                            Title = Get-TrimmedValue -Value $contact.title
                            Office = Get-TrimmedValue -Value $contact.physicalDeliveryOfficeName
                            Manager = Get-TrimmedValue -Value $contact.manager
                            OfficePhone = Get-TrimmedValue -Value $contact.telephoneNumber
                            MobilePhone = Get-TrimmedValue -Value $contact.mobile
                            HomePhone = Get-TrimmedValue -Value $contact.homePhone
                            IpPhone = Get-TrimmedValue -Value $contact.ipPhone
                            Fax = Get-TrimmedValue -Value $contact.facsimileTelephoneNumber
                            Pager = Get-TrimmedValue -Value $contact.pager
                            Mail = Get-TrimmedValue -Value $contact.mail
                            MailNickname = Get-TrimmedValue -Value $contact.mailNickname
                            ProxyAddresses = Convert-MultiValueToString -Value $contact.proxyAddresses
                            TargetAddress = Get-TrimmedValue -Value $contact.targetAddress
                            HideFromAddressLists = [string]$contact.msExchHideFromAddressLists
                            StreetAddress = Get-TrimmedValue -Value $contact.streetAddress
                            PostOfficeBox = Get-TrimmedValue -Value $contact.postOfficeBox
                            City = Get-TrimmedValue -Value $contact.l
                            StateOrProvince = Get-TrimmedValue -Value $contact.st
                            PostalCode = Get-TrimmedValue -Value $contact.postalCode
                            CountryCode = Get-TrimmedValue -Value $contact.c
                            CountryName = Get-TrimmedValue -Value $contact.co
                            CountryNumericCode = Get-TrimmedValue -Value $contact.countryCode
                            WebPage = Get-TrimmedValue -Value $contact.wWWHomePage
                            ExtensionAttribute1 = Get-TrimmedValue -Value $contact.extensionAttribute1
                            ExtensionAttribute2 = Get-TrimmedValue -Value $contact.extensionAttribute2
                            ExtensionAttribute3 = Get-TrimmedValue -Value $contact.extensionAttribute3
                            ExtensionAttribute4 = Get-TrimmedValue -Value $contact.extensionAttribute4
                            ExtensionAttribute5 = Get-TrimmedValue -Value $contact.extensionAttribute5
                            ExtensionAttribute6 = Get-TrimmedValue -Value $contact.extensionAttribute6
                            ExtensionAttribute7 = Get-TrimmedValue -Value $contact.extensionAttribute7
                            ExtensionAttribute8 = Get-TrimmedValue -Value $contact.extensionAttribute8
                            ExtensionAttribute9 = Get-TrimmedValue -Value $contact.extensionAttribute9
                            ExtensionAttribute10 = Get-TrimmedValue -Value $contact.extensionAttribute10
                            ExtensionAttribute11 = Get-TrimmedValue -Value $contact.extensionAttribute11
                            ExtensionAttribute12 = Get-TrimmedValue -Value $contact.extensionAttribute12
                            ExtensionAttribute13 = Get-TrimmedValue -Value $contact.extensionAttribute13
                            ExtensionAttribute14 = Get-TrimmedValue -Value $contact.extensionAttribute14
                            ExtensionAttribute15 = Get-TrimmedValue -Value $contact.extensionAttribute15
                            CanonicalName = Get-TrimmedValue -Value $contact.CanonicalName
                            WhenCreated = Get-TrimmedValue -Value $contact.whenCreated
                            WhenChanged = Get-TrimmedValue -Value $contact.whenChanged
                            ShowInAddressBook = Convert-MultiValueToString -Value $contact.showInAddressBook
                            LegacyExchangeDn = Get-TrimmedValue -Value $contact.legacyExchangeDN
                            MemberOf = Convert-MultiValueToString -Value $contact.memberOf
                        })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $data = New-EmptyContactData
        $data['IdentityTypeRequested'] = $identityType
        $data['IdentityValueRequested'] = $identityValue
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryContact' -Status 'Failed' -Message $_.Exception.Message -Data $data))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeSearchBase' -NotePropertyValue $resolvedSearchBase -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeServer' -NotePropertyValue $resolvedServer -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeMaxObjects' -NotePropertyValue ($(if ($scopeMode -eq 'DiscoverAll') { [string]$MaxObjects } else { '' })) -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeWasTruncated' -NotePropertyValue ([string]$runWasTruncated) -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory contact inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}


