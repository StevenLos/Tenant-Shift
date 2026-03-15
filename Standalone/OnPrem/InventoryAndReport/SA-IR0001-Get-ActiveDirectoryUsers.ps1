<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260305-130500

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\..\Standalone_OutputCsvPath') -ChildPath ("Results_SA-IR0001-Get-ActiveDirectoryUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Resolve-UsersByScope {
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
                Filter      = '*'
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADUser @params)
        }
        'samaccountname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            $params = @{
                Filter      = "SamAccountName -eq '$escaped'"
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADUser @params)
        }
        'userprincipalname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            $params = @{
                Filter      = "UserPrincipalName -eq '$escaped'"
                Properties  = '*'
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
                $params['SearchBase'] = $SearchBase
            }
            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            return @(Get-ADUser @params)
        }
        'distinguishedname' {
            $params = @{
                Identity    = $IdentityValue
                Properties  = '*'
                ErrorAction = 'SilentlyContinue'
            }

            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            $user = Get-ADUser @params
            if ($user) {
                return @($user)
            }

            return @()
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            $params = @{
                Identity    = $guid
                Properties  = '*'
                ErrorAction = 'SilentlyContinue'
            }

            if (-not [string]::IsNullOrWhiteSpace($Server)) {
                $params['Server'] = $Server
            }

            $user = Get-ADUser @params
            if ($user) {
                return @($user)
            }

            return @()
        }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use All, SamAccountName, UserPrincipalName, DistinguishedName, or ObjectGuid."
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
    'SID',
    'SamAccountName',
    'UserPrincipalName',
    'GivenName',
    'Initials',
    'Surname',
    'DisplayName',
    'Name',
    'Description',
    'EmployeeID',
    'EmployeeNumber',
    'EmployeeType',
    'Enabled',
    'ChangePasswordAtLogon',
    'PasswordNeverExpires',
    'CannotChangePassword',
    'SmartcardLogonRequired',
    'AccountExpirationDate',
    'LogonWorkstations',
    'ScriptPath',
    'ProfilePath',
    'HomeDirectory',
    'HomeDrive',
    'Title',
    'Department',
    'Company',
    'Division',
    'Office',
    'Manager',
    'OfficePhone',
    'MobilePhone',
    'Mail',
    'mailNickname',
    'proxyAddresses',
    'targetAddress',
    'msExchHideFromAddressLists',
    'StreetAddress',
    'POBox',
    'City',
    'State',
    'PostalCode',
    'c',
    'co',
    'countryCode',
    'HomePhone',
    'ipPhone',
    'Fax',
    'Pager',
    'HomePage',
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
    'LastLogonDate',
    'LastBadPasswordAttempt',
    'PasswordLastSet',
    'LockedOut',
    'MemberOf'
)

Write-Status -Message 'Starting Active Directory user inventory script.'
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

    Write-Status -Message "DiscoverAll enabled for Active Directory users. SearchBase='$resolvedSearchBase'." -Level WARN
    $rows = @([PSCustomObject]@{
            IdentityType  = 'All'
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
        $users = Invoke-WithRetry -OperationName "Load users for $primaryKey" -ScriptBlock {
            Resolve-UsersByScope -IdentityType $identityType -IdentityValue $identityValue -PropertyNames $propertyNames -SearchBase $effectiveSearchBase -Server $resolvedServer
        }

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $users.Count -gt $MaxObjects) {
            $users = @($users | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($users.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryUser' -Status 'NotFound' -Message 'No matching users were found.' -Data ([ordered]@{
                            IdentityTypeRequested = $identityType
                            IdentityValueRequested = $identityValue
                            DistinguishedName = ''
                            ObjectGuid = ''
                            SID = ''
                            SamAccountName = ''
                            UserPrincipalName = ''
                            GivenName = ''
                            Initials = ''
                            Surname = ''
                            DisplayName = ''
                            Name = ''
                            Description = ''
                            EmployeeID = ''
                            EmployeeNumber = ''
                            EmployeeType = ''
                            Enabled = ''
                            ChangePasswordAtLogon = ''
                            PasswordNeverExpires = ''
                            CannotChangePassword = ''
                            SmartcardLogonRequired = ''
                            AccountExpirationDate = ''
                            UserWorkstations = ''
                            ScriptPath = ''
                            ProfilePath = ''
                            HomeDirectory = ''
                            HomeDrive = ''
                            Title = ''
                            Department = ''
                            Company = ''
                            Division = ''
                            Office = ''
                            Manager = ''
                            OfficePhone = ''
                            MobilePhone = ''
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
                            HomePhone = ''
                            IpPhone = ''
                            Fax = ''
                            Pager = ''
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
                            LastLogonDate = ''
                            LastBadPasswordAttempt = ''
                            PasswordLastSet = ''
                            LockedOut = ''
                            MemberOf = ''
                        })))

            $rowNumber++
            continue
        }

        foreach ($user in @($users | Sort-Object -Property UserPrincipalName, SamAccountName, DistinguishedName)) {
            $userPrimaryKey = if (-not [string]::IsNullOrWhiteSpace([string]$user.UserPrincipalName)) { ([string]$user.UserPrincipalName).Trim() } else { ([string]$user.SamAccountName).Trim() }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $userPrimaryKey -Action 'GetActiveDirectoryUser' -Status 'Completed' -Message 'User exported.' -Data ([ordered]@{
                            IdentityTypeRequested = $identityType
                            IdentityValueRequested = $identityValue
                            DistinguishedName = Get-TrimmedValue -Value $user.DistinguishedName
                            ObjectGuid = Get-TrimmedValue -Value $user.ObjectGuid
                            SID = Get-TrimmedValue -Value $user.SID
                            SamAccountName = Get-TrimmedValue -Value $user.SamAccountName
                            UserPrincipalName = Get-TrimmedValue -Value $user.UserPrincipalName
                            GivenName = Get-TrimmedValue -Value $user.GivenName
                            Initials = Get-TrimmedValue -Value $user.Initials
                            Surname = Get-TrimmedValue -Value $user.Surname
                            DisplayName = Get-TrimmedValue -Value $user.DisplayName
                            Name = Get-TrimmedValue -Value $user.Name
                            Description = Get-TrimmedValue -Value $user.Description
                            EmployeeID = Get-TrimmedValue -Value $user.EmployeeID
                            EmployeeNumber = Get-TrimmedValue -Value $user.EmployeeNumber
                            EmployeeType = Get-TrimmedValue -Value $user.EmployeeType
                            Enabled = [string]$user.Enabled
                            ChangePasswordAtLogon = [string]$user.ChangePasswordAtLogon
                            PasswordNeverExpires = [string]$user.PasswordNeverExpires
                            CannotChangePassword = [string]$user.CannotChangePassword
                            SmartcardLogonRequired = [string]$user.SmartcardLogonRequired
                            AccountExpirationDate = Get-TrimmedValue -Value $user.AccountExpirationDate
                            UserWorkstations = Get-TrimmedValue -Value $user.LogonWorkstations
                            ScriptPath = Get-TrimmedValue -Value $user.ScriptPath
                            ProfilePath = Get-TrimmedValue -Value $user.ProfilePath
                            HomeDirectory = Get-TrimmedValue -Value $user.HomeDirectory
                            HomeDrive = Get-TrimmedValue -Value $user.HomeDrive
                            Title = Get-TrimmedValue -Value $user.Title
                            Department = Get-TrimmedValue -Value $user.Department
                            Company = Get-TrimmedValue -Value $user.Company
                            Division = Get-TrimmedValue -Value $user.Division
                            Office = Get-TrimmedValue -Value $user.Office
                            Manager = Get-TrimmedValue -Value $user.Manager
                            OfficePhone = Get-TrimmedValue -Value $user.OfficePhone
                            MobilePhone = Get-TrimmedValue -Value $user.MobilePhone
                            Mail = Get-TrimmedValue -Value $user.Mail
                            MailNickname = Get-TrimmedValue -Value $user.mailNickname
                            ProxyAddresses = Convert-MultiValueToString -Value $user.proxyAddresses
                            TargetAddress = Get-TrimmedValue -Value $user.targetAddress
                            HideFromAddressLists = [string]$user.msExchHideFromAddressLists
                            StreetAddress = Get-TrimmedValue -Value $user.StreetAddress
                            PostOfficeBox = Get-TrimmedValue -Value $user.POBox
                            City = Get-TrimmedValue -Value $user.City
                            StateOrProvince = Get-TrimmedValue -Value $user.State
                            PostalCode = Get-TrimmedValue -Value $user.PostalCode
                            CountryCode = Get-TrimmedValue -Value $user.c
                            CountryName = Get-TrimmedValue -Value $user.co
                            CountryNumericCode = Get-TrimmedValue -Value $user.countryCode
                            HomePhone = Get-TrimmedValue -Value $user.HomePhone
                            IpPhone = Get-TrimmedValue -Value $user.ipPhone
                            Fax = Get-TrimmedValue -Value $user.Fax
                            Pager = Get-TrimmedValue -Value $user.Pager
                            WebPage = Get-TrimmedValue -Value $user.HomePage
                            ExtensionAttribute1 = Get-TrimmedValue -Value $user.extensionAttribute1
                            ExtensionAttribute2 = Get-TrimmedValue -Value $user.extensionAttribute2
                            ExtensionAttribute3 = Get-TrimmedValue -Value $user.extensionAttribute3
                            ExtensionAttribute4 = Get-TrimmedValue -Value $user.extensionAttribute4
                            ExtensionAttribute5 = Get-TrimmedValue -Value $user.extensionAttribute5
                            ExtensionAttribute6 = Get-TrimmedValue -Value $user.extensionAttribute6
                            ExtensionAttribute7 = Get-TrimmedValue -Value $user.extensionAttribute7
                            ExtensionAttribute8 = Get-TrimmedValue -Value $user.extensionAttribute8
                            ExtensionAttribute9 = Get-TrimmedValue -Value $user.extensionAttribute9
                            ExtensionAttribute10 = Get-TrimmedValue -Value $user.extensionAttribute10
                            ExtensionAttribute11 = Get-TrimmedValue -Value $user.extensionAttribute11
                            ExtensionAttribute12 = Get-TrimmedValue -Value $user.extensionAttribute12
                            ExtensionAttribute13 = Get-TrimmedValue -Value $user.extensionAttribute13
                            ExtensionAttribute14 = Get-TrimmedValue -Value $user.extensionAttribute14
                            ExtensionAttribute15 = Get-TrimmedValue -Value $user.extensionAttribute15
                            CanonicalName = Get-TrimmedValue -Value $user.CanonicalName
                            WhenCreated = Get-TrimmedValue -Value $user.whenCreated
                            WhenChanged = Get-TrimmedValue -Value $user.whenChanged
                            LastLogonDate = Get-TrimmedValue -Value $user.LastLogonDate
                            LastBadPasswordAttempt = Get-TrimmedValue -Value $user.LastBadPasswordAttempt
                            PasswordLastSet = Get-TrimmedValue -Value $user.PasswordLastSet
                            LockedOut = [string]$user.LockedOut
                            MemberOf = Convert-MultiValueToString -Value $user.MemberOf
                        })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetActiveDirectoryUser' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                        IdentityTypeRequested = $identityType
                        IdentityValueRequested = $identityValue
                        DistinguishedName = ''
                        ObjectGuid = ''
                        SID = ''
                        SamAccountName = ''
                        UserPrincipalName = ''
                        GivenName = ''
                        Initials = ''
                        Surname = ''
                        DisplayName = ''
                        Name = ''
                        Description = ''
                        EmployeeID = ''
                        EmployeeNumber = ''
                        EmployeeType = ''
                        Enabled = ''
                        ChangePasswordAtLogon = ''
                        PasswordNeverExpires = ''
                        CannotChangePassword = ''
                        SmartcardLogonRequired = ''
                        AccountExpirationDate = ''
                        UserWorkstations = ''
                        ScriptPath = ''
                        ProfilePath = ''
                        HomeDirectory = ''
                        HomeDrive = ''
                        Title = ''
                        Department = ''
                        Company = ''
                        Division = ''
                        Office = ''
                        Manager = ''
                        OfficePhone = ''
                        MobilePhone = ''
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
                        HomePhone = ''
                        IpPhone = ''
                        Fax = ''
                        Pager = ''
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
                        LastLogonDate = ''
                        LastBadPasswordAttempt = ''
                        PasswordLastSet = ''
                        LockedOut = ''
                        MemberOf = ''
                    })))
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
Write-Status -Message 'Active Directory user inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
