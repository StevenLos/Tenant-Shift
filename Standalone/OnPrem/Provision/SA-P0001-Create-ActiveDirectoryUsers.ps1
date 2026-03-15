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

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\..\Standalone_OutputCsvPath') -ChildPath ("Results_SA-P0001-Create-ActiveDirectoryUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Add-IfValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Hashtable,

        [Parameter(Mandatory)]
        [string]$Key,

        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        $Hashtable[$Key] = $text
    }
}

function Add-IfNullableBool {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Hashtable,

        [Parameter(Mandatory)]
        [string]$Key,

        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $boolValue = Get-NullableBool -Value $Value
    if ($null -ne $boolValue) {
        $Hashtable[$Key] = $boolValue
    }
}

$requiredHeaders = @(
    'Action',
    'Notes',
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
    'AccountPassword',
    'ChangePasswordAtLogon',
    'PasswordNeverExpires',
    'CannotChangePassword',
    'SmartcardLogonRequired',
    'AccountExpirationDate',
    'UserWorkstations',
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
    'HomePhone',
    'IpPhone',
    'Fax',
    'Pager',
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

Write-Status -Message 'Starting Active Directory user creation script.'
Ensure-ActiveDirectoryConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $samAccountName = Get-TrimmedValue -Value $row.SamAccountName
    $userPrincipalName = Get-TrimmedValue -Value $row.UserPrincipalName
    $primaryKey = if (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) { $userPrincipalName } else { $samAccountName }

    try {
        if ([string]::IsNullOrWhiteSpace($samAccountName)) {
            throw 'SamAccountName is required.'
        }

        if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
            throw 'UserPrincipalName is required.'
        }

        $givenName = Get-TrimmedValue -Value $row.GivenName
        $surname = Get-TrimmedValue -Value $row.Surname
        if ([string]::IsNullOrWhiteSpace($givenName) -or [string]::IsNullOrWhiteSpace($surname)) {
            throw 'GivenName and Surname are required.'
        }

        $path = Get-TrimmedValue -Value $row.Path
        if ([string]::IsNullOrWhiteSpace($path)) {
            throw 'Path (target OU distinguished name) is required.'
        }

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            $displayName = "$givenName $surname".Trim()
        }

        $name = Get-TrimmedValue -Value $row.Name
        if ([string]::IsNullOrWhiteSpace($name)) {
            $name = $displayName
        }

        $enabled = ConvertTo-Bool -Value $row.Enabled -Default $true
        $passwordText = Get-TrimmedValue -Value $row.AccountPassword
        if ($enabled -and [string]::IsNullOrWhiteSpace($passwordText)) {
            throw 'AccountPassword is required when Enabled is TRUE.'
        }

        $existingUser = $null
        $escapedSam = Escape-AdFilterValue -Value $samAccountName
        $existingBySam = Invoke-WithRetry -OperationName "Lookup user by SamAccountName $samAccountName" -ScriptBlock {
            Get-ADUser -Filter "SamAccountName -eq '$escapedSam'" -ErrorAction SilentlyContinue | Select-Object -First 1
        }

        if ($existingBySam) {
            $existingUser = $existingBySam
        }
        else {
            $escapedUpn = Escape-AdFilterValue -Value $userPrincipalName
            $existingByUpn = Invoke-WithRetry -OperationName "Lookup user by UPN $userPrincipalName" -ScriptBlock {
                Get-ADUser -Filter "UserPrincipalName -eq '$escapedUpn'" -ErrorAction SilentlyContinue | Select-Object -First 1
            }

            if ($existingByUpn) {
                $existingUser = $existingByUpn
            }
        }

        if ($existingUser) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryUser' -Status 'Skipped' -Message "User already exists as '$($existingUser.DistinguishedName)'."))
            $rowNumber++
            continue
        }

        $newUserParams = @{
            Name              = $name
            SamAccountName    = $samAccountName
            UserPrincipalName = $userPrincipalName
            GivenName         = $givenName
            Surname           = $surname
            DisplayName       = $displayName
            Enabled           = $enabled
            Path              = $path
        }

        Add-IfValue -Hashtable $newUserParams -Key 'Initials' -Value $row.Initials
        Add-IfValue -Hashtable $newUserParams -Key 'Description' -Value $row.Description
        Add-IfValue -Hashtable $newUserParams -Key 'EmployeeID' -Value $row.EmployeeID
        Add-IfValue -Hashtable $newUserParams -Key 'EmployeeNumber' -Value $row.EmployeeNumber
        Add-IfNullableBool -Hashtable $newUserParams -Key 'ChangePasswordAtLogon' -Value $row.ChangePasswordAtLogon
        Add-IfNullableBool -Hashtable $newUserParams -Key 'PasswordNeverExpires' -Value $row.PasswordNeverExpires
        Add-IfNullableBool -Hashtable $newUserParams -Key 'CannotChangePassword' -Value $row.CannotChangePassword
        Add-IfNullableBool -Hashtable $newUserParams -Key 'SmartcardLogonRequired' -Value $row.SmartcardLogonRequired

        $accountExpirationDate = ConvertTo-NullableDateTime -Value $row.AccountExpirationDate
        if ($null -ne $accountExpirationDate) {
            $newUserParams['AccountExpirationDate'] = $accountExpirationDate
        }

        Add-IfValue -Hashtable $newUserParams -Key 'LogonWorkstations' -Value $row.UserWorkstations
        Add-IfValue -Hashtable $newUserParams -Key 'ScriptPath' -Value $row.ScriptPath
        Add-IfValue -Hashtable $newUserParams -Key 'ProfilePath' -Value $row.ProfilePath
        Add-IfValue -Hashtable $newUserParams -Key 'HomeDirectory' -Value $row.HomeDirectory
        Add-IfValue -Hashtable $newUserParams -Key 'HomeDrive' -Value $row.HomeDrive
        Add-IfValue -Hashtable $newUserParams -Key 'Title' -Value $row.Title
        Add-IfValue -Hashtable $newUserParams -Key 'Department' -Value $row.Department
        Add-IfValue -Hashtable $newUserParams -Key 'Company' -Value $row.Company
        Add-IfValue -Hashtable $newUserParams -Key 'Division' -Value $row.Division
        Add-IfValue -Hashtable $newUserParams -Key 'Office' -Value $row.Office
        Add-IfValue -Hashtable $newUserParams -Key 'Manager' -Value $row.Manager
        Add-IfValue -Hashtable $newUserParams -Key 'OfficePhone' -Value $row.OfficePhone
        Add-IfValue -Hashtable $newUserParams -Key 'MobilePhone' -Value $row.MobilePhone
        Add-IfValue -Hashtable $newUserParams -Key 'HomePhone' -Value $row.HomePhone
        Add-IfValue -Hashtable $newUserParams -Key 'EmailAddress' -Value $row.Mail
        Add-IfValue -Hashtable $newUserParams -Key 'StreetAddress' -Value $row.StreetAddress
        Add-IfValue -Hashtable $newUserParams -Key 'POBox' -Value $row.PostOfficeBox
        Add-IfValue -Hashtable $newUserParams -Key 'City' -Value $row.City
        Add-IfValue -Hashtable $newUserParams -Key 'State' -Value $row.StateOrProvince
        Add-IfValue -Hashtable $newUserParams -Key 'PostalCode' -Value $row.PostalCode
        Add-IfValue -Hashtable $newUserParams -Key 'Fax' -Value $row.Fax
        Add-IfValue -Hashtable $newUserParams -Key 'Pager' -Value $row.Pager
        Add-IfValue -Hashtable $newUserParams -Key 'HomePage' -Value $row.WebPage

        if (-not [string]::IsNullOrWhiteSpace($passwordText)) {
            $newUserParams['AccountPassword'] = ConvertTo-SecureString -String $passwordText -AsPlainText -Force
        }

        $otherAttributes = @{}

        Add-IfValue -Hashtable $otherAttributes -Key 'employeeType' -Value $row.EmployeeType
        Add-IfValue -Hashtable $otherAttributes -Key 'mailNickname' -Value $row.MailNickname
        Add-IfValue -Hashtable $otherAttributes -Key 'targetAddress' -Value $row.TargetAddress
        Add-IfValue -Hashtable $otherAttributes -Key 'ipPhone' -Value $row.IpPhone

        $proxyAddresses = @(
            ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ProxyAddresses) |
                ForEach-Object { [string]$_ }
        )
        if ($proxyAddresses.Count -gt 0) {
            $otherAttributes['proxyAddresses'] = [string[]]$proxyAddresses
        }

        $hideFromAddressLists = Get-NullableBool -Value $row.HideFromAddressLists
        if ($null -ne $hideFromAddressLists) {
            $otherAttributes['msExchHideFromAddressLists'] = $hideFromAddressLists
        }

        Add-IfValue -Hashtable $otherAttributes -Key 'c' -Value $row.CountryCode
        Add-IfValue -Hashtable $otherAttributes -Key 'co' -Value $row.CountryName
        $countryNumericCode = Get-TrimmedValue -Value $row.CountryNumericCode
        if (-not [string]::IsNullOrWhiteSpace($countryNumericCode)) {
            try {
                $otherAttributes['countryCode'] = [int]$countryNumericCode
            }
            catch {
                throw "CountryNumericCode '$countryNumericCode' must be an integer value."
            }
        }

        for ($i = 1; $i -le 15; $i++) {
            $columnName = "ExtensionAttribute$i"
            $attributeName = "extensionAttribute$i"
            Add-IfValue -Hashtable $otherAttributes -Key $attributeName -Value $row.$columnName
        }

        if ($otherAttributes.Count -gt 0) {
            $newUserParams['OtherAttributes'] = $otherAttributes
        }

        if ($PSCmdlet.ShouldProcess($primaryKey, 'Create Active Directory user')) {
            Invoke-WithRetry -OperationName "Create AD user $primaryKey" -ScriptBlock {
                New-ADUser @newUserParams -ErrorAction Stop
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryUser' -Status 'Created' -Message 'User created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryUser' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'CreateActiveDirectoryUser' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory user creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
