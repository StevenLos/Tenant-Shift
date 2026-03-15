<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-120000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\\..\\Standalone_OutputCsvPath') -ChildPath ("Results_SA-M3001-Update-EntraUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
        return [string[]]@()
    }

    $items = [System.Collections.Generic.List[string]]::new()
    foreach ($rawPart in ($Value -split [Regex]::Escape($Delimiter))) {
        $part = ([string]$rawPart).Trim()
        if (-not [string]::IsNullOrWhiteSpace($part)) {
            [void]$items.Add($part)
        }
    }

    return [string[]]$items.ToArray()
}

function Escape-ODataString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    return $Value.Replace("'", "''")
}

function Assert-ModuleCurrent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ModuleNames
    )

    foreach ($moduleName in $ModuleNames) {
        Write-Status -Message "Checking module '$moduleName'."

        $installed = Get-Module -ListAvailable -Name $moduleName |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if (-not $installed) {
            throw "Required module '$moduleName' is not installed. Install with: Install-Module $moduleName -Scope CurrentUser"
        }

        Write-Status -Message "Installed version for '$moduleName': $($installed.Version)."

        try {
            $gallery = Find-Module -Name $moduleName -Repository PSGallery -ErrorAction Stop
        }
        catch {
            throw "Unable to verify the latest version for '$moduleName' from PSGallery. Ensure internet access and try again. Error: $($_.Exception.Message)"
        }

        if ($installed.Version -lt $gallery.Version) {
            throw "Module '$moduleName' is outdated. Installed: $($installed.Version), current: $($gallery.Version). Update with: Update-Module $moduleName"
        }

        Write-Status -Message "Module '$moduleName' is current ($($installed.Version))." -Level SUCCESS
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

function Ensure-GraphConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$RequiredScopes
    )

    $context = Get-MgContext -ErrorAction SilentlyContinue
    $needsConnect = $true

    if ($context) {
        $missingScopes = @(
            $RequiredScopes | Where-Object { $_ -notin $context.Scopes }
        )

        if ($missingScopes.Count -eq 0) {
            Write-Status -Message "Already connected to Microsoft Graph as '$($context.Account)'." -Level SUCCESS
            $needsConnect = $false
        }
        else {
            Write-Status -Message "Graph connection exists but is missing scopes: $($missingScopes -join ', '). Reconnecting." -Level WARN
        }
    }
    else {
        Write-Status -Message 'No active Microsoft Graph connection detected. Connecting now.' -Level WARN
    }

    if ($needsConnect) {
        Connect-MgGraph -Scopes $RequiredScopes -NoWelcome -ErrorAction Stop | Out-Null

        $context = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $context) {
            throw 'Microsoft Graph connection failed. No active context was returned.'
        }

        Write-Status -Message "Connected to Microsoft Graph tenant '$($context.TenantId)' as '$($context.Account)'." -Level SUCCESS
    }
}

function Test-ExchangeConnection {
    [CmdletBinding()]
    param()

    if (Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        try {
            $connection = Get-ConnectionInformation -ErrorAction Stop |
                Where-Object { $_.State -eq 'Connected' } |
                Select-Object -First 1

            if ($connection) {
                return $true
            }
        }
        catch {
            # Continue to fallback probe.
        }
    }

    try {
        Get-EXORecipient -ResultSize 1 -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Ensure-ExchangeConnection {
    [CmdletBinding()]
    param()

    if (Test-ExchangeConnection) {
        Write-Status -Message 'Already connected to Exchange Online.' -Level SUCCESS
        return
    }

    Write-Status -Message 'No active Exchange Online connection detected. Connecting now.' -Level WARN

    $connectCommand = Get-Command -Name Connect-ExchangeOnline -ErrorAction Stop
    $supportsDisableWam = $connectCommand.Parameters.ContainsKey('DisableWAM')
    $supportsDevice = $connectCommand.Parameters.ContainsKey('Device')

    $getCombinedExceptionMessage = {
        param(
            [Parameter(Mandatory)]
            [System.Exception]$Exception
        )

        $messageParts = [System.Collections.Generic.List[string]]::new()
        $cursor = $Exception

        while ($null -ne $cursor) {
            if (-not [string]::IsNullOrWhiteSpace($cursor.Message)) {
                $messageParts.Add($cursor.Message)
            }

            $cursor = $cursor.InnerException
        }

        return ($messageParts -join ' ').ToLowerInvariant()
    }

    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
    }
    catch {
        $initialException = $_.Exception
        $combinedMessage = & $getCombinedExceptionMessage -Exception $initialException
        $looksLikeBrokerIssue = $combinedMessage -match 'runtimebroker|acquiring token|nullreferenceexception|object reference not set|broker'

        if ($looksLikeBrokerIssue -and $supportsDisableWam) {
            try {
                Write-Status -Message 'Exchange sign-in failed with broker/WAM error. Retrying with -DisableWAM.' -Level WARN
                Connect-ExchangeOnline -ShowBanner:$false -DisableWAM -ErrorAction Stop | Out-Null
            }
            catch {
                $disableWamException = $_.Exception

                if ($supportsDevice) {
                    Write-Status -Message 'Retry with -DisableWAM failed. Retrying with device code sign-in (-Device).' -Level WARN
                    Connect-ExchangeOnline -ShowBanner:$false -Device -DisableWAM:$supportsDisableWam -ErrorAction Stop | Out-Null
                }
                else {
                    throw "Exchange sign-in failed with broker/WAM error and -DisableWAM retry also failed. Original error: $($initialException.Message) Secondary error: $($disableWamException.Message)"
                }
            }
        }
        elseif ($looksLikeBrokerIssue -and -not $supportsDisableWam) {
            if ($supportsDevice) {
                Write-Status -Message 'Exchange sign-in failed with broker/WAM error. Retrying with device code sign-in (-Device).' -Level WARN
                Connect-ExchangeOnline -ShowBanner:$false -Device -ErrorAction Stop | Out-Null
            }
            else {
                throw "Exchange sign-in failed with broker/WAM error, and this ExchangeOnlineManagement version does not support -DisableWAM or -Device. Update the module and retry. Original error: $($initialException.Message)"
            }
        }
        else {
            throw
        }
    }

    if (-not (Test-ExchangeConnection)) {
        throw 'Exchange Online connection failed. Unable to verify an active session.'
    }

    Write-Status -Message 'Connected to Exchange Online.' -Level SUCCESS
}

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
    return ($combinedMessage -match 'too many request|throttl|temporar|timeout|timed out|service unavailable|rate limit|try again|gateway|429|500|502|503|504')
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

    $retryAfterValue = $null
    foreach ($propertyName in @('ResponseHeaders', 'Headers')) {
        if ($Exception.PSObject.Properties.Name -contains $propertyName) {
            $headers = $Exception.$propertyName
            if ($headers) {
                if ($headers.PSObject.Properties.Name -contains 'RetryAfter') {
                    $retryAfterValue = $headers.RetryAfter
                    break
                }

                if ($headers.PSObject.Properties.Name -contains 'Retry-After') {
                    $retryAfterValue = $headers.'Retry-After'
                    break
                }

                try {
                    if ($headers.ContainsKey('Retry-After')) {
                        $retryAfterValue = $headers['Retry-After']
                        break
                    }
                }
                catch {
                    # Best effort.
                }
            }
        }
    }

    if ($null -ne $retryAfterValue) {
        $retryAfterString = [string]$retryAfterValue
        if ($retryAfterString -match '^\d+$') {
            return [Math]::Min([int]$retryAfterString, $MaxDelaySeconds)
        }
    }

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

function Get-TrimmedValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    return ([string]$Value).Trim()
}

function Convert-ToIsoDateTimeOffsetString {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value,

        [Parameter(Mandatory)]
        [string]$FieldName
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    try {
        $parsed = [datetimeoffset]::Parse($text)
        return $parsed.ToString('o')
    }
    catch {
        throw "$FieldName value '$text' is invalid. Use an ISA-8601 compatible date/time value."
    }
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'UserPrincipalName',
    'DisplayName',
    'GivenName',
    'Surname',
    'MailNickname',
    'UserType',
    'Password',
    'ForceChangePasswordNextSignIn',
    'ForceChangePasswordNextSignInWithMfa',
    'AccountEnabled',
    'UsageLocation',
    'PreferredLanguage',
    'Department',
    'JobTitle',
    'CompanyName',
    'OfficeLocation',
    'EmployeeId',
    'EmployeeType',
    'EmployeeHireDate',
    'MobilePhone',
    'BusinessPhones',
    'FaxNumber',
    'OtherMails',
    'StreetAddress',
    'City',
    'State',
    'PostalCode',
    'Country',
    'PasswordPolicies',
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
    'ClearAttributes'
)

$clearFieldMap = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$clearFieldMap['GivenName'] = 'givenName'
$clearFieldMap['Surname'] = 'surname'
$clearFieldMap['UsageLocation'] = 'usageLocation'
$clearFieldMap['PreferredLanguage'] = 'preferredLanguage'
$clearFieldMap['Department'] = 'department'
$clearFieldMap['JobTitle'] = 'jobTitle'
$clearFieldMap['CompanyName'] = 'companyName'
$clearFieldMap['OfficeLocation'] = 'officeLocation'
$clearFieldMap['EmployeeId'] = 'employeeId'
$clearFieldMap['EmployeeType'] = 'employeeType'
$clearFieldMap['EmployeeHireDate'] = 'employeeHireDate'
$clearFieldMap['MobilePhone'] = 'mobilePhone'
$clearFieldMap['BusinessPhones'] = 'businessPhones'
$clearFieldMap['FaxNumber'] = 'faxNumber'
$clearFieldMap['OtherMails'] = 'otherMails'
$clearFieldMap['StreetAddress'] = 'streetAddress'
$clearFieldMap['City'] = 'city'
$clearFieldMap['State'] = 'state'
$clearFieldMap['PostalCode'] = 'postalCode'
$clearFieldMap['Country'] = 'country'
$clearFieldMap['PasswordPolicies'] = 'passwordPolicies'

Write-Status -Message 'Starting Entra ID user update script (expanded field model).'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('User.ReadWrite.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $upn = Get-TrimmedValue -Value $row.UserPrincipalName

    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw 'UserPrincipalName is required.'
        }

        $escapedUpn = Escape-ODataString -Value $upn
        $users = @(Invoke-WithRetry -OperationName "Lookup user $upn" -ScriptBlock {
            Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -Property 'id,userPrincipalName,displayName,accountEnabled' -ErrorAction Stop
        })

        if ($users.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'UpdateUser' -Status 'NotFound' -Message 'User not found.'))
            $rowNumber++
            continue
        }

        if ($users.Count -gt 1) {
            throw "Multiple users were returned for UPN '$upn'. Resolve duplicate directory objects before retrying."
        }

        $user = $users[0]
        $body = @{}
        $setFields = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        $scalarMappings = @(
            @{ Column = 'DisplayName'; Property = 'displayName' },
            @{ Column = 'GivenName'; Property = 'givenName' },
            @{ Column = 'Surname'; Property = 'surname' },
            @{ Column = 'MailNickname'; Property = 'mailNickname' },
            @{ Column = 'UserType'; Property = 'userType' },
            @{ Column = 'UsageLocation'; Property = 'usageLocation' },
            @{ Column = 'PreferredLanguage'; Property = 'preferredLanguage' },
            @{ Column = 'Department'; Property = 'department' },
            @{ Column = 'JobTitle'; Property = 'jobTitle' },
            @{ Column = 'CompanyName'; Property = 'companyName' },
            @{ Column = 'OfficeLocation'; Property = 'officeLocation' },
            @{ Column = 'EmployeeId'; Property = 'employeeId' },
            @{ Column = 'EmployeeType'; Property = 'employeeType' },
            @{ Column = 'MobilePhone'; Property = 'mobilePhone' },
            @{ Column = 'FaxNumber'; Property = 'faxNumber' },
            @{ Column = 'StreetAddress'; Property = 'streetAddress' },
            @{ Column = 'City'; Property = 'city' },
            @{ Column = 'State'; Property = 'state' },
            @{ Column = 'PostalCode'; Property = 'postalCode' },
            @{ Column = 'Country'; Property = 'country' },
            @{ Column = 'PasswordPolicies'; Property = 'passwordPolicies' }
        )

        foreach ($mapping in $scalarMappings) {
            $value = Get-TrimmedValue -Value $row.($mapping.Column)
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                if ($mapping.Property -eq 'userType' -and $value -notin @('Member', 'Guest')) {
                    throw "UserType '$value' is invalid. Use Member or Guest."
                }

                $body[$mapping.Property] = $value
                $null = $setFields.Add($mapping.Column)
            }
        }

        $accountEnabledRaw = Get-TrimmedValue -Value $row.AccountEnabled
        if (-not [string]::IsNullOrWhiteSpace($accountEnabledRaw)) {
            $body['accountEnabled'] = ConvertTo-Bool -Value $accountEnabledRaw
            $null = $setFields.Add('AccountEnabled')
        }

        $businessPhonesRaw = Get-TrimmedValue -Value $row.BusinessPhones
        if (-not [string]::IsNullOrWhiteSpace($businessPhonesRaw)) {
            $body['businessPhones'] = ConvertTo-Array -Value $businessPhonesRaw
            $null = $setFields.Add('BusinessPhones')
        }

        $otherMailsRaw = Get-TrimmedValue -Value $row.OtherMails
        if (-not [string]::IsNullOrWhiteSpace($otherMailsRaw)) {
            $body['otherMails'] = ConvertTo-Array -Value $otherMailsRaw
            $null = $setFields.Add('OtherMails')
        }

        $employeeHireDate = Convert-ToIsoDateTimeOffsetString -Value $row.EmployeeHireDate -FieldName 'EmployeeHireDate'
        if (-not [string]::IsNullOrWhiteSpace($employeeHireDate)) {
            $body['employeeHireDate'] = $employeeHireDate
            $null = $setFields.Add('EmployeeHireDate')
        }

        $password = [string]$row.Password
        if (-not [string]::IsNullOrWhiteSpace($password)) {
            $passwordProfile = @{ password = $password }

            $forceChange = Get-TrimmedValue -Value $row.ForceChangePasswordNextSignIn
            if (-not [string]::IsNullOrWhiteSpace($forceChange)) {
                $passwordProfile['forceChangePasswordNextSignIn'] = ConvertTo-Bool -Value $forceChange
            }

            $forceChangeWithMfa = Get-TrimmedValue -Value $row.ForceChangePasswordNextSignInWithMfa
            if (-not [string]::IsNullOrWhiteSpace($forceChangeWithMfa)) {
                $passwordProfile['forceChangePasswordNextSignInWithMfa'] = ConvertTo-Bool -Value $forceChangeWithMfa
            }

            $body['passwordProfile'] = $passwordProfile
            $null = $setFields.Add('Password')
        }

        $extensionAttributes = @{}
        $extensionTouched = $false
        for ($i = 1; $i -le 15; $i++) {
            $columnName = "ExtensionAttribute$i"
            $attributeName = "extensionAttribute$i"
            $value = Get-TrimmedValue -Value $row.$columnName
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $extensionAttributes[$attributeName] = $value
                $extensionTouched = $true
                $null = $setFields.Add($columnName)
            }
        }

        $clearRequested = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ClearAttributes)
        foreach ($clearName in $clearRequested) {
            if ($setFields.Contains($clearName)) {
                throw "Field '$clearName' is set and cleared in the same row. Use only one behavior per field."
            }

            if ($clearFieldMap.ContainsKey($clearName)) {
                $propertyName = $clearFieldMap[$clearName]
                if ($propertyName -in @('businessPhones', 'otherMails')) {
                    $body[$propertyName] = @()
                }
                else {
                    $body[$propertyName] = $null
                }

                continue
            }

            if ($clearName -match '^ExtensionAttribute([1-9]|1[0-5])$') {
                $attributeName = 'extensionAttribute{0}' -f $Matches[1]
                $extensionAttributes[$attributeName] = $null
                $extensionTouched = $true
                continue
            }

            throw "ClearAttributes value '$clearName' is not supported."
        }

        if ($extensionTouched) {
            $body['onPremisesExtensionAttributes'] = $extensionAttributes
        }

        if ($body.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'UpdateUser' -Status 'Skipped' -Message 'No updates were requested.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($upn, 'Update Entra ID user attributes')) {
            Invoke-WithRetry -OperationName "Update user $upn" -ScriptBlock {
                Update-MgUser -UserId $user.Id -BodyParameter $body -ErrorAction Stop | Out-Null
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'UpdateUser' -Status 'Updated' -Message 'User updated successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'UpdateUser' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'UpdateUser' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID user update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}

