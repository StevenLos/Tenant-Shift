<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-141500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Groups

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\\..\\Standalone_OutputCsvPath') -ChildPath ("Results_SA-M3009-Set-EntraGroupCreatorsPolicy_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-GraphPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$Object,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    if ($null -eq $Object) {
        return $null
    }

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($PropertyName)) {
            return $Object[$PropertyName]
        }
    }

    if ($Object.PSObject.Properties.Name -contains $PropertyName) {
        return $Object.$PropertyName
    }

    if ($Object.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $additional = $Object.AdditionalProperties
        if ($additional -is [System.Collections.IDictionary] -and $additional.Contains($PropertyName)) {
            return $additional[$PropertyName]
        }
    }

    return $null
}

function Convert-SettingValuesToMap {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$Values
    )

    $map = @{}

    foreach ($entry in @($Values)) {
        $name = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $entry -PropertyName 'name')
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $map[$name] = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $entry -PropertyName 'value')
    }

    return $map
}

function Convert-SettingMapToValues {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Map
    )

    $values = [System.Collections.Generic.List[object]]::new()

    foreach ($key in @($Map.Keys | Sort-Object)) {
        $values.Add([ordered]@{
                name  = [string]$key
                value = [string]$Map[$key]
            })
    }

    return $values.ToArray()
}

function Get-GroupUnifiedSetting {
    [CmdletBinding()]
    param()

    $response = Invoke-WithRetry -OperationName 'Load Group.Unified setting' -ScriptBlock {
        Invoke-MgGraphRequest -Method GET -Uri "/v1.0/groupSettings?`$filter=displayName eq 'Group.Unified'" -OutputType PSObject -ErrorAction Stop
    }

    $settings = @((Get-GraphPropertyValue -Object $response -PropertyName 'value'))
    if ($settings.Count -eq 0) {
        return $null
    }

    if ($settings.Count -gt 1) {
        Write-Status -Message 'Multiple Group.Unified settings found. The first setting object will be used.' -Level WARN
    }

    return $settings[0]
}

function Get-GroupUnifiedTemplate {
    [CmdletBinding()]
    param()

    $response = Invoke-WithRetry -OperationName 'Load Group.Unified template' -ScriptBlock {
        Invoke-MgGraphRequest -Method GET -Uri "/v1.0/groupSettingTemplates?`$filter=displayName eq 'Group.Unified'" -OutputType PSObject -ErrorAction Stop
    }

    $templates = @((Get-GraphPropertyValue -Object $response -PropertyName 'value'))
    if ($templates.Count -eq 0) {
        throw "Unable to locate the 'Group.Unified' setting template in Microsoft Graph."
    }

    return $templates[0]
}

function Resolve-AllowedGroup {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$IdentityType,

        [AllowEmptyString()]
        [string]$IdentityValue
    )

    $resolvedType = Get-TrimmedValue -Value $IdentityType
    $resolvedValue = Get-TrimmedValue -Value $IdentityValue

    if ([string]::IsNullOrWhiteSpace($resolvedType) -and [string]::IsNullOrWhiteSpace($resolvedValue)) {
        return [PSCustomObject]@{
            Id           = ''
            DisplayName  = ''
            MailNickname = ''
        }
    }

    if ([string]::IsNullOrWhiteSpace($resolvedType) -or [string]::IsNullOrWhiteSpace($resolvedValue)) {
        throw 'AllowedGroupIdentityType and AllowedGroupIdentityValue must both be set when resolving an allowed group.'
    }

    switch ($resolvedType.Trim().ToLowerInvariant()) {
        'groupid' {
            $group = Invoke-WithRetry -OperationName "Lookup allowed group by id $resolvedValue" -ScriptBlock {
                Get-MgGroup -GroupId $resolvedValue -Property 'id,displayName,mailNickname' -ErrorAction SilentlyContinue
            }

            if (-not $group) {
                throw "Allowed group '$resolvedValue' was not found."
            }

            return [PSCustomObject]@{
                Id           = Get-TrimmedValue -Value $group.Id
                DisplayName  = Get-TrimmedValue -Value $group.DisplayName
                MailNickname = Get-TrimmedValue -Value $group.MailNickname
            }
        }
        'displayname' {
            $escaped = Escape-ODataString -Value $resolvedValue
            $groups = @(Invoke-WithRetry -OperationName "Lookup allowed group by displayName $resolvedValue" -ScriptBlock {
                    Get-MgGroup -Filter "displayName eq '$escaped'" -ConsistencyLevel eventual -Property 'id,displayName,mailNickname' -ErrorAction Stop
                })

            if ($groups.Count -eq 0) {
                throw "Allowed group displayName '$resolvedValue' was not found."
            }

            if ($groups.Count -gt 1) {
                throw "Multiple groups matched displayName '$resolvedValue'. Use AllowedGroupIdentityType='GroupId' for a unique match."
            }

            $group = $groups[0]
            return [PSCustomObject]@{
                Id           = Get-TrimmedValue -Value $group.Id
                DisplayName  = Get-TrimmedValue -Value $group.DisplayName
                MailNickname = Get-TrimmedValue -Value $group.MailNickname
            }
        }
        'mailnickname' {
            $escaped = Escape-ODataString -Value $resolvedValue
            $groups = @(Invoke-WithRetry -OperationName "Lookup allowed group by mailNickname $resolvedValue" -ScriptBlock {
                    Get-MgGroup -Filter "mailNickname eq '$escaped'" -ConsistencyLevel eventual -Property 'id,displayName,mailNickname' -ErrorAction Stop
                })

            if ($groups.Count -eq 0) {
                throw "Allowed group mailNickname '$resolvedValue' was not found."
            }

            if ($groups.Count -gt 1) {
                throw "Multiple groups matched mailNickname '$resolvedValue'. Use AllowedGroupIdentityType='GroupId' for a unique match."
            }

            $group = $groups[0]
            return [PSCustomObject]@{
                Id           = Get-TrimmedValue -Value $group.Id
                DisplayName  = Get-TrimmedValue -Value $group.DisplayName
                MailNickname = Get-TrimmedValue -Value $group.MailNickname
            }
        }
        default {
            throw "AllowedGroupIdentityType '$resolvedType' is invalid. Use GroupId, DisplayName, or MailNickname."
        }
    }
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'AllowGroupCreation',
    'AllowedGroupIdentityType',
    'AllowedGroupIdentityValue',
    'ClearAllowedGroup'
)

Write-Status -Message 'Starting Entra group creators policy update script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')
Ensure-GraphConnection -RequiredScopes @('Directory.ReadWrite.All', 'Group.ReadWrite.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $policyKey = 'Group.Unified'

    try {
        $allowGroupCreation = Get-NullableBool -Value $row.AllowGroupCreation
        $allowedGroupIdentityType = Get-TrimmedValue -Value $row.AllowedGroupIdentityType
        $allowedGroupIdentityValue = Get-TrimmedValue -Value $row.AllowedGroupIdentityValue
        $clearAllowedGroup = Get-NullableBool -Value $row.ClearAllowedGroup

        if ($null -eq $clearAllowedGroup) {
            $clearAllowedGroup = $false
        }

        if ($clearAllowedGroup -and (-not [string]::IsNullOrWhiteSpace($allowedGroupIdentityType) -or -not [string]::IsNullOrWhiteSpace($allowedGroupIdentityValue))) {
            throw 'ClearAllowedGroup cannot be TRUE when an allowed group identity is also supplied.'
        }

        $resolvedAllowedGroup = Resolve-AllowedGroup -IdentityType $allowedGroupIdentityType -IdentityValue $allowedGroupIdentityValue
        $hasRequestedChange = ($null -ne $allowGroupCreation) -or $clearAllowedGroup -or (-not [string]::IsNullOrWhiteSpace($resolvedAllowedGroup.Id))

        if (-not $hasRequestedChange) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $policyKey -Action 'SetEntraGroupCreatorsPolicy' -Status 'Skipped' -Message 'No policy updates were requested.'))
            $rowNumber++
            continue
        }

        $setting = Get-GroupUnifiedSetting
        $settingId = ''
        $templateId = ''
        $valuesMap = @{}
        $settingWasCreated = $false

        if ($setting) {
            $settingId = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $setting -PropertyName 'id')
            $valuesMap = Convert-SettingValuesToMap -Values @((Get-GraphPropertyValue -Object $setting -PropertyName 'values'))
        }
        else {
            $template = Get-GroupUnifiedTemplate
            $templateId = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $template -PropertyName 'id')
            $valuesMap = Convert-SettingValuesToMap -Values @((Get-GraphPropertyValue -Object $template -PropertyName 'values'))
            $settingWasCreated = $true
        }

        if (-not $valuesMap.ContainsKey('EnableGroupCreation')) {
            $valuesMap['EnableGroupCreation'] = ''
        }
        if (-not $valuesMap.ContainsKey('GroupCreationAllowedGroupId')) {
            $valuesMap['GroupCreationAllowedGroupId'] = ''
        }

        $allowGroupCreationBefore = Get-TrimmedValue -Value $valuesMap['EnableGroupCreation']
        $allowedGroupIdBefore = Get-TrimmedValue -Value $valuesMap['GroupCreationAllowedGroupId']

        if ($null -ne $allowGroupCreation) {
            $valuesMap['EnableGroupCreation'] = if ($allowGroupCreation) { 'true' } else { 'false' }
        }

        if ($clearAllowedGroup) {
            $valuesMap['GroupCreationAllowedGroupId'] = ''
        }
        elseif (-not [string]::IsNullOrWhiteSpace($resolvedAllowedGroup.Id)) {
            $valuesMap['GroupCreationAllowedGroupId'] = $resolvedAllowedGroup.Id
        }

        $allowGroupCreationAfter = Get-TrimmedValue -Value $valuesMap['EnableGroupCreation']
        $allowedGroupIdAfter = Get-TrimmedValue -Value $valuesMap['GroupCreationAllowedGroupId']

        $isValueChange = ($allowGroupCreationBefore -ne $allowGroupCreationAfter) -or ($allowedGroupIdBefore -ne $allowedGroupIdAfter)
        if ((-not $isValueChange) -and (-not $settingWasCreated)) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $policyKey -Action 'SetEntraGroupCreatorsPolicy' -Status 'Skipped' -Message 'Group creators policy is already set to the requested values.'))
            $rowNumber++
            continue
        }

        $targetDescription = if ($settingWasCreated) { 'Create Group.Unified setting and apply group creators policy' } else { 'Update Group.Unified group creators policy' }
        if ($PSCmdlet.ShouldProcess($policyKey, $targetDescription)) {
            $bodyObject = @{
                values = Convert-SettingMapToValues -Map $valuesMap
            }

            if ($settingWasCreated) {
                $bodyObject['templateId'] = $templateId
                $createBody = $bodyObject | ConvertTo-Json -Depth 8 -Compress

                $createdSetting = Invoke-WithRetry -OperationName 'Create Group.Unified setting' -ScriptBlock {
                    Invoke-MgGraphRequest -Method POST -Uri '/v1.0/groupSettings' -Body $createBody -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
                }

                $settingId = Get-TrimmedValue -Value (Get-GraphPropertyValue -Object $createdSetting -PropertyName 'id')
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $policyKey -Action 'SetEntraGroupCreatorsPolicy' -Status 'Created' -Message 'Group creators policy setting created and updated.'))
            }
            else {
                $updateBody = $bodyObject | ConvertTo-Json -Depth 8 -Compress

                Invoke-WithRetry -OperationName 'Update Group.Unified setting' -ScriptBlock {
                    Invoke-MgGraphRequest -Method PATCH -Uri "/v1.0/groupSettings/$settingId" -Body $updateBody -ContentType 'application/json' -ErrorAction Stop | Out-Null
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $policyKey -Action 'SetEntraGroupCreatorsPolicy' -Status 'Updated' -Message 'Group creators policy updated successfully.'))
            }
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $policyKey -Action 'SetEntraGroupCreatorsPolicy' -Status 'WhatIf' -Message 'Policy update skipped due to WhatIf.'))
        }

        $resultIndex = $results.Count - 1
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'SettingId' -NotePropertyValue $settingId -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'AllowGroupCreationBefore' -NotePropertyValue $allowGroupCreationBefore -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'AllowGroupCreationAfter' -NotePropertyValue $allowGroupCreationAfter -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'AllowedGroupIdBefore' -NotePropertyValue $allowedGroupIdBefore -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'AllowedGroupIdAfter' -NotePropertyValue $allowedGroupIdAfter -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'AllowedGroupIdentityTypeRequested' -NotePropertyValue $allowedGroupIdentityType -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'AllowedGroupIdentityValueRequested' -NotePropertyValue $allowedGroupIdentityValue -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'ResolvedAllowedGroupDisplayName' -NotePropertyValue $resolvedAllowedGroup.DisplayName -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'ResolvedAllowedGroupMailNickname' -NotePropertyValue $resolvedAllowedGroup.MailNickname -Force
        Add-Member -InputObject $results[$resultIndex] -NotePropertyName 'SettingWasCreated' -NotePropertyValue ([string]$settingWasCreated) -Force
    }
    catch {
        Write-Status -Message "Row $rowNumber ($policyKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $policyKey -Action 'SetEntraGroupCreatorsPolicy' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

$extendedColumns = @(
    'SettingId',
    'AllowGroupCreationBefore',
    'AllowGroupCreationAfter',
    'AllowedGroupIdBefore',
    'AllowedGroupIdAfter',
    'AllowedGroupIdentityTypeRequested',
    'AllowedGroupIdentityValueRequested',
    'ResolvedAllowedGroupDisplayName',
    'ResolvedAllowedGroupMailNickname',
    'SettingWasCreated'
)

foreach ($result in $results) {
    foreach ($column in $extendedColumns) {
        if ($result.PSObject.Properties.Name -notcontains $column) {
            Add-Member -InputObject $result -NotePropertyName $column -NotePropertyValue '' -Force
        }
    }
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra group creators policy update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}

