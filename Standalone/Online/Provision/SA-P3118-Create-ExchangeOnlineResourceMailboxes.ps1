<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-005957

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\\..\\Standalone_OutputCsvPath') -ChildPath ("Results_SA-P3118-Create-ExchangeOnlineResourceMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$requiredHeaders = @(
    'ResourceType',
    'Name',
    'Alias',
    'DisplayName',
    'UserPrincipalName',
    'PrimarySmtpAddress',
    'HiddenFromAddressListsEnabled',
    'ResourceCapacity',
    'Office',
    'Phone',
    'AutomateProcessing',
    'ForwardRequestsToDelegates',
    'AllBookInPolicy',
    'AllRequestInPolicy',
    'AllRequestOutOfPolicy',
    'BookInPolicy',
    'RequestInPolicy',
    'RequestOutOfPolicy',
    'BookingWindowInDays',
    'MaximumDurationInMinutes',
    'AllowRecurringMeetings',
    'EnforceSchedulingHorizon',
    'ScheduleOnlyDuringWorkHours'
)

Write-Status -Message 'Starting Exchange Online resource mailbox creation script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$newMailboxCommand = Get-Command -Name New-Mailbox -ErrorAction Stop
$newMailboxSupportsUserPrincipalName = $newMailboxCommand.Parameters.ContainsKey('UserPrincipalName')
$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop
$setCalendarCommand = Get-Command -Name Set-CalendarProcessing -ErrorAction Stop

$setMailboxSupportsResourceCapacity = $setMailboxCommand.Parameters.ContainsKey('ResourceCapacity')
$setMailboxSupportsOffice = $setMailboxCommand.Parameters.ContainsKey('Office')
$setMailboxSupportsPhone = $setMailboxCommand.Parameters.ContainsKey('Phone')

$setCalendarSupports = @{
    AutomateProcessing        = $setCalendarCommand.Parameters.ContainsKey('AutomateProcessing')
    ForwardRequestsToDelegates= $setCalendarCommand.Parameters.ContainsKey('ForwardRequestsToDelegates')
    AllBookInPolicy           = $setCalendarCommand.Parameters.ContainsKey('AllBookInPolicy')
    AllRequestInPolicy        = $setCalendarCommand.Parameters.ContainsKey('AllRequestInPolicy')
    AllRequestOutOfPolicy     = $setCalendarCommand.Parameters.ContainsKey('AllRequestOutOfPolicy')
    BookInPolicy              = $setCalendarCommand.Parameters.ContainsKey('BookInPolicy')
    RequestInPolicy           = $setCalendarCommand.Parameters.ContainsKey('RequestInPolicy')
    RequestOutOfPolicy        = $setCalendarCommand.Parameters.ContainsKey('RequestOutOfPolicy')
    BookingWindowInDays       = $setCalendarCommand.Parameters.ContainsKey('BookingWindowInDays')
    MaximumDurationInMinutes  = $setCalendarCommand.Parameters.ContainsKey('MaximumDurationInMinutes')
    AllowRecurringMeetings    = $setCalendarCommand.Parameters.ContainsKey('AllowRecurringMeetings')
    EnforceSchedulingHorizon  = $setCalendarCommand.Parameters.ContainsKey('EnforceSchedulingHorizon')
    ScheduleOnlyDuringWorkHours = $setCalendarCommand.Parameters.ContainsKey('ScheduleOnlyDuringWorkHours')
}

if (-not $newMailboxSupportsUserPrincipalName) {
    Write-Status -Message "New-Mailbox in this session does not support -UserPrincipalName. The 'UserPrincipalName' CSV value will be ignored." -Level WARN
}

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $name = ([string]$row.Name).Trim()
    $resourceTypeRaw = ([string]$row.ResourceType).Trim()
    $resourceType = $resourceTypeRaw.ToLowerInvariant()

    try {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Name is required.'
        }

        if ($resourceType -notin @('room', 'equipment')) {
            throw "ResourceType '$resourceTypeRaw' is invalid. Use Room or Equipment."
        }

        $alias = ([string]$row.Alias).Trim()
        $displayName = ([string]$row.DisplayName).Trim()
        $userPrincipalName = ([string]$row.UserPrincipalName).Trim()
        $primarySmtpAddress = ([string]$row.PrimarySmtpAddress).Trim()
        $office = ([string]$row.Office).Trim()
        $phone = ([string]$row.Phone).Trim()

        $lookupIdentity = if (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) {
            $userPrincipalName
        }
        elseif (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $primarySmtpAddress
        }
        elseif (-not [string]::IsNullOrWhiteSpace($alias)) {
            $alias
        }
        else {
            $name
        }

        $existingMailbox = Invoke-WithRetry -OperationName "Lookup resource mailbox $lookupIdentity" -ScriptBlock {
            Get-EXOMailbox -Identity $lookupIdentity -ErrorAction SilentlyContinue
        }

        if ($existingMailbox) {
            $recipientTypeDetails = ([string]$existingMailbox.RecipientTypeDetails).Trim()
            if ($resourceType -eq 'room' -and $recipientTypeDetails -eq 'RoomMailbox') {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateResourceMailbox' -Status 'Skipped' -Message 'Room mailbox already exists.'))
                $rowNumber++
                continue
            }

            if ($resourceType -eq 'equipment' -and $recipientTypeDetails -eq 'EquipmentMailbox') {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateResourceMailbox' -Status 'Skipped' -Message 'Equipment mailbox already exists.'))
                $rowNumber++
                continue
            }

            throw "Mailbox '$lookupIdentity' already exists with recipient type '$recipientTypeDetails', which does not match requested resource type '$resourceTypeRaw'."
        }

        $createParams = @{
            Name = $name
        }

        if ($resourceType -eq 'room') {
            $createParams.Room = $true
        }
        else {
            $createParams.Equipment = $true
        }

        if (-not [string]::IsNullOrWhiteSpace($alias)) {
            $createParams.Alias = $alias
        }

        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $createParams.DisplayName = $displayName
        }

        $upnIgnored = $false
        if (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) {
            if ($newMailboxSupportsUserPrincipalName) {
                $createParams.UserPrincipalName = $userPrincipalName
            }
            else {
                $upnIgnored = $true
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $createParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $resourceCapacity = 0
        $resourceCapacityRaw = ([string]$row.ResourceCapacity).Trim()
        $setResourceCapacity = $false
        if (-not [string]::IsNullOrWhiteSpace($resourceCapacityRaw)) {
            if (-not [int]::TryParse($resourceCapacityRaw, [ref]$resourceCapacity)) {
                throw "ResourceCapacity '$resourceCapacityRaw' is not a valid integer."
            }
            if ($resourceCapacity -lt 0) {
                throw 'ResourceCapacity must be zero or greater.'
            }
            $setResourceCapacity = $true
        }

        if ($PSCmdlet.ShouldProcess($lookupIdentity, 'Create Exchange Online resource mailbox')) {
            $createdMailbox = Invoke-WithRetry -OperationName "Create resource mailbox $lookupIdentity" -ScriptBlock {
                New-Mailbox @createParams -ErrorAction Stop
            }

            $setMailboxParams = @{
                Identity = $createdMailbox.Identity
            }
            $messages = [System.Collections.Generic.List[string]]::new()

            $hiddenRaw = ([string]$row.HiddenFromAddressListsEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
                $setMailboxParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
            }

            if ($setResourceCapacity) {
                if ($setMailboxSupportsResourceCapacity) {
                    $setMailboxParams.ResourceCapacity = $resourceCapacity
                }
                else {
                    $messages.Add('ResourceCapacity ignored (unsupported parameter).')
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($office)) {
                if ($setMailboxSupportsOffice) {
                    $setMailboxParams.Office = $office
                }
                else {
                    $messages.Add('Office ignored (unsupported parameter).')
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($phone)) {
                if ($setMailboxSupportsPhone) {
                    $setMailboxParams.Phone = $phone
                }
                else {
                    $messages.Add('Phone ignored (unsupported parameter).')
                }
            }

            if ($setMailboxParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Set resource mailbox options $lookupIdentity" -ScriptBlock {
                    Set-Mailbox @setMailboxParams -ErrorAction Stop
                }
            }

            $setCalendarParams = @{
                Identity = $createdMailbox.Identity
            }

            $automateProcessing = ([string]$row.AutomateProcessing).Trim()
            if (-not [string]::IsNullOrWhiteSpace($automateProcessing) -and $setCalendarSupports.AutomateProcessing) {
                $setCalendarParams.AutomateProcessing = $automateProcessing
            }

            $forwardRequestsToDelegatesRaw = ([string]$row.ForwardRequestsToDelegates).Trim()
            if (-not [string]::IsNullOrWhiteSpace($forwardRequestsToDelegatesRaw) -and $setCalendarSupports.ForwardRequestsToDelegates) {
                $setCalendarParams.ForwardRequestsToDelegates = ConvertTo-Bool -Value $forwardRequestsToDelegatesRaw
            }

            $allBookInPolicyRaw = ([string]$row.AllBookInPolicy).Trim()
            if (-not [string]::IsNullOrWhiteSpace($allBookInPolicyRaw) -and $setCalendarSupports.AllBookInPolicy) {
                $setCalendarParams.AllBookInPolicy = ConvertTo-Bool -Value $allBookInPolicyRaw
            }

            $allRequestInPolicyRaw = ([string]$row.AllRequestInPolicy).Trim()
            if (-not [string]::IsNullOrWhiteSpace($allRequestInPolicyRaw) -and $setCalendarSupports.AllRequestInPolicy) {
                $setCalendarParams.AllRequestInPolicy = ConvertTo-Bool -Value $allRequestInPolicyRaw
            }

            $allRequestOutOfPolicyRaw = ([string]$row.AllRequestOutOfPolicy).Trim()
            if (-not [string]::IsNullOrWhiteSpace($allRequestOutOfPolicyRaw) -and $setCalendarSupports.AllRequestOutOfPolicy) {
                $setCalendarParams.AllRequestOutOfPolicy = ConvertTo-Bool -Value $allRequestOutOfPolicyRaw
            }

            $bookInPolicy = ConvertTo-Array -Value ([string]$row.BookInPolicy)
            if ($bookInPolicy.Count -gt 0 -and $setCalendarSupports.BookInPolicy) {
                $setCalendarParams.BookInPolicy = $bookInPolicy
            }

            $requestInPolicy = ConvertTo-Array -Value ([string]$row.RequestInPolicy)
            if ($requestInPolicy.Count -gt 0 -and $setCalendarSupports.RequestInPolicy) {
                $setCalendarParams.RequestInPolicy = $requestInPolicy
            }

            $requestOutOfPolicy = ConvertTo-Array -Value ([string]$row.RequestOutOfPolicy)
            if ($requestOutOfPolicy.Count -gt 0 -and $setCalendarSupports.RequestOutOfPolicy) {
                $setCalendarParams.RequestOutOfPolicy = $requestOutOfPolicy
            }

            $bookingWindowInDaysRaw = ([string]$row.BookingWindowInDays).Trim()
            if (-not [string]::IsNullOrWhiteSpace($bookingWindowInDaysRaw) -and $setCalendarSupports.BookingWindowInDays) {
                $bookingWindowInDays = 0
                if (-not [int]::TryParse($bookingWindowInDaysRaw, [ref]$bookingWindowInDays)) {
                    throw "BookingWindowInDays '$bookingWindowInDaysRaw' is not a valid integer."
                }
                $setCalendarParams.BookingWindowInDays = $bookingWindowInDays
            }

            $maximumDurationRaw = ([string]$row.MaximumDurationInMinutes).Trim()
            if (-not [string]::IsNullOrWhiteSpace($maximumDurationRaw) -and $setCalendarSupports.MaximumDurationInMinutes) {
                $maximumDuration = 0
                if (-not [int]::TryParse($maximumDurationRaw, [ref]$maximumDuration)) {
                    throw "MaximumDurationInMinutes '$maximumDurationRaw' is not a valid integer."
                }
                $setCalendarParams.MaximumDurationInMinutes = $maximumDuration
            }

            $allowRecurringMeetingsRaw = ([string]$row.AllowRecurringMeetings).Trim()
            if (-not [string]::IsNullOrWhiteSpace($allowRecurringMeetingsRaw) -and $setCalendarSupports.AllowRecurringMeetings) {
                $setCalendarParams.AllowRecurringMeetings = ConvertTo-Bool -Value $allowRecurringMeetingsRaw
            }

            $enforceSchedulingHorizonRaw = ([string]$row.EnforceSchedulingHorizon).Trim()
            if (-not [string]::IsNullOrWhiteSpace($enforceSchedulingHorizonRaw) -and $setCalendarSupports.EnforceSchedulingHorizon) {
                $setCalendarParams.EnforceSchedulingHorizon = ConvertTo-Bool -Value $enforceSchedulingHorizonRaw
            }

            $scheduleOnlyDuringWorkHoursRaw = ([string]$row.ScheduleOnlyDuringWorkHours).Trim()
            if (-not [string]::IsNullOrWhiteSpace($scheduleOnlyDuringWorkHoursRaw) -and $setCalendarSupports.ScheduleOnlyDuringWorkHours) {
                $setCalendarParams.ScheduleOnlyDuringWorkHours = ConvertTo-Bool -Value $scheduleOnlyDuringWorkHoursRaw
            }

            if ($setCalendarParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Set calendar processing for $lookupIdentity" -ScriptBlock {
                    Set-CalendarProcessing @setCalendarParams -ErrorAction Stop
                }
            }

            $successMessage = 'Resource mailbox created successfully.'
            if ($upnIgnored) {
                $successMessage = "$successMessage UserPrincipalName was provided but ignored because this New-Mailbox session does not support -UserPrincipalName."
            }
            if ($messages.Count -gt 0) {
                $successMessage = "$successMessage $($messages -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateResourceMailbox' -Status 'Created' -Message $successMessage))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateResourceMailbox' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($name) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $name -Action 'CreateResourceMailbox' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online resource mailbox creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}






