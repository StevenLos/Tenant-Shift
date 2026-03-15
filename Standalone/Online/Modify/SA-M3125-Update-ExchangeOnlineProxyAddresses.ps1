<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-151500

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\\..\\Standalone_OutputCsvPath') -ChildPath ("Results_SA-M3125-Update-ExchangeOnlineProxyAddresses_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function ConvertTo-NormalizedSmtpAddress {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    $trimmed = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return ''
    }

    if ($trimmed.StartsWith('SMTP:', [System.StringComparison]::OrdinalIgnoreCase)) {
        $trimmed = $trimmed.Substring(5)
    }

    return $trimmed.ToLowerInvariant()
}

function ConvertTo-ProxyAddressArray {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    $items = ConvertTo-Array -Value $Value
    $deduped = [System.Collections.Generic.List[string]]::new()

    foreach ($item in $items) {
        $trimmed = Get-TrimmedValue -Value $item
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }

        if ($trimmed -notmatch ':') {
            if ($trimmed -match '@') {
                $trimmed = "smtp:$trimmed"
            }
        }

        if (-not ($deduped.Contains($trimmed))) {
            $deduped.Add($trimmed)
        }
    }

    return $deduped.ToArray()
}

function ConvertTo-CanonicalProxyAddressSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PrimarySmtpAddress,

        [Parameter(Mandatory)]
        [string[]]$InputAddresses
    )

    $normalizedPrimary = ConvertTo-NormalizedSmtpAddress -Value $PrimarySmtpAddress
    if ([string]::IsNullOrWhiteSpace($normalizedPrimary)) {
        throw 'Primary SMTP address cannot be empty when building canonical proxy address set.'
    }

    $finalList = [System.Collections.Generic.List[string]]::new()
    $finalList.Add("SMTP:$normalizedPrimary")

    foreach ($entry in @($InputAddresses)) {
        $trimmedEntry = Get-TrimmedValue -Value $entry
        if ([string]::IsNullOrWhiteSpace($trimmedEntry)) {
            continue
        }

        $candidate = ''
        if ($trimmedEntry -match '^(?<prefix>[^:]+):(?<value>.+)$') {
            $prefix = $matches['prefix']
            $value = Get-TrimmedValue -Value $matches['value']

            if ($prefix.Equals('smtp', [System.StringComparison]::OrdinalIgnoreCase)) {
                $smtp = ConvertTo-NormalizedSmtpAddress -Value $value
                if ($smtp -eq $normalizedPrimary) {
                    continue
                }

                $candidate = "smtp:$smtp"
            }
            else {
                $candidate = "{0}:{1}" -f $prefix, $value
            }
        }
        else {
            $smtp = ConvertTo-NormalizedSmtpAddress -Value $trimmedEntry
            if ([string]::IsNullOrWhiteSpace($smtp) -or $smtp -eq $normalizedPrimary) {
                continue
            }

            $candidate = "smtp:$smtp"
        }

        if (-not [string]::IsNullOrWhiteSpace($candidate) -and -not ($finalList.Contains($candidate))) {
            $finalList.Add($candidate)
        }
    }

    return $finalList.ToArray()
}

$requiredHeaders = @(
    'MailboxIdentity',
    'PrimarySmtpAddress',
    'AddProxyAddresses',
    'RemoveProxyAddresses',
    'ReplaceAllProxyAddresses',
    'ClearSecondaryProxyAddresses',
    'Notes'
)

Write-Status -Message 'Starting Exchange Online proxy address update script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = Get-TrimmedValue -Value $row.MailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
            throw 'MailboxIdentity is required.'
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentity" -ScriptBlock {
            Get-EXOMailbox -Identity $mailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'UpdateMailboxProxyAddresses' -Status 'NotFound' -Message 'Mailbox not found.'))
            $rowNumber++
            continue
        }

        $primarySmtpAddress = Get-TrimmedValue -Value $row.PrimarySmtpAddress
        $addAddresses = ConvertTo-ProxyAddressArray -Value ([string]$row.AddProxyAddresses)
        $removeAddresses = ConvertTo-ProxyAddressArray -Value ([string]$row.RemoveProxyAddresses)
        $replaceAddresses = ConvertTo-ProxyAddressArray -Value ([string]$row.ReplaceAllProxyAddresses)

        $clearSecondaryRaw = Get-TrimmedValue -Value $row.ClearSecondaryProxyAddresses
        $clearSecondary = $false
        if (-not [string]::IsNullOrWhiteSpace($clearSecondaryRaw)) {
            $clearSecondary = ConvertTo-Bool -Value $clearSecondaryRaw
        }

        if ($replaceAddresses.Count -gt 0 -and ($addAddresses.Count -gt 0 -or $removeAddresses.Count -gt 0 -or $clearSecondary)) {
            throw 'ReplaceAllProxyAddresses cannot be combined with AddProxyAddresses, RemoveProxyAddresses, or ClearSecondaryProxyAddresses.'
        }

        if ($clearSecondary -and $removeAddresses.Count -gt 0) {
            throw 'ClearSecondaryProxyAddresses cannot be combined with RemoveProxyAddresses.'
        }

        $requestedChange = (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) -or ($addAddresses.Count -gt 0) -or ($removeAddresses.Count -gt 0) -or ($replaceAddresses.Count -gt 0) -or $clearSecondary
        if (-not $requestedChange) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'UpdateMailboxProxyAddresses' -Status 'Skipped' -Message 'No proxy address updates were requested.'))
            $rowNumber++
            continue
        }

        $setParams = @{
            Identity = $mailbox.Identity
        }

        if ($replaceAddresses.Count -gt 0) {
            $targetPrimary = ConvertTo-NormalizedSmtpAddress -Value $primarySmtpAddress
            if ([string]::IsNullOrWhiteSpace($targetPrimary)) {
                $primaryInList = @($replaceAddresses | Where-Object { $_.StartsWith('SMTP:', [System.StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1)
                if ($primaryInList.Count -gt 0) {
                    $targetPrimary = ConvertTo-NormalizedSmtpAddress -Value $primaryInList[0]
                }
            }
            if ([string]::IsNullOrWhiteSpace($targetPrimary) -and $replaceAddresses.Count -gt 0) {
                $targetPrimary = ConvertTo-NormalizedSmtpAddress -Value $replaceAddresses[0]
            }
            if ([string]::IsNullOrWhiteSpace($targetPrimary)) {
                $targetPrimary = ConvertTo-NormalizedSmtpAddress -Value $mailbox.PrimarySmtpAddress
            }

            $setParams.EmailAddresses = ConvertTo-CanonicalProxyAddressSet -PrimarySmtpAddress $targetPrimary -InputAddresses $replaceAddresses

            if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
                $setParams.PrimarySmtpAddress = $primarySmtpAddress
            }
        }
        elseif ($clearSecondary) {
            $targetPrimary = if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) { $primarySmtpAddress } else { ([string]$mailbox.PrimarySmtpAddress).Trim() }
            if ([string]::IsNullOrWhiteSpace($targetPrimary)) {
                throw 'Unable to determine primary SMTP address while clearing secondary proxy addresses.'
            }

            $setParams.EmailAddresses = ConvertTo-CanonicalProxyAddressSet -PrimarySmtpAddress $targetPrimary -InputAddresses $addAddresses

            if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
                $setParams.PrimarySmtpAddress = $primarySmtpAddress
            }
        }
        else {
            if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
                $setParams.PrimarySmtpAddress = $primarySmtpAddress
            }

            $emailAddressOps = @{}
            if ($addAddresses.Count -gt 0) {
                $emailAddressOps['Add'] = $addAddresses
            }
            if ($removeAddresses.Count -gt 0) {
                $emailAddressOps['Remove'] = $removeAddresses
            }

            if ($emailAddressOps.Count -gt 0) {
                $setParams.EmailAddresses = $emailAddressOps
            }
        }

        if ($setParams.Count -eq 1) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'UpdateMailboxProxyAddresses' -Status 'Skipped' -Message 'No effective proxy address updates were generated.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($mailboxIdentity, 'Update mailbox proxy addresses')) {
            Invoke-WithRetry -OperationName "Update mailbox proxy addresses $mailboxIdentity" -ScriptBlock {
                Set-Mailbox @setParams -ErrorAction Stop
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'UpdateMailboxProxyAddresses' -Status 'Updated' -Message 'Mailbox proxy addresses updated successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'UpdateMailboxProxyAddresses' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'UpdateMailboxProxyAddresses' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online proxy address update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}

