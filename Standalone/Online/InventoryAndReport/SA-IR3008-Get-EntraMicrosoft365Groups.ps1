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
Microsoft.Graph.Groups
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\\..\\Standalone_OutputCsvPath') -ChildPath ("Results_SA-IR3008-Get-EntraMicrosoft365Groups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-TrimmedValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    return ([string]$Value).Trim()
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

function Test-IsMicrosoft365Group {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Group
    )

    $groupTypes = @($Group.GroupTypes)
    return (($groupTypes -contains 'Unified') -and ($Group.MailEnabled -eq $true) -and ($Group.SecurityEnabled -eq $false))
}

function Get-DirectoryObjectType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$DirectoryObject
    )

    $odataType = ''
    if ($DirectoryObject.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $additional = $DirectoryObject.AdditionalProperties
        if ($additional) {
            try {
                if ($additional.ContainsKey('@odata.type')) {
                    $odataType = ([string]$additional['@odata.type']).Trim()
                }
            }
            catch {
                # Best effort only.
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($odataType)) {
        return 'unknown'
    }

    $normalized = $odataType.TrimStart('#').ToLowerInvariant()
    if ($normalized.StartsWith('microsoft.graph.')) {
        $normalized = $normalized.Substring('microsoft.graph.'.Length)
    }

    return $normalized
}

$requiredHeaders = @(
    'GroupMailNickname'
)

Write-Status -Message 'Starting Entra ID Microsoft 365 group inventory script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('Group.Read.All', 'User.Read.All', 'Directory.Read.All')

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

$groupSelect = 'id,displayName,description,mail,mailNickname,proxyAddresses,groupTypes,visibility,classification,preferredDataLocation,mailEnabled,securityEnabled,isAssignableToRole,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesDomainName,onPremisesNetBiosName,onPremisesSamAccountName,onPremisesSecurityIdentifier,createdDateTime,renewedDateTime,deletedDateTime'
$allM365GroupsCache = $null
$userById = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

$rowNumber = 1
foreach ($row in $rows) {
    $groupMailNickname = ([string]$row.GroupMailNickname).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($groupMailNickname)) {
            throw 'GroupMailNickname is required. Use * to inventory all Microsoft 365 groups.'
        }

        $groups = @()
        if ($groupMailNickname -eq '*') {
            if ($null -eq $allM365GroupsCache) {
                $allGroups = @(Invoke-WithRetry -OperationName 'Load all groups for Microsoft 365 group inventory' -ScriptBlock {
                    Get-MgGroup -All -Property $groupSelect -ErrorAction Stop
                })
                $allM365GroupsCache = @($allGroups | Where-Object { Test-IsMicrosoft365Group -Group $_ } | Sort-Object -Property MailNickname, DisplayName, Id)
            }

            $groups = @($allM365GroupsCache)
        }
        else {
            $escapedAlias = Escape-ODataString -Value $groupMailNickname
            $candidateGroups = @(Invoke-WithRetry -OperationName "Lookup group alias $groupMailNickname" -ScriptBlock {
                Get-MgGroup -Filter "mailNickname eq '$escapedAlias'" -ConsistencyLevel eventual -Property $groupSelect -ErrorAction Stop
            })

            $groups = @($candidateGroups | Where-Object { Test-IsMicrosoft365Group -Group $_ })
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'GetEntraMicrosoft365Group' -Status 'NotFound' -Message 'No matching Microsoft 365 groups were found.' -Data ([ordered]@{
                        GroupId                        = ''
                        GroupDisplayName               = ''
                        GroupMailNickname              = $groupMailNickname
                        Description                    = ''
                        Mail                           = ''
                        ProxyAddresses                 = ''
                        Visibility                     = ''
                        Classification                 = ''
                        PreferredDataLocation          = ''
                        GroupTypes                     = ''
                        MailEnabled                    = ''
                        SecurityEnabled                = ''
                        IsAssignableToRole             = ''
                        OnPremisesSyncEnabled          = ''
                        OnPremisesLastSyncDateTime     = ''
                        OnPremisesDomainName           = ''
                        OnPremisesNetBiosName          = ''
                        OnPremisesSamAccountName       = ''
                        OnPremisesSecurityIdentifier   = ''
                        CreatedDateTime                = ''
                        RenewedDateTime                = ''
                        DeletedDateTime                = ''
                        OwnersUserPrincipalNames       = ''
                        OwnersObjectIds                = ''
                        OwnersCount                    = ''
                        MembersUserPrincipalNames      = ''
                        MembersObjectIds               = ''
                        MembersCount                   = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property MailNickname, DisplayName, Id)) {
            $groupId = Get-TrimmedValue -Value $group.Id
            $groupDisplayName = Get-TrimmedValue -Value $group.DisplayName
            $resolvedAlias = Get-TrimmedValue -Value $group.MailNickname

            $ownerUpns = [System.Collections.Generic.List[string]]::new()
            $ownerObjectIds = [System.Collections.Generic.List[string]]::new()
            $memberUpns = [System.Collections.Generic.List[string]]::new()
            $memberObjectIds = [System.Collections.Generic.List[string]]::new()

            $owners = @(Invoke-WithRetry -OperationName "Load owners for group $resolvedAlias" -ScriptBlock {
                Get-MgGroupOwner -GroupId $groupId -All -ErrorAction Stop
            })

            foreach ($owner in @($owners | Sort-Object -Property Id)) {
                $ownerId = Get-TrimmedValue -Value $owner.Id
                if ([string]::IsNullOrWhiteSpace($ownerId)) {
                    continue
                }

                if (-not $ownerObjectIds.Contains($ownerId)) {
                    $ownerObjectIds.Add($ownerId)
                }

                if ((Get-DirectoryObjectType -DirectoryObject $owner) -ne 'user') {
                    continue
                }

                try {
                    $ownerUser = $null
                    if ($userById.ContainsKey($ownerId)) {
                        $ownerUser = $userById[$ownerId]
                    }
                    else {
                        $ownerUser = Invoke-WithRetry -OperationName "Load owner user details $ownerId" -ScriptBlock {
                            Get-MgUser -UserId $ownerId -Property 'id,userPrincipalName' -ErrorAction Stop
                        }
                        $userById[$ownerId] = $ownerUser
                    }

                    $ownerUpn = Get-TrimmedValue -Value $ownerUser.UserPrincipalName
                    if (-not [string]::IsNullOrWhiteSpace($ownerUpn) -and -not $ownerUpns.Contains($ownerUpn)) {
                        $ownerUpns.Add($ownerUpn)
                    }
                }
                catch {
                    Write-Status -Message "Owner detail lookup failed for owner ID '$ownerId' in group '$groupDisplayName': $($_.Exception.Message)" -Level WARN
                }
            }

            $members = @(Invoke-WithRetry -OperationName "Load members for group $resolvedAlias" -ScriptBlock {
                Get-MgGroupMember -GroupId $groupId -All -ErrorAction Stop
            })

            foreach ($member in @($members | Sort-Object -Property Id)) {
                $memberId = Get-TrimmedValue -Value $member.Id
                if ([string]::IsNullOrWhiteSpace($memberId)) {
                    continue
                }

                if (-not $memberObjectIds.Contains($memberId)) {
                    $memberObjectIds.Add($memberId)
                }

                if ((Get-DirectoryObjectType -DirectoryObject $member) -ne 'user') {
                    continue
                }

                try {
                    $memberUser = $null
                    if ($userById.ContainsKey($memberId)) {
                        $memberUser = $userById[$memberId]
                    }
                    else {
                        $memberUser = Invoke-WithRetry -OperationName "Load member user details $memberId" -ScriptBlock {
                            Get-MgUser -UserId $memberId -Property 'id,userPrincipalName' -ErrorAction Stop
                        }
                        $userById[$memberId] = $memberUser
                    }

                    $memberUpn = Get-TrimmedValue -Value $memberUser.UserPrincipalName
                    if (-not [string]::IsNullOrWhiteSpace($memberUpn) -and -not $memberUpns.Contains($memberUpn)) {
                        $memberUpns.Add($memberUpn)
                    }
                }
                catch {
                    Write-Status -Message "Member detail lookup failed for member ID '$memberId' in group '$groupDisplayName': $($_.Exception.Message)" -Level WARN
                }
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$resolvedAlias|$groupId" -Action 'GetEntraMicrosoft365Group' -Status 'Completed' -Message 'Microsoft 365 group exported.' -Data ([ordered]@{
                        GroupId                        = $groupId
                        GroupDisplayName               = $groupDisplayName
                        GroupMailNickname              = $resolvedAlias
                        Description                    = Get-TrimmedValue -Value $group.Description
                        Mail                           = Get-TrimmedValue -Value $group.Mail
                        ProxyAddresses                 = Convert-MultiValueToString -Value $group.ProxyAddresses
                        Visibility                     = Get-TrimmedValue -Value $group.Visibility
                        Classification                 = Get-TrimmedValue -Value $group.Classification
                        PreferredDataLocation          = Get-TrimmedValue -Value $group.PreferredDataLocation
                        GroupTypes                     = Convert-MultiValueToString -Value $group.GroupTypes
                        MailEnabled                    = [string]$group.MailEnabled
                        SecurityEnabled                = [string]$group.SecurityEnabled
                        IsAssignableToRole             = [string]$group.IsAssignableToRole
                        OnPremisesSyncEnabled          = [string]$group.OnPremisesSyncEnabled
                        OnPremisesLastSyncDateTime     = [string]$group.OnPremisesLastSyncDateTime
                        OnPremisesDomainName           = Get-TrimmedValue -Value $group.OnPremisesDomainName
                        OnPremisesNetBiosName          = Get-TrimmedValue -Value $group.OnPremisesNetBiosName
                        OnPremisesSamAccountName       = Get-TrimmedValue -Value $group.OnPremisesSamAccountName
                        OnPremisesSecurityIdentifier   = Get-TrimmedValue -Value $group.OnPremisesSecurityIdentifier
                        CreatedDateTime                = [string]$group.CreatedDateTime
                        RenewedDateTime                = [string]$group.RenewedDateTime
                        DeletedDateTime                = [string]$group.DeletedDateTime
                        OwnersUserPrincipalNames       = (@($ownerUpns | Sort-Object) -join ';')
                        OwnersObjectIds                = (@($ownerObjectIds | Sort-Object) -join ';')
                        OwnersCount                    = [string]$ownerObjectIds.Count
                        MembersUserPrincipalNames      = (@($memberUpns | Sort-Object) -join ';')
                        MembersObjectIds               = (@($memberObjectIds | Sort-Object) -join ';')
                        MembersCount                   = [string]$memberObjectIds.Count
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupMailNickname) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupMailNickname -Action 'GetEntraMicrosoft365Group' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    GroupId                        = ''
                    GroupDisplayName               = ''
                    GroupMailNickname              = $groupMailNickname
                    Description                    = ''
                    Mail                           = ''
                    ProxyAddresses                 = ''
                    Visibility                     = ''
                    Classification                 = ''
                    PreferredDataLocation          = ''
                    GroupTypes                     = ''
                    MailEnabled                    = ''
                    SecurityEnabled                = ''
                    IsAssignableToRole             = ''
                    OnPremisesSyncEnabled          = ''
                    OnPremisesLastSyncDateTime     = ''
                    OnPremisesDomainName           = ''
                    OnPremisesNetBiosName          = ''
                    OnPremisesSamAccountName       = ''
                    OnPremisesSecurityIdentifier   = ''
                    CreatedDateTime                = ''
                    RenewedDateTime                = ''
                    DeletedDateTime                = ''
                    OwnersUserPrincipalNames       = ''
                    OwnersObjectIds                = ''
                    OwnersCount                    = ''
                    MembersUserPrincipalNames      = ''
                    MembersObjectIds               = ''
                    MembersCount                   = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID Microsoft 365 group inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}






