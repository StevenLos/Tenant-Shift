<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-004416

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Users
Microsoft.Graph.Users.Actions
Microsoft.Graph.Identity.DirectoryManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\\..\\Standalone_OutputCsvPath') -ChildPath ("Results_SA-M3003-Set-EntraUserLicenses_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'UserPrincipalName',
    'SkuPartNumber',
    'DisabledPlans'
)

Write-Status -Message 'Starting Entra ID user license assignment script.'
Assert-ModuleCurrent -ModuleNames @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Users.Actions',
    'Microsoft.Graph.Identity.DirectoryManagement'
)
Ensure-GraphConnection -RequiredScopes @('User.ReadWrite.All', 'Directory.Read.All', 'Organization.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

Write-Status -Message 'Loading subscribed SKUs from Microsoft Graph.'
$subscribedSkus = @(Invoke-WithRetry -OperationName 'Load subscribed SKUs' -ScriptBlock {
    Get-MgSubscribedSku -All -ErrorAction Stop
})

if ($subscribedSkus.Count -eq 0) {
    throw 'No subscribed SKUs were returned from Microsoft Graph. Confirm that your tenant has at least one available license SKU.'
}

$skuByPartNumber = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$availableUnitsByPartNumber = [System.Collections.Generic.Dictionary[string, int]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($sku in $subscribedSkus) {
    $partNumber = ([string]$sku.SkuPartNumber).Trim()
    if ([string]::IsNullOrWhiteSpace($partNumber)) {
        continue
    }

    if (-not $skuByPartNumber.ContainsKey($partNumber)) {
        $skuByPartNumber[$partNumber] = $sku
    }

    try {
        $enabledUnits = $null
        $consumedUnits = $null

        if ($sku.PSObject.Properties.Name -contains 'PrepaidUnits' -and $sku.PrepaidUnits -and $sku.PrepaidUnits.PSObject.Properties.Name -contains 'Enabled') {
            $enabledUnits = [int]$sku.PrepaidUnits.Enabled
        }

        if ($sku.PSObject.Properties.Name -contains 'ConsumedUnits' -and $null -ne $sku.ConsumedUnits) {
            $consumedUnits = [int]$sku.ConsumedUnits
        }

        if ($null -ne $enabledUnits -and $null -ne $consumedUnits) {
            $availableUnitsByPartNumber[$partNumber] = [Math]::Max($enabledUnits - $consumedUnits, 0)
        }
    }
    catch {
        Write-Status -Message "Unable to calculate available unit count for SKU '$partNumber': $($_.Exception.Message)" -Level WARN
    }
}

if ($skuByPartNumber.Count -eq 0) {
    throw 'No subscribed SKUs with usable SkuPartNumber values were returned from Microsoft Graph.'
}

$userByUpn = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$assignedLicensesByUpn = [System.Collections.Generic.Dictionary[string, object[]]]::new([System.StringComparer]::OrdinalIgnoreCase)

$rowNumber = 1
foreach ($row in $rows) {
    $upn = ([string]$row.UserPrincipalName).Trim()
    $skuPartNumber = ([string]$row.SkuPartNumber).Trim()
    $primaryKey = "$upn|$skuPartNumber"

    try {
        if ([string]::IsNullOrWhiteSpace($upn) -or [string]::IsNullOrWhiteSpace($skuPartNumber)) {
            throw 'UserPrincipalName and SkuPartNumber are required.'
        }

        if (-not $skuByPartNumber.ContainsKey($skuPartNumber)) {
            $availableSkuPartNumbers = @($skuByPartNumber.Keys | Sort-Object)
            throw "SkuPartNumber '$skuPartNumber' was not found in tenant subscriptions. Available values: $($availableSkuPartNumbers -join ', ')"
        }

        $sku = $skuByPartNumber[$skuPartNumber]
        $skuId = [Guid]$sku.SkuId
        $servicePlans = @($sku.ServicePlans)

        $disabledPlanTokens = ConvertTo-Array -Value ([string]$row.DisabledPlans)
        $resolvedDisabledPlans = [System.Collections.Generic.List[Guid]]::new()

        foreach ($token in $disabledPlanTokens) {
            $tokenValue = ([string]$token).Trim()
            if ([string]::IsNullOrWhiteSpace($tokenValue)) {
                continue
            }

            $planId = [Guid]::Empty
            if ([Guid]::TryParse($tokenValue, [ref]$planId)) {
                $planExists = $false
                foreach ($servicePlan in $servicePlans) {
                    $existingPlanId = [Guid]::Empty
                    if ([Guid]::TryParse([string]$servicePlan.ServicePlanId, [ref]$existingPlanId) -and $existingPlanId -eq $planId) {
                        $planExists = $true
                        break
                    }
                }

                if (-not $planExists) {
                    throw "Disabled plan GUID '$tokenValue' is not valid for SKU '$skuPartNumber'."
                }
            }
            else {
                $matchingPlanIds = [System.Collections.Generic.List[Guid]]::new()

                foreach ($servicePlan in $servicePlans) {
                    $servicePlanName = ([string]$servicePlan.ServicePlanName).Trim()
                    if ([string]::IsNullOrWhiteSpace($servicePlanName)) {
                        continue
                    }

                    if ($servicePlanName.Equals($tokenValue, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $servicePlanId = [Guid]::Empty
                        if (-not [Guid]::TryParse([string]$servicePlan.ServicePlanId, [ref]$servicePlanId)) {
                            throw "Service plan '$servicePlanName' on SKU '$skuPartNumber' has an invalid ID."
                        }

                        $matchingPlanIds.Add($servicePlanId)
                    }
                }

                if ($matchingPlanIds.Count -eq 0) {
                    $availablePlans = @(
                        $servicePlans |
                            ForEach-Object { ([string]$_.ServicePlanName).Trim() } |
                            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                            Sort-Object -Unique
                    )
                    throw "Disabled plan '$tokenValue' is not valid for SKU '$skuPartNumber'. Available plans: $($availablePlans -join ', ')"
                }

                if ($matchingPlanIds.Count -gt 1) {
                    throw "Disabled plan '$tokenValue' matched multiple service plans for SKU '$skuPartNumber'. Use a service plan GUID to disambiguate."
                }

                $planId = $matchingPlanIds[0]
            }

            if (-not ($resolvedDisabledPlans -contains $planId)) {
                $resolvedDisabledPlans.Add($planId)
            }
        }

        if ($userByUpn.ContainsKey($upn)) {
            $user = $userByUpn[$upn]
        }
        else {
            $escapedUpn = Escape-ODataString -Value $upn
            $users = @(Invoke-WithRetry -OperationName "Lookup user $upn" -ScriptBlock {
                Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -ErrorAction Stop
            })

            if ($users.Count -eq 0) {
                throw "User '$upn' was not found."
            }

            if ($users.Count -gt 1) {
                throw "Multiple users were returned for UPN '$upn'. Resolve duplicate directory objects before retrying."
            }

            $user = $users[0]
            $userByUpn[$upn] = $user
        }

        if ($assignedLicensesByUpn.ContainsKey($upn)) {
            $assignedLicenses = @($assignedLicensesByUpn[$upn])
        }
        else {
            $userWithLicenses = Invoke-WithRetry -OperationName "Load assigned licenses for $upn" -ScriptBlock {
                Get-MgUser -UserId $user.Id -Property 'id,userPrincipalName,assignedLicenses' -ErrorAction Stop
            }
            $assignedLicenses = @($userWithLicenses.AssignedLicenses)
            $assignedLicensesByUpn[$upn] = $assignedLicenses
        }

        $currentAssignments = @($assignedLicenses | Where-Object { [Guid]$_.SkuId -eq $skuId })
        if ($currentAssignments.Count -gt 1) {
            throw "User '$upn' has duplicate assignments for SKU '$skuPartNumber'. Resolve license state before retrying."
        }

        $requestedDisabledCanonical = @(
            $resolvedDisabledPlans |
                ForEach-Object { $_.Guid.ToLowerInvariant() } |
                Sort-Object -Unique
        )

        $currentDisabledCanonical = @()
        if ($currentAssignments.Count -eq 1) {
            $currentDisabledCanonical = @(
                @($currentAssignments[0].DisabledPlans) |
                    Where-Object { $null -ne $_ } |
                    ForEach-Object {
                        $currentDisabledGuid = [Guid]::Empty
                        if (-not [Guid]::TryParse([string]$_, [ref]$currentDisabledGuid)) {
                            throw "User '$upn' has an invalid disabled plan ID '$([string]$_)' for SKU '$skuPartNumber'."
                        }

                        $currentDisabledGuid.Guid.ToLowerInvariant()
                    } |
                    Sort-Object -Unique
            )
        }

        $requestedMinusCurrent = @($requestedDisabledCanonical | Where-Object { $_ -notin $currentDisabledCanonical })
        $currentMinusRequested = @($currentDisabledCanonical | Where-Object { $_ -notin $requestedDisabledCanonical })
        $disabledPlansMatch = ($requestedMinusCurrent.Count -eq 0 -and $currentMinusRequested.Count -eq 0)

        if ($currentAssignments.Count -eq 1 -and $disabledPlansMatch) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetUserLicense' -Status 'Skipped' -Message "License '$skuPartNumber' is already assigned with the requested disabled plans."))
            $rowNumber++
            continue
        }

        $isNewAssignment = ($currentAssignments.Count -eq 0)

        $addLicense = @{
            SkuId         = $skuId
            DisabledPlans = @($resolvedDisabledPlans)
        }

        $operationLabel = if ($isNewAssignment) { 'Assign license' } else { 'Update license options' }
        if ($PSCmdlet.ShouldProcess($primaryKey, $operationLabel)) {
            if ($isNewAssignment -and $availableUnitsByPartNumber.ContainsKey($skuPartNumber) -and $availableUnitsByPartNumber[$skuPartNumber] -le 0) {
                throw "No available units remain for SKU '$skuPartNumber'."
            }

            Invoke-WithRetry -OperationName "$operationLabel $primaryKey" -ScriptBlock {
                Set-MgUserLicense -UserId $user.Id -AddLicenses @($addLicense) -RemoveLicenses @() -ErrorAction Stop | Out-Null
            }

            $refreshedUser = Invoke-WithRetry -OperationName "Refresh assigned licenses for $upn" -ScriptBlock {
                Get-MgUser -UserId $user.Id -Property 'id,userPrincipalName,assignedLicenses' -ErrorAction Stop
            }
            $assignedLicensesByUpn[$upn] = @($refreshedUser.AssignedLicenses)

            if ($isNewAssignment -and $availableUnitsByPartNumber.ContainsKey($skuPartNumber)) {
                $availableUnitsByPartNumber[$skuPartNumber] = [Math]::Max(($availableUnitsByPartNumber[$skuPartNumber] - 1), 0)
            }

            $resultStatus = if ($isNewAssignment) { 'Assigned' } else { 'Updated' }
            $disabledPlanSummary = if ($resolvedDisabledPlans.Count -gt 0) {
                " Disabled plans applied: $($resolvedDisabledPlans.Count)."
            }
            else {
                ' No disabled plans were specified.'
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetUserLicense' -Status $resultStatus -Message "License '$skuPartNumber' processed successfully.$disabledPlanSummary"))
        }
        else {
            $whatIfMessage = if ($isNewAssignment) {
                'License assignment skipped due to WhatIf.'
            }
            else {
                'License option update skipped due to WhatIf.'
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetUserLicense' -Status 'WhatIf' -Message $whatIfMessage))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetUserLicense' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID user license assignment script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}








