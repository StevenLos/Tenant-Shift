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
None declared in this file

.MODULEVERSIONPOLICY
Not declared in this file
#>
Set-StrictMode -Version Latest

# Load shared utility functions (dot-sourced, not Import-Module).
# Write-Status is used as sentinel: if it is already in scope, Shared.Common.ps1 was
# already dot-sourced by a prior module load in this session — skip to avoid redefining.
if (-not (Get-Command Write-Status -ErrorAction SilentlyContinue)) {
    . "$PSScriptRoot\..\Shared\Shared.Common.ps1"
}

function Escape-ODataString {
    [CmdletBinding()]
    # Suppression justification: Internal platform helper. Verb is intentional and
    # predates PSScriptAnalyzer enforcement; renaming requires updating all dependent scripts.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
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

function Ensure-GraphConnection {
    [CmdletBinding()]
    # Suppression justification: Internal platform helper. Verb is intentional and
    # predates PSScriptAnalyzer enforcement; renaming requires updating all dependent scripts.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
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
    # Suppression justification: Internal platform helper. Verb is intentional and
    # predates PSScriptAnalyzer enforcement; renaming requires updating all dependent scripts.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
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

function Get-ExchangeOnlineCommandDefinition {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CommandName
    )

    $cacheVariable = Get-Variable -Name ExchangeOnlineCommandDefinitionCache -Scope Script -ErrorAction SilentlyContinue
    if (-not $cacheVariable) {
        $script:ExchangeOnlineCommandDefinitionCache = @{}
    }

    if (-not $script:ExchangeOnlineCommandDefinitionCache.ContainsKey($CommandName)) {
        $definition = Get-Command -Name $CommandName -ErrorAction SilentlyContinue
        if (-not $definition) {
            throw "Required Exchange Online cmdlet '$CommandName' was not found. Ensure ExchangeOnlineManagement is installed and connected."
        }

        $script:ExchangeOnlineCommandDefinitionCache[$CommandName] = $definition
    }

    return $script:ExchangeOnlineCommandDefinitionCache[$CommandName]
}

function Add-SupportedParameter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$ParameterHashtable,

        [Parameter(Mandatory)]
        [string]$CommandName,

        [Parameter(Mandatory)]
        [string]$ParameterName,

        [AllowNull()]
        [object]$Value,

        [switch]$AllowNull,
        [switch]$AllowEmptyString
    )

    if ($null -eq $Value -and -not $AllowNull) {
        return
    }

    if (-not $AllowEmptyString -and $Value -is [string] -and [string]::IsNullOrWhiteSpace([string]$Value)) {
        return
    }

    $definition = Get-ExchangeOnlineCommandDefinition -CommandName $CommandName
    if ($definition.Parameters.ContainsKey($ParameterName)) {
        $ParameterHashtable[$ParameterName] = $Value
    }
}

function ConvertTo-ExchangeResultSize {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [int] -or $Value -is [long]) {
        return $Value
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    if ($text -match '^\d+$') {
        return [int]$text
    }

    return $text
}

function Get-EffectiveErrorActionPreference {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$BoundParameters
    )

    if ($BoundParameters.ContainsKey('ErrorAction')) {
        return [System.Management.Automation.ActionPreference]$BoundParameters['ErrorAction']
    }

    return [System.Management.Automation.ActionPreference]::Stop
}

function Invoke-ExchangeOnlineGetCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CommandName,

        [Parameter(Mandatory)]
        [hashtable]$Parameters,

        [Parameter(Mandatory)]
        [string]$OperationName,

        [System.Management.Automation.ActionPreference]$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
    )

    Get-ExchangeOnlineCommandDefinition -CommandName $CommandName | Out-Null

    if ($ErrorActionPreference -eq [System.Management.Automation.ActionPreference]::Stop) {
        return (Invoke-WithRetry -OperationName $OperationName -ScriptBlock {
                & $CommandName @Parameters -ErrorAction Stop
            })
    }

    return (& $CommandName @Parameters -ErrorAction $ErrorActionPreference)
}

function Get-ExchangeOnlineMailbox {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Identity,

        [AllowNull()]
        [string[]]$RecipientTypeDetails,

        [AllowNull()]
        [object]$ResultSize,

        [AllowNull()]
        [string[]]$Properties,

        [AllowNull()]
        [string[]]$PropertySets
    )

    $commandName = 'Get-EXOMailbox'
    $params = @{}

    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Identity' -Value $Identity
    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'RecipientTypeDetails' -Value $RecipientTypeDetails

    if ($PSBoundParameters.ContainsKey('ResultSize')) {
        $resultSizeValue = ConvertTo-ExchangeResultSize -Value $ResultSize
        Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'ResultSize' -Value $resultSizeValue
    }

    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Properties' -Value $Properties
    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'PropertySets' -Value $PropertySets

    return (Invoke-ExchangeOnlineGetCommand -CommandName $commandName -Parameters $params -OperationName 'Get Exchange Online mailbox (EXO)' -ErrorActionPreference (Get-EffectiveErrorActionPreference -BoundParameters $PSBoundParameters))
}

function Get-ExchangeOnlineRecipient {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Identity,

        [AllowNull()]
        [object]$ResultSize,

        [AllowNull()]
        [string[]]$RecipientTypeDetails,

        [AllowNull()]
        [string[]]$Properties
    )

    $commandName = 'Get-EXORecipient'
    $params = @{}

    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Identity' -Value $Identity
    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'RecipientTypeDetails' -Value $RecipientTypeDetails

    if ($PSBoundParameters.ContainsKey('ResultSize')) {
        $resultSizeValue = ConvertTo-ExchangeResultSize -Value $ResultSize
        Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'ResultSize' -Value $resultSizeValue
    }

    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Properties' -Value $Properties

    return (Invoke-ExchangeOnlineGetCommand -CommandName $commandName -Parameters $params -OperationName 'Get Exchange Online recipient (EXO)' -ErrorActionPreference (Get-EffectiveErrorActionPreference -BoundParameters $PSBoundParameters))
}

function Get-ExchangeOnlineCasMailbox {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Identity,

        [AllowNull()]
        [object]$ResultSize,

        [AllowNull()]
        [string[]]$Properties
    )

    $commandName = 'Get-EXOCasMailbox'
    $params = @{}

    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Identity' -Value $Identity

    if ($PSBoundParameters.ContainsKey('ResultSize')) {
        $resultSizeValue = ConvertTo-ExchangeResultSize -Value $ResultSize
        Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'ResultSize' -Value $resultSizeValue
    }

    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Properties' -Value $Properties

    return (Invoke-ExchangeOnlineGetCommand -CommandName $commandName -Parameters $params -OperationName 'Get Exchange Online CAS mailbox (EXO)' -ErrorActionPreference (Get-EffectiveErrorActionPreference -BoundParameters $PSBoundParameters))
}

function Get-ExchangeOnlineMailboxPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$User
    )

    $commandName = 'Get-EXOMailboxPermission'
    $params = @{}

    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Identity' -Value $Identity
    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'User' -Value $User

    return (Invoke-ExchangeOnlineGetCommand -CommandName $commandName -Parameters $params -OperationName 'Get Exchange Online mailbox permissions (EXO)' -ErrorActionPreference (Get-EffectiveErrorActionPreference -BoundParameters $PSBoundParameters))
}

function Get-ExchangeOnlineRecipientPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$Trustee
    )

    $commandName = 'Get-EXORecipientPermission'
    $params = @{}

    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Identity' -Value $Identity
    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Trustee' -Value $Trustee

    return (Invoke-ExchangeOnlineGetCommand -CommandName $commandName -Parameters $params -OperationName 'Get Exchange Online recipient permissions (EXO)' -ErrorActionPreference (Get-EffectiveErrorActionPreference -BoundParameters $PSBoundParameters))
}

function Get-ExchangeOnlineMailboxStatistics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,

        [switch]$Archive
    )

    $commandName = 'Get-EXOMailboxStatistics'
    $params = @{}

    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Identity' -Value $Identity
    if ($Archive.IsPresent) {
        Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Archive' -Value $true
    }

    return (Invoke-ExchangeOnlineGetCommand -CommandName $commandName -Parameters $params -OperationName 'Get Exchange Online mailbox statistics (EXO)' -ErrorActionPreference (Get-EffectiveErrorActionPreference -BoundParameters $PSBoundParameters))
}

function Get-ExchangeOnlineMailboxFolderStatistics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$FolderScope,

        [switch]$IncludeOldestAndNewestItems
    )

    $commandName = 'Get-EXOMailboxFolderStatistics'
    $params = @{}

    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Identity' -Value $Identity
    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'FolderScope' -Value $FolderScope
    if ($IncludeOldestAndNewestItems.IsPresent) {
        Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'IncludeOldestAndNewestItems' -Value $true
    }

    return (Invoke-ExchangeOnlineGetCommand -CommandName $commandName -Parameters $params -OperationName 'Get Exchange Online mailbox folder statistics (EXO)' -ErrorActionPreference (Get-EffectiveErrorActionPreference -BoundParameters $PSBoundParameters))
}

function Get-ExchangeOnlineMailboxFolderPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$User
    )

    $commandName = 'Get-EXOMailboxFolderPermission'
    $params = @{}

    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Identity' -Value $Identity
    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'User' -Value $User

    return (Invoke-ExchangeOnlineGetCommand -CommandName $commandName -Parameters $params -OperationName 'Get Exchange Online mailbox folder permissions (EXO)' -ErrorActionPreference (Get-EffectiveErrorActionPreference -BoundParameters $PSBoundParameters))
}

function Get-ExchangeOnlineMobileDeviceStatistics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,

        [switch]$IncludeAnalysis
    )

    $commandName = 'Get-EXOMobileDeviceStatistics'
    $params = @{}

    Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'Identity' -Value $Identity
    if ($IncludeAnalysis.IsPresent) {
        Add-SupportedParameter -ParameterHashtable $params -CommandName $commandName -ParameterName 'IncludeAnalysis' -Value $true
    }

    return (Invoke-ExchangeOnlineGetCommand -CommandName $commandName -Parameters $params -OperationName 'Get Exchange Online mobile device statistics (EXO)' -ErrorActionPreference (Get-EffectiveErrorActionPreference -BoundParameters $PSBoundParameters))
}

function Import-SharePointOnlineManagementModule {
    [CmdletBinding()]
    param()

    $moduleName = 'Microsoft.Online.SharePoint.PowerShell'
    $requiredCommands = @(
        'Connect-SPOService',
        'Get-SPOTenant',
        'Get-SPOSite'
    )

    $missingCommands = @($requiredCommands | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($missingCommands.Count -eq 0) {
        return
    }

    $importParams = @{
        Name        = $moduleName
        Global      = $true
        ErrorAction = 'Stop'
    }

    if ($PSVersionTable.PSEdition -eq 'Core') {
        if ($IsWindows) {
            $importParams['UseWindowsPowerShell'] = $true
        }
        else {
            throw "Module '$moduleName' requires Windows PowerShell compatibility when used from PowerShell 7."
        }
    }

    try {
        Import-Module @importParams | Out-Null
    }
    catch {
        throw "Failed to import required module '$moduleName'. Error: $($_.Exception.Message)"
    }

    $missingCommands = @($requiredCommands | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($missingCommands.Count -gt 0) {
        throw "Module '$moduleName' imported, but required cmdlets are still unavailable: $($missingCommands -join ', ')."
    }
}

function Test-SharePointConnection {
    [CmdletBinding()]
    param()

    try {
        Import-SharePointOnlineManagementModule
        Get-SPOTenant -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Ensure-SharePointConnection {
    [CmdletBinding()]
    # Suppression justification: Internal platform helper. Verb is intentional and
    # predates PSScriptAnalyzer enforcement; renaming requires updating all dependent scripts.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
    param(
        [Parameter(Mandatory)]
        [string]$AdminUrl
    )

    Import-SharePointOnlineManagementModule

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

Export-ModuleMember -Function @(
    # Shared utility functions — loaded via dot-source of Shared.Common.ps1 (guard at top of file)
    'Write-Status',
    'Start-RunTranscript',
    'Stop-RunTranscript',
    'ConvertTo-Bool',
    'ConvertTo-Array',
    'Import-ValidatedCsv',
    'New-ResultObject',
    'Export-ResultsCsv',
    'Get-TrimmedValue',
    'Convert-MultiValueToString',
    'Convert-ToOrderedReportObject',
    # Online-only functions
    'Escape-ODataString',
    'Assert-ModuleCurrent',
    'Ensure-GraphConnection',
    'Ensure-ExchangeConnection',
    'Get-ExchangeOnlineMailbox',
    'Get-ExchangeOnlineRecipient',
    'Get-ExchangeOnlineCasMailbox',
    'Get-ExchangeOnlineMailboxPermission',
    'Get-ExchangeOnlineRecipientPermission',
    'Get-ExchangeOnlineMailboxStatistics',
    'Get-ExchangeOnlineMailboxFolderStatistics',
    'Get-ExchangeOnlineMailboxFolderPermission',
    'Get-ExchangeOnlineMobileDeviceStatistics',
    'Ensure-SharePointConnection',
    'Invoke-WithRetry'
)
