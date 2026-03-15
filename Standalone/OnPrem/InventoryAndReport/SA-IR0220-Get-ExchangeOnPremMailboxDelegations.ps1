<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-235500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)
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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\\..\\Standalone_OutputCsvPath') -ChildPath ("Results_SA-IR0220-Get-ExchangeOnPremMailboxDelegations_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return
    }

    $command = Get-Command -Name $CommandName -ErrorAction Stop
    if ($command.Parameters.ContainsKey($ParameterName)) {
        $ParameterHashtable[$ParameterName] = $text
    }
}

function Resolve-MailboxesByScope {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,

        [AllowEmptyString()]
        [string]$SearchBase,

        [AllowEmptyString()]
        [string]$Server
    )

    if ($Identity -eq '*') {
        $params = @{
            RecipientTypeDetails = @('UserMailbox', 'SharedMailbox')
            ResultSize           = 'Unlimited'
            ErrorAction          = 'Stop'
        }

        Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-Mailbox' -ParameterName 'OrganizationalUnit' -Value $SearchBase
        Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-Mailbox' -ParameterName 'DomainController' -Value $Server

        return @(Get-Mailbox @params)
    }

    $params = @{
        Identity    = $Identity
        ErrorAction = 'SilentlyContinue'
    }

    Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-Mailbox' -ParameterName 'DomainController' -Value $Server

    $mailbox = Get-Mailbox @params
    if ($mailbox -and (Get-TrimmedValue -Value $mailbox.RecipientTypeDetails) -in @('UserMailbox', 'SharedMailbox')) {
        return @($mailbox)
    }

    return @()
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

function Normalize-TrusteeKey {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    return $text.ToLowerInvariant()
}

function Get-DelegateHint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$DelegateObject
    )

    $candidateProperties = @('DistinguishedName', 'Name', 'PrimarySmtpAddress', 'Identity')
    foreach ($propertyName in $candidateProperties) {
        $property = $DelegateObject.PSObject.Properties[$propertyName]
        if ($property) {
            $text = Get-TrimmedValue -Value $property.Value
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                return $text
            }
        }
    }

    return Get-TrimmedValue -Value $DelegateObject
}

$requiredHeaders = @(
    'MailboxIdentity'
)

Write-Status -Message 'Starting Exchange on-prem mailbox delegation inventory script.'
Ensure-ExchangeOnPremConnection

$hasRecipientPermissionCmdlet = [bool](Get-Command -Name Get-RecipientPermission -ErrorAction SilentlyContinue)
$hasAdPermissionCmdlet = [bool](Get-Command -Name Get-ADPermission -ErrorAction SilentlyContinue)
if (-not $hasRecipientPermissionCmdlet -and -not $hasAdPermissionCmdlet) {
    Write-Status -Message 'Neither Get-RecipientPermission nor Get-ADPermission is available. SendAs export will be blank.' -Level WARN
}

$scopeMode = 'Csv'
$resolvedServer = Get-TrimmedValue -Value $Server
$resolvedSearchBase = ''
$runWasTruncated = $false

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    $resolvedSearchBase = Get-TrimmedValue -Value $SearchBase
    Write-Status -Message "DiscoverAll enabled for Exchange on-prem mailbox delegations. SearchBase='$resolvedSearchBase'." -Level WARN
    $rows = @([PSCustomObject]@{
            MailboxIdentity = '*'
        })
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results = [System.Collections.Generic.List[object]]::new()
$recipientSummaryByKey = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

$resolveRecipientSummary = {
    param(
        [Parameter(Mandatory)]
        [string]$IdentityHint
    )

    $normalized = Normalize-TrusteeKey -Value $IdentityHint
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return [PSCustomObject]@{
            TrusteeIdentity           = ''
            TrusteePrimarySmtpAddress = ''
            TrusteeRecipientType      = ''
        }
    }

    if ($recipientSummaryByKey.ContainsKey($normalized)) {
        return $recipientSummaryByKey[$normalized]
    }

    $summary = $null
    try {
        $recipient = Invoke-WithRetry -OperationName "Lookup recipient $IdentityHint" -ScriptBlock {
            Get-Recipient -Identity $IdentityHint -ErrorAction Stop
        }

        $summary = [PSCustomObject]@{
            TrusteeIdentity           = Get-TrimmedValue -Value $recipient.Identity
            TrusteePrimarySmtpAddress = Get-TrimmedValue -Value $recipient.PrimarySmtpAddress
            TrusteeRecipientType      = Get-TrimmedValue -Value $recipient.RecipientType
        }
    }
    catch {
        $summary = [PSCustomObject]@{
            TrusteeIdentity           = $IdentityHint
            TrusteePrimarySmtpAddress = ''
            TrusteeRecipientType      = ''
        }
    }

    $recipientSummaryByKey[$normalized] = $summary
    return $summary
}

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = Get-TrimmedValue -Value $row.MailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
            throw 'MailboxIdentity is required. Use * to inventory delegations for all user/shared mailboxes.'
        }

        $effectiveSearchBase = if ($scopeMode -eq 'DiscoverAll') { $resolvedSearchBase } else { '' }
        $mailboxes = @(Invoke-WithRetry -OperationName "Load mailboxes for $mailboxIdentity" -ScriptBlock {
            Resolve-MailboxesByScope -Identity $mailboxIdentity -SearchBase $effectiveSearchBase -Server $resolvedServer
        })

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $mailboxes.Count -gt $MaxObjects) {
            $mailboxes = @($mailboxes | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($mailboxes.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'GetExchangeMailboxDelegation' -Status 'NotFound' -Message 'No matching mailboxes were found.' -Data ([ordered]@{
                        MailboxIdentity             = $mailboxIdentity
                        MailboxRecipientTypeDetails = ''
                        TrusteeIdentity             = ''
                        TrusteePrimarySmtpAddress   = ''
                        TrusteeRecipientType        = ''
                        FullAccess                  = ''
                        ReadOnly                    = ''
                        SendAs                      = ''
                        SendOnBehalf                = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $mailboxIdentityResolved = Get-TrimmedValue -Value $mailbox.Identity
            $permissionMap = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

            $ensureEntry = {
                param(
                    [Parameter(Mandatory)]
                    [string]$TrusteeHint
                )

                $summary = & $resolveRecipientSummary -IdentityHint $TrusteeHint
                $key = Normalize-TrusteeKey -Value $summary.TrusteeIdentity
                if ([string]::IsNullOrWhiteSpace($key)) {
                    $key = Normalize-TrusteeKey -Value $TrusteeHint
                }
                if ([string]::IsNullOrWhiteSpace($key)) {
                    return $null
                }

                if ($permissionMap.ContainsKey($key)) {
                    return $permissionMap[$key]
                }

                $entry = [PSCustomObject]@{
                    TrusteeIdentity           = $summary.TrusteeIdentity
                    TrusteePrimarySmtpAddress = $summary.TrusteePrimarySmtpAddress
                    TrusteeRecipientType      = $summary.TrusteeRecipientType
                    FullAccess                = $false
                    ReadOnly                  = $false
                    SendAs                    = $false
                    SendOnBehalf              = $false
                }

                $permissionMap[$key] = $entry
                return $entry
            }

            $mailboxPermissions = @(Invoke-WithRetry -OperationName "Load mailbox permissions $mailboxIdentityResolved" -ScriptBlock {
                Get-MailboxPermission -Identity $mailbox.Identity -ErrorAction Stop
            })

            foreach ($permission in $mailboxPermissions) {
                if ($permission.Deny -eq $true) { continue }
                if ($permission.IsInherited -eq $true) { continue }

                $trustee = Get-TrimmedValue -Value $permission.User
                if ([string]::IsNullOrWhiteSpace($trustee)) { continue }
                if ($trustee.Equals('NT AUTHORITY\\SELF', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                if ($trustee -match '^S-1-5-') { continue }

                $entry = & $ensureEntry -TrusteeHint $trustee
                if ($null -eq $entry) { continue }

                $accessRights = @($permission.AccessRights | ForEach-Object { Get-TrimmedValue -Value $_ })
                if ($accessRights -contains 'FullAccess') {
                    $entry.FullAccess = $true
                }
                if ($accessRights -contains 'ReadPermission') {
                    $entry.ReadOnly = $true
                }
            }

            if ($hasRecipientPermissionCmdlet) {
                $recipientPermissions = @(Invoke-WithRetry -OperationName "Load recipient permissions $mailboxIdentityResolved" -ScriptBlock {
                    Get-RecipientPermission -Identity $mailbox.Identity -ErrorAction SilentlyContinue
                })

                foreach ($permission in $recipientPermissions) {
                    if ($permission.Deny -eq $true) { continue }

                    $accessRights = @($permission.AccessRights | ForEach-Object { Get-TrimmedValue -Value $_ })
                    if ($accessRights -notcontains 'SendAs') { continue }

                    $trustee = Get-TrimmedValue -Value $permission.Trustee
                    if ([string]::IsNullOrWhiteSpace($trustee)) { continue }

                    $entry = & $ensureEntry -TrusteeHint $trustee
                    if ($null -eq $entry) { continue }

                    $entry.SendAs = $true
                }
            }
            elseif ($hasAdPermissionCmdlet) {
                $adPermissions = @(Invoke-WithRetry -OperationName "Load AD permissions $mailboxIdentityResolved" -ScriptBlock {
                    Get-ADPermission -Identity $mailbox.Identity -ErrorAction SilentlyContinue
                })

                foreach ($permission in $adPermissions) {
                    if ($permission.Deny -eq $true) { continue }
                    if ($permission.IsInherited -eq $true) { continue }

                    $extendedRights = @($permission.ExtendedRights | ForEach-Object { Get-TrimmedValue -Value $_ })
                    if ($extendedRights -notcontains 'Send As') { continue }

                    $trustee = Get-TrimmedValue -Value $permission.User
                    if ([string]::IsNullOrWhiteSpace($trustee)) { continue }
                    if ($trustee -match '^S-1-5-') { continue }

                    $entry = & $ensureEntry -TrusteeHint $trustee
                    if ($null -eq $entry) { continue }

                    $entry.SendAs = $true
                }
            }

            foreach ($delegate in @($mailbox.GrantSendOnBehalfTo)) {
                $delegateHint = Get-DelegateHint -DelegateObject $delegate
                if ([string]::IsNullOrWhiteSpace($delegateHint)) { continue }

                $entry = & $ensureEntry -TrusteeHint $delegateHint
                if ($null -eq $entry) { continue }

                $entry.SendOnBehalf = $true
            }

            if ($permissionMap.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityResolved -Action 'GetExchangeMailboxDelegation' -Status 'Completed' -Message 'No explicit delegated permissions found for mailbox.' -Data ([ordered]@{
                            MailboxIdentity             = $mailboxIdentityResolved
                            MailboxRecipientTypeDetails = Get-TrimmedValue -Value $mailbox.RecipientTypeDetails
                            TrusteeIdentity             = ''
                            TrusteePrimarySmtpAddress   = ''
                            TrusteeRecipientType        = ''
                            FullAccess                  = ''
                            ReadOnly                    = ''
                            SendAs                      = ''
                            SendOnBehalf                = ''
                        })))
                continue
            }

            foreach ($entry in @($permissionMap.Values | Sort-Object -Property TrusteeIdentity)) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$mailboxIdentityResolved|$($entry.TrusteeIdentity)" -Action 'GetExchangeMailboxDelegation' -Status 'Completed' -Message 'Mailbox delegation row exported.' -Data ([ordered]@{
                            MailboxIdentity             = $mailboxIdentityResolved
                            MailboxRecipientTypeDetails = Get-TrimmedValue -Value $mailbox.RecipientTypeDetails
                            TrusteeIdentity             = $entry.TrusteeIdentity
                            TrusteePrimarySmtpAddress   = $entry.TrusteePrimarySmtpAddress
                            TrusteeRecipientType        = $entry.TrusteeRecipientType
                            FullAccess                  = [string]$entry.FullAccess
                            ReadOnly                    = [string]$entry.ReadOnly
                            SendAs                      = [string]$entry.SendAs
                            SendOnBehalf                = [string]$entry.SendOnBehalf
                        })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'GetExchangeMailboxDelegation' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    MailboxIdentity             = $mailboxIdentity
                    MailboxRecipientTypeDetails = ''
                    TrusteeIdentity             = ''
                    TrusteePrimarySmtpAddress   = ''
                    TrusteeRecipientType        = ''
                    FullAccess                  = ''
                    ReadOnly                    = ''
                    SendAs                      = ''
                    SendOnBehalf                = ''
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
Write-Status -Message 'Exchange on-prem mailbox delegation inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}

