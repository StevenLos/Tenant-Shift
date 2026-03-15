<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-220000

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\\..\\Standalone_OutputCsvPath') -ChildPath ("Results_SA-M0220-Set-ExchangeOnPremMailboxDelegations_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-RecipientKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Recipient
    )

    $primary = Get-TrimmedValue -Value $Recipient.PrimarySmtpAddress
    if (-not [string]::IsNullOrWhiteSpace($primary)) {
        return $primary.ToLowerInvariant()
    }

    $identity = Get-TrimmedValue -Value $Recipient.Identity
    if (-not [string]::IsNullOrWhiteSpace($identity)) {
        return $identity.ToLowerInvariant()
    }

    return ''
}

$requiredHeaders = @(
    'MailboxIdentity',
    'TrusteeIdentity',
    'PermissionType',
    'PermissionAction',
    'AutoMapping'
)

Write-Status -Message 'Starting Exchange on-prem mailbox delegation script.'
Ensure-ExchangeOnPremConnection

$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop
$supportsGrantSendOnBehalf = $setMailboxCommand.Parameters.ContainsKey('GrantSendOnBehalfTo')
$hasRecipientPermissionCmdlets = (Get-Command -Name Get-RecipientPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Add-RecipientPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Remove-RecipientPermission -ErrorAction SilentlyContinue)
$hasAdPermissionCmdlets = (Get-Command -Name Get-ADPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Add-ADPermission -ErrorAction SilentlyContinue) -and (Get-Command -Name Remove-ADPermission -ErrorAction SilentlyContinue)

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = Get-TrimmedValue -Value $row.MailboxIdentity
    $trusteeIdentity = Get-TrimmedValue -Value $row.TrusteeIdentity
    $permissionTypeRaw = Get-TrimmedValue -Value $row.PermissionType
    $permissionActionRaw = Get-TrimmedValue -Value $row.PermissionAction

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity) -or [string]::IsNullOrWhiteSpace($trusteeIdentity)) {
            throw 'MailboxIdentity and TrusteeIdentity are required.'
        }

        $permissionType = if ([string]::IsNullOrWhiteSpace($permissionTypeRaw)) { 'FullAccess' } else { $permissionTypeRaw }
        if ($permissionType -notin @('FullAccess', 'SendAs', 'SendOnBehalf')) {
            throw "PermissionType '$permissionType' is invalid. Use FullAccess, SendAs, or SendOnBehalf."
        }

        $permissionAction = if ([string]::IsNullOrWhiteSpace($permissionActionRaw)) { 'Add' } else { $permissionActionRaw }
        if ($permissionAction -notin @('Add', 'Remove')) {
            throw "PermissionAction '$permissionAction' is invalid. Use Add or Remove."
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $mailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|$permissionType" -Action 'SetMailboxDelegation' -Status 'NotFound' -Message 'Mailbox not found.'))
            $rowNumber++
            continue
        }

        $trusteeRecipient = Invoke-WithRetry -OperationName "Lookup trustee recipient $trusteeIdentity" -ScriptBlock {
            Get-Recipient -Identity $trusteeIdentity -ErrorAction SilentlyContinue
        }
        if (-not $trusteeRecipient) {
            throw "Trustee '$trusteeIdentity' was not found."
        }

        $trusteeKey = Get-RecipientKey -Recipient $trusteeRecipient

        if ($permissionType -eq 'FullAccess') {
            $autoMappingRaw = Get-TrimmedValue -Value $row.AutoMapping
            $autoMapping = if ([string]::IsNullOrWhiteSpace($autoMappingRaw)) { $true } else { ConvertTo-Bool -Value $autoMappingRaw }

            $existingPermissions = @(Invoke-WithRetry -OperationName "Load FullAccess delegations for $mailboxIdentity" -ScriptBlock {
                Get-MailboxPermission -Identity $mailbox.Identity -ErrorAction Stop
            })

            $existing = $false
            foreach ($perm in $existingPermissions) {
                if ($perm.IsInherited -or $perm.Deny) {
                    continue
                }

                $permUser = Get-TrimmedValue -Value $perm.User
                if ($permUser.Equals((Get-TrimmedValue -Value $trusteeRecipient.Identity), [System.StringComparison]::OrdinalIgnoreCase) -or $permUser.Equals((Get-TrimmedValue -Value $trusteeRecipient.Name), [System.StringComparison]::OrdinalIgnoreCase) -or $permUser.Equals((Get-TrimmedValue -Value $trusteeRecipient.PrimarySmtpAddress), [System.StringComparison]::OrdinalIgnoreCase)) {
                    if (@($perm.AccessRights) -contains 'FullAccess') {
                        $existing = $true
                        break
                    }
                }
            }

            if ($permissionAction -eq 'Add') {
                if ($existing) {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetMailboxDelegation' -Status 'Skipped' -Message 'FullAccess delegation already exists.'))
                    $rowNumber++
                    continue
                }

                if ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeIdentity", 'Add FullAccess delegation')) {
                    Invoke-WithRetry -OperationName "Add FullAccess delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                        Add-MailboxPermission -Identity $mailbox.Identity -User $trusteeRecipient.Identity -AccessRights FullAccess -InheritanceType All -AutoMapping:$autoMapping -ErrorAction Stop | Out-Null
                    }

                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetMailboxDelegation' -Status 'Added' -Message 'FullAccess delegation added.'))
                }
                else {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetMailboxDelegation' -Status 'WhatIf' -Message 'Delegation change skipped due to WhatIf.'))
                }
            }
            else {
                if (-not $existing) {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetMailboxDelegation' -Status 'Skipped' -Message 'FullAccess delegation does not exist.'))
                    $rowNumber++
                    continue
                }

                if ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeIdentity", 'Remove FullAccess delegation')) {
                    Invoke-WithRetry -OperationName "Remove FullAccess delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                        Remove-MailboxPermission -Identity $mailbox.Identity -User $trusteeRecipient.Identity -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                    }

                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetMailboxDelegation' -Status 'Removed' -Message 'FullAccess delegation removed.'))
                }
                else {
                    $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|FullAccess" -Action 'SetMailboxDelegation' -Status 'WhatIf' -Message 'Delegation change skipped due to WhatIf.'))
                }
            }

            $rowNumber++
            continue
        }

        if ($permissionType -eq 'SendAs') {
            if (-not $hasRecipientPermissionCmdlets -and -not $hasAdPermissionCmdlets) {
                throw 'No supported cmdlet set for SendAs delegations was found (Get/Add/Remove-RecipientPermission or Get/Add/Remove-ADPermission).'
            }

            if ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeIdentity", "$permissionAction SendAs delegation")) {
                if ($hasRecipientPermissionCmdlets) {
                    try {
                        if ($permissionAction -eq 'Add') {
                            Invoke-WithRetry -OperationName "Add SendAs delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                                Add-RecipientPermission -Identity $mailbox.Identity -Trustee $trusteeRecipient.Identity -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                            }
                            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'Added' -Message 'SendAs delegation add attempted via RecipientPermission cmdlets.'))
                        }
                        else {
                            Invoke-WithRetry -OperationName "Remove SendAs delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                                Remove-RecipientPermission -Identity $mailbox.Identity -Trustee $trusteeRecipient.Identity -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                            }
                            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'Removed' -Message 'SendAs delegation remove attempted via RecipientPermission cmdlets.'))
                        }
                    }
                    catch {
                        $messageLower = ([string]$_.Exception.Message).ToLowerInvariant()
                        if ($permissionAction -eq 'Add' -and $messageLower -match 'already|exists|duplicate') {
                            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'Skipped' -Message 'SendAs delegation already exists.'))
                        }
                        elseif ($permissionAction -eq 'Remove' -and $messageLower -match 'cannot find|not found|doesn''t exist') {
                            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'Skipped' -Message 'SendAs delegation does not exist.'))
                        }
                        else {
                            throw
                        }
                    }
                }
                else {
                    if ($permissionAction -eq 'Add') {
                        Invoke-WithRetry -OperationName "Add SendAs AD delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                            Add-ADPermission -Identity $mailbox.Identity -User $trusteeRecipient.Identity -ExtendedRights 'Send As' -Confirm:$false -ErrorAction Stop | Out-Null
                        }
                        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'Added' -Message 'SendAs delegation add attempted via ADPermission cmdlets.'))
                    }
                    else {
                        Invoke-WithRetry -OperationName "Remove SendAs AD delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                            Remove-ADPermission -Identity $mailbox.Identity -User $trusteeRecipient.Identity -ExtendedRights 'Send As' -Confirm:$false -ErrorAction Stop | Out-Null
                        }
                        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'Removed' -Message 'SendAs delegation remove attempted via ADPermission cmdlets.'))
                    }
                }
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendAs" -Action 'SetMailboxDelegation' -Status 'WhatIf' -Message 'Delegation change skipped due to WhatIf.'))
            }

            $rowNumber++
            continue
        }

        if (-not $supportsGrantSendOnBehalf) {
            throw 'Set-Mailbox -GrantSendOnBehalfTo is not available in this session.'
        }

        $currentDelegates = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($delegate in @($mailbox.GrantSendOnBehalfTo)) {
            $delegateRecipient = Invoke-WithRetry -OperationName "Resolve current SendOnBehalf delegate $delegate" -ScriptBlock {
                Get-Recipient -Identity $delegate -ErrorAction SilentlyContinue
            }

            if ($delegateRecipient) {
                $key = Get-RecipientKey -Recipient $delegateRecipient
                if (-not [string]::IsNullOrWhiteSpace($key)) {
                    $null = $currentDelegates.Add($key)
                }
            }
            else {
                $raw = Get-TrimmedValue -Value $delegate
                if (-not [string]::IsNullOrWhiteSpace($raw)) {
                    $null = $currentDelegates.Add($raw.ToLowerInvariant())
                }
            }
        }

        if ($permissionAction -eq 'Add') {
            if ($currentDelegates.Contains($trusteeKey)) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetMailboxDelegation' -Status 'Skipped' -Message 'SendOnBehalf delegation already exists.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeIdentity", 'Add SendOnBehalf delegation')) {
                Invoke-WithRetry -OperationName "Add SendOnBehalf delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                    Set-Mailbox -Identity $mailbox.Identity -GrantSendOnBehalfTo @{ Add = $trusteeRecipient.Identity } -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetMailboxDelegation' -Status 'Added' -Message 'SendOnBehalf delegation added.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetMailboxDelegation' -Status 'WhatIf' -Message 'Delegation change skipped due to WhatIf.'))
            }
        }
        else {
            if (-not $currentDelegates.Contains($trusteeKey)) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetMailboxDelegation' -Status 'Skipped' -Message 'SendOnBehalf delegation does not exist.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$mailboxIdentity -> $trusteeIdentity", 'Remove SendOnBehalf delegation')) {
                Invoke-WithRetry -OperationName "Remove SendOnBehalf delegation $mailboxIdentity -> $trusteeIdentity" -ScriptBlock {
                    Set-Mailbox -Identity $mailbox.Identity -GrantSendOnBehalfTo @{ Remove = $trusteeRecipient.Identity } -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetMailboxDelegation' -Status 'Removed' -Message 'SendOnBehalf delegation removed.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|SendOnBehalf" -Action 'SetMailboxDelegation' -Status 'WhatIf' -Message 'Delegation change skipped due to WhatIf.'))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity|$trusteeIdentity|$permissionTypeRaw|$permissionActionRaw) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|$permissionTypeRaw|$permissionActionRaw" -Action 'SetMailboxDelegation' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem mailbox delegation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}

