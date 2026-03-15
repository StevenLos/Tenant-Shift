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

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\\..\\Standalone_OutputCsvPath') -ChildPath ("Results_SA-M0221-Set-ExchangeOnPremMailboxFolderPermissions_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function ConvertTo-CanonicalStringSet {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$Values
    )

    if ($null -eq $Values) {
        return ,@()
    }

    return ,@(
        @($Values) |
            ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
}

function Test-SameStringSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Left,

        [Parameter(Mandatory)]
        [string[]]$Right
    )

    $leftOnly = @($Left | Where-Object { $_ -notin $Right })
    $rightOnly = @($Right | Where-Object { $_ -notin $Left })
    return ($leftOnly.Count -eq 0 -and $rightOnly.Count -eq 0)
}

$requiredHeaders = @(
    'MailboxIdentity',
    'TrusteeIdentity',
    'FolderPath',
    'PermissionAction',
    'AccessRights',
    'CalendarDelegate',
    'CalendarCanViewPrivateItems',
    'SendNotificationToUser'
)

Write-Status -Message 'Starting Exchange on-prem mailbox folder permission script.'
Ensure-ExchangeOnPremConnection

$addMailboxFolderPermissionCommand = Get-Command -Name Add-MailboxFolderPermission -ErrorAction Stop
$setMailboxFolderPermissionCommand = Get-Command -Name Set-MailboxFolderPermission -ErrorAction Stop
$removeMailboxFolderPermissionCommand = Get-Command -Name Remove-MailboxFolderPermission -ErrorAction Stop
$addSupportsSharingFlags = $addMailboxFolderPermissionCommand.Parameters.ContainsKey('SharingPermissionFlags')
$setSupportsSharingFlags = $setMailboxFolderPermissionCommand.Parameters.ContainsKey('SharingPermissionFlags')
$addSupportsSendNotification = $addMailboxFolderPermissionCommand.Parameters.ContainsKey('SendNotificationToUser')
$setSupportsSendNotification = $setMailboxFolderPermissionCommand.Parameters.ContainsKey('SendNotificationToUser')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = Get-TrimmedValue -Value $row.MailboxIdentity
    $trusteeIdentity = Get-TrimmedValue -Value $row.TrusteeIdentity
    $folderPathRaw = Get-TrimmedValue -Value $row.FolderPath

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity) -or [string]::IsNullOrWhiteSpace($trusteeIdentity) -or [string]::IsNullOrWhiteSpace($folderPathRaw)) {
            throw 'MailboxIdentity, TrusteeIdentity, and FolderPath are required.'
        }

        $permissionActionRaw = Get-TrimmedValue -Value $row.PermissionAction
        $permissionAction = if ([string]::IsNullOrWhiteSpace($permissionActionRaw)) { 'Set' } else { $permissionActionRaw }
        if ($permissionAction -notin @('Add', 'Set', 'Remove')) {
            throw "PermissionAction '$permissionAction' is invalid. Use Add, Set, or Remove."
        }

        $accessRights = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.AccessRights)
        if ($permissionAction -ne 'Remove' -and $accessRights.Count -eq 0) {
            throw 'AccessRights is required for Add/Set actions and must contain at least one value.'
        }

        $calendarDelegate = ConvertTo-Bool -Value $row.CalendarDelegate -Default $false
        $calendarCanViewPrivateItems = ConvertTo-Bool -Value $row.CalendarCanViewPrivateItems -Default $false
        $sendNotificationToUser = ConvertTo-Bool -Value $row.SendNotificationToUser -Default $false

        if ($calendarCanViewPrivateItems -and -not $calendarDelegate) {
            throw 'CalendarCanViewPrivateItems requires CalendarDelegate = TRUE.'
        }

        $folderIdentity = if ($folderPathRaw -match '^[^:]+:\\') {
            $folderPathRaw
        }
        else {
            "${mailboxIdentity}:\\$folderPathRaw"
        }

        $folderPathForCheck = if ($folderPathRaw -match '^[^:]+:\\') {
            $folderPathRaw.Split(':\\', 2)[1]
        }
        else {
            $folderPathRaw
        }

        $folderSegments = @($folderPathForCheck -split '[\\/]' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $folderLeaf = ''
        if ($folderSegments.Count -gt 0) {
            $folderLeaf = Get-TrimmedValue -Value $folderSegments[-1]
        }
        $isCalendarFolder = $folderLeaf.Equals('Calendar', [System.StringComparison]::OrdinalIgnoreCase)

        if (($calendarDelegate -or $calendarCanViewPrivateItems -or $sendNotificationToUser) -and -not $isCalendarFolder) {
            throw 'CalendarDelegate, CalendarCanViewPrivateItems, and SendNotificationToUser are only supported when FolderPath targets Calendar.'
        }

        if ($isCalendarFolder -and ($calendarDelegate -or $calendarCanViewPrivateItems) -and (-not $addSupportsSharingFlags -or -not $setSupportsSharingFlags)) {
            throw 'This Exchange on-prem cmdlet session does not support -SharingPermissionFlags required for calendar delegate configuration.'
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $mailboxIdentity -ErrorAction SilentlyContinue
        }
        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|$folderPathRaw" -Action 'SetMailboxFolderPermission' -Status 'NotFound' -Message 'Mailbox not found.'))
            $rowNumber++
            continue
        }

        $trustee = Invoke-WithRetry -OperationName "Lookup trustee $trusteeIdentity" -ScriptBlock {
            Get-Recipient -Identity $trusteeIdentity -ErrorAction SilentlyContinue
        }
        if (-not $trustee) {
            throw "Trustee '$trusteeIdentity' was not found."
        }

        $existingPermissions = @(Invoke-WithRetry -OperationName "Check folder permission $folderIdentity -> $trusteeIdentity" -ScriptBlock {
            Get-MailboxFolderPermission -Identity $folderIdentity -User $trustee.Identity -ErrorAction SilentlyContinue
        })
        $existingPermission = if ($existingPermissions.Count -gt 0) { $existingPermissions[0] } else { $null }

        if ($permissionAction -eq 'Remove') {
            if (-not $existingPermission) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Remove" -Action 'SetMailboxFolderPermission' -Status 'Skipped' -Message 'Folder permission does not exist.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$folderIdentity -> $trusteeIdentity", 'Remove mailbox folder permission')) {
                Invoke-WithRetry -OperationName "Remove folder permission $folderIdentity -> $trusteeIdentity" -ScriptBlock {
                    Remove-MailboxFolderPermission -Identity $folderIdentity -User $trustee.Identity -Confirm:$false -ErrorAction Stop | Out-Null
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Remove" -Action 'SetMailboxFolderPermission' -Status 'Removed' -Message 'Folder permission removed successfully.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Remove" -Action 'SetMailboxFolderPermission' -Status 'WhatIf' -Message 'Folder permission removal skipped due to WhatIf.'))
            }

            $rowNumber++
            continue
        }

        $desiredAccessRightsCanonical = ConvertTo-CanonicalStringSet -Values $accessRights
        $desiredSharingFlagsCanonical = @()
        if ($isCalendarFolder -and $setSupportsSharingFlags) {
            if ($calendarDelegate) {
                $flags = [System.Collections.Generic.List[string]]::new()
                $flags.Add('Delegate')
                if ($calendarCanViewPrivateItems) {
                    $flags.Add('CanViewPrivateItems')
                }

                $desiredSharingFlagsCanonical = ConvertTo-CanonicalStringSet -Values $flags.ToArray()
            }
            else {
                $desiredSharingFlagsCanonical = @('none')
            }
        }

        if ($existingPermission) {
            if ($permissionAction -eq 'Add') {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Add" -Action 'SetMailboxFolderPermission' -Status 'Skipped' -Message 'Folder permission already exists. Use PermissionAction=Set to update rights.'))
                $rowNumber++
                continue
            }

            $currentAccessRightsCanonical = ConvertTo-CanonicalStringSet -Values @($existingPermission.AccessRights)
            $rightsMatch = Test-SameStringSet -Left $currentAccessRightsCanonical -Right $desiredAccessRightsCanonical

            $flagsMatch = $true
            if ($isCalendarFolder -and $setSupportsSharingFlags) {
                $currentFlagsCanonical = ConvertTo-CanonicalStringSet -Values @($existingPermission.SharingPermissionFlags)
                if ($currentFlagsCanonical.Count -eq 0) {
                    $currentFlagsCanonical = @('none')
                }

                $flagsMatch = Test-SameStringSet -Left $currentFlagsCanonical -Right $desiredSharingFlagsCanonical
            }

            if ($rightsMatch -and $flagsMatch) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Set" -Action 'SetMailboxFolderPermission' -Status 'Skipped' -Message 'Folder permissions already match requested values.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$folderIdentity -> $trusteeIdentity", 'Set mailbox folder permission')) {
                $params = @{
                    Identity     = $folderIdentity
                    User         = $trustee.Identity
                    AccessRights = $accessRights
                    ErrorAction  = 'Stop'
                }

                if ($isCalendarFolder -and $setSupportsSharingFlags) {
                    if ($desiredSharingFlagsCanonical.Count -eq 1 -and $desiredSharingFlagsCanonical[0] -eq 'none') {
                        $params.SharingPermissionFlags = @('None')
                    }
                    else {
                        $params.SharingPermissionFlags = $desiredSharingFlagsCanonical
                    }
                }

                if ($sendNotificationToUser) {
                    if ($setSupportsSendNotification) {
                        $params.SendNotificationToUser = $true
                    }
                    else {
                        Write-Status -Message "Set-MailboxFolderPermission in this session does not support -SendNotificationToUser. Notification request for '$folderIdentity' was ignored." -Level WARN
                    }
                }

                Invoke-WithRetry -OperationName "Set folder permission $folderIdentity -> $trusteeIdentity" -ScriptBlock {
                    Set-MailboxFolderPermission @params | Out-Null
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Set" -Action 'SetMailboxFolderPermission' -Status 'Updated' -Message 'Folder permission updated successfully.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Set" -Action 'SetMailboxFolderPermission' -Status 'WhatIf' -Message 'Folder permission update skipped due to WhatIf.'))
            }

            $rowNumber++
            continue
        }

        if ($permissionAction -eq 'Set') {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Set" -Action 'SetMailboxFolderPermission' -Status 'Skipped' -Message 'Folder permission does not exist. Use PermissionAction=Add to create it.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess("$folderIdentity -> $trusteeIdentity", 'Add mailbox folder permission')) {
            $params = @{
                Identity     = $folderIdentity
                User         = $trustee.Identity
                AccessRights = $accessRights
                ErrorAction  = 'Stop'
            }

            if ($isCalendarFolder -and $addSupportsSharingFlags) {
                if ($desiredSharingFlagsCanonical.Count -eq 1 -and $desiredSharingFlagsCanonical[0] -eq 'none') {
                    $params.SharingPermissionFlags = @('None')
                }
                else {
                    $params.SharingPermissionFlags = $desiredSharingFlagsCanonical
                }
            }

            if ($sendNotificationToUser) {
                if ($addSupportsSendNotification) {
                    $params.SendNotificationToUser = $true
                }
                else {
                    Write-Status -Message "Add-MailboxFolderPermission in this session does not support -SendNotificationToUser. Notification request for '$folderIdentity' was ignored." -Level WARN
                }
            }

            Invoke-WithRetry -OperationName "Add folder permission $folderIdentity -> $trusteeIdentity" -ScriptBlock {
                Add-MailboxFolderPermission @params | Out-Null
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Add" -Action 'SetMailboxFolderPermission' -Status 'Added' -Message 'Folder permission added successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Add" -Action 'SetMailboxFolderPermission' -Status 'WhatIf' -Message 'Folder permission add skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity|$trusteeIdentity|$folderPathRaw) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|$folderPathRaw" -Action 'SetMailboxFolderPermission' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem mailbox folder permission script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}

