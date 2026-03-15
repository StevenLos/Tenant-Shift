<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260302-191500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
ActiveDirectory

.MODULEVERSIONPOLICY
Best-effort PSGallery check (offline-safe)
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\..\Standalone_OutputCsvPath') -ChildPath ("Results_SA-M0009-Move-ActiveDirectoryObjects_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Escape-LdapFilterValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    $builder = [System.Text.StringBuilder]::new()
    foreach ($char in $Value.ToCharArray()) {
        switch ($char) {
            '\\' { [void]$builder.Append('\\5c') }
            '*' { [void]$builder.Append('\\2a') }
            '(' { [void]$builder.Append('\\28') }
            ')' { [void]$builder.Append('\\29') }
            ([char]0) { [void]$builder.Append('\\00') }
            default { [void]$builder.Append($char) }
        }
    }

    return $builder.ToString()
}

function Get-ObjectTypeLdapClause {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ObjectType
    )

    switch ($ObjectType.Trim().ToLowerInvariant()) {
        'any' { return '' }
        'user' { return '(&(objectClass=user)(!(objectClass=computer)))' }
        'group' { return '(objectClass=group)' }
        'contact' { return '(objectClass=contact)' }
        'organizationalunit' { return '(objectClass=organizationalUnit)' }
        'ou' { return '(objectClass=organizationalUnit)' }
        'computer' { return '(objectClass=computer)' }
        default {
            throw "ObjectType '$ObjectType' is invalid. Use Any, User, Group, Contact, OrganizationalUnit, OU, or Computer."
        }
    }
}

function Test-ObjectTypeMatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$AdObject,

        [Parameter(Mandatory)]
        [string]$ObjectType
    )

    $normalizedType = $ObjectType.Trim().ToLowerInvariant()
    if ($normalizedType -eq 'any') {
        return $true
    }

    $classes = @($AdObject.ObjectClass | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() })
    switch ($normalizedType) {
        'user' { return (($classes -contains 'user') -and (-not ($classes -contains 'computer'))) }
        'group' { return ($classes -contains 'group') }
        'contact' { return ($classes -contains 'contact') }
        'organizationalunit' { return ($classes -contains 'organizationalunit') }
        'ou' { return ($classes -contains 'organizationalunit') }
        'computer' { return ($classes -contains 'computer') }
    }

    return $false
}

function Resolve-TargetAdObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ObjectType,

        [Parameter(Mandatory)]
        [string]$IdentityType,

        [Parameter(Mandatory)]
        [string]$IdentityValue
    )

    $normalizedIdentityType = $IdentityType.Trim().ToLowerInvariant()

    if ($normalizedIdentityType -eq 'distinguishedname') {
        $candidate = Get-ADObject -Identity $IdentityValue -Properties * -ErrorAction SilentlyContinue
        if (-not $candidate) {
            return $null
        }

        if (-not (Test-ObjectTypeMatch -AdObject $candidate -ObjectType $ObjectType)) {
            throw "Resolved object does not match ObjectType '$ObjectType'."
        }

        return $candidate
    }

    if ($normalizedIdentityType -eq 'objectguid') {
        $guid = [guid]$IdentityValue
        $candidate = Get-ADObject -Identity $guid -Properties * -ErrorAction SilentlyContinue
        if (-not $candidate) {
            return $null
        }

        if (-not (Test-ObjectTypeMatch -AdObject $candidate -ObjectType $ObjectType)) {
            throw "Resolved object does not match ObjectType '$ObjectType'."
        }

        return $candidate
    }

    $escapedIdentityValue = Escape-LdapFilterValue -Value $IdentityValue
    $identityClause = switch ($normalizedIdentityType) {
        'samaccountname' { "(sAMAccountName=$escapedIdentityValue)" }
        'userprincipalname' { "(userPrincipalName=$escapedIdentityValue)" }
        'name' { "(name=$escapedIdentityValue)" }
        default {
            throw "IdentityType '$IdentityType' is invalid. Use SamAccountName, UserPrincipalName, Name, DistinguishedName, or ObjectGuid."
        }
    }

    $typeClause = Get-ObjectTypeLdapClause -ObjectType $ObjectType
    $ldapFilter = if ([string]::IsNullOrWhiteSpace($typeClause)) { $identityClause } else { "(&$typeClause$identityClause)" }

    $matches = @(Get-ADObject -LDAPFilter $ldapFilter -Properties * -ErrorAction SilentlyContinue)
    if ($matches.Count -eq 0) {
        return $null
    }

    if ($matches.Count -gt 1) {
        throw "Identity '$IdentityType=$IdentityValue' resolved to $($matches.Count) objects. Use DistinguishedName or ObjectGuid for an unambiguous target."
    }

    return $matches[0]
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'ObjectType',
    'IdentityType',
    'IdentityValue',
    'TargetPath',
    'NewName',
    'ProtectionEnabled'
)

Write-Status -Message 'Starting Active Directory object move script.'
Ensure-ActiveDirectoryConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $objectType = Get-TrimmedValue -Value $row.ObjectType
    if ([string]::IsNullOrWhiteSpace($objectType)) {
        $objectType = 'Any'
    }

    $identityType = Get-TrimmedValue -Value $row.IdentityType
    $identityValue = Get-TrimmedValue -Value $row.IdentityValue
    $primaryKey = "$objectType|${identityType}:$identityValue"

    try {
        if ([string]::IsNullOrWhiteSpace($identityType) -or [string]::IsNullOrWhiteSpace($identityValue)) {
            throw 'IdentityType and IdentityValue are required.'
        }

        $targetObject = Invoke-WithRetry -OperationName "Resolve AD object $primaryKey" -ScriptBlock {
            Resolve-TargetAdObject -ObjectType $objectType -IdentityType $identityType -IdentityValue $identityValue
        }

        if (-not $targetObject) {
            throw 'Target object was not found.'
        }

        $resolvedKey = Get-TrimmedValue -Value $targetObject.DistinguishedName
        if ([string]::IsNullOrWhiteSpace($resolvedKey)) {
            $resolvedKey = "ObjectGuid:$($targetObject.ObjectGuid)"
        }

        $messages = [System.Collections.Generic.List[string]]::new()
        $changeCount = 0

        $newName = Get-TrimmedValue -Value $row.NewName
        if (-not [string]::IsNullOrWhiteSpace($newName) -and $newName -ne (Get-TrimmedValue -Value $targetObject.Name)) {
            $changeCount++

            if ($PSCmdlet.ShouldProcess($resolvedKey, "Rename AD object to '$newName'")) {
                Invoke-WithRetry -OperationName "Rename AD object $resolvedKey" -ScriptBlock {
                    Rename-ADObject -Identity $targetObject.ObjectGuid -NewName $newName -ErrorAction Stop
                }

                $messages.Add("Object renamed to '$newName'.")
            }
            else {
                $messages.Add('Rename skipped due to WhatIf.')
            }
        }

        $targetPath = Get-TrimmedValue -Value $row.TargetPath
        if (-not [string]::IsNullOrWhiteSpace($targetPath)) {
            $currentObject = Invoke-WithRetry -OperationName "Reload AD object $resolvedKey" -ScriptBlock {
                Get-ADObject -Identity $targetObject.ObjectGuid -Properties DistinguishedName -ErrorAction Stop
            }

            $currentParentPath = ($currentObject.DistinguishedName -split ',', 2)[1]
            if ($targetPath -ieq $currentParentPath) {
                $messages.Add('Object already in requested OU path.')
            }
            else {
                $changeCount++
                if ($PSCmdlet.ShouldProcess($resolvedKey, "Move AD object to '$targetPath'")) {
                    Invoke-WithRetry -OperationName "Move AD object $resolvedKey" -ScriptBlock {
                        Move-ADObject -Identity $targetObject.ObjectGuid -TargetPath $targetPath -ErrorAction Stop
                    }

                    $messages.Add("Object moved to '$targetPath'.")
                }
                else {
                    $messages.Add('Move skipped due to WhatIf.')
                }
            }
        }

        $protectionEnabled = Get-NullableBool -Value $row.ProtectionEnabled
        if ($null -ne $protectionEnabled) {
            $changeCount++
            if ($PSCmdlet.ShouldProcess($resolvedKey, "Set ProtectedFromAccidentalDeletion = $protectionEnabled")) {
                Invoke-WithRetry -OperationName "Set AD object protection $resolvedKey" -ScriptBlock {
                    Set-ADObject -Identity $targetObject.ObjectGuid -ProtectedFromAccidentalDeletion $protectionEnabled -ErrorAction Stop
                }

                $messages.Add("ProtectedFromAccidentalDeletion set to '$protectionEnabled'.")
            }
            else {
                $messages.Add('Protection change skipped due to WhatIf.')
            }
        }

        if ($changeCount -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'MoveActiveDirectoryObject' -Status 'Skipped' -Message 'No changes were requested for this row.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resolvedKey -Action 'MoveActiveDirectoryObject' -Status 'Completed' -Message ($messages -join ' ')))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'MoveActiveDirectoryObject' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory object move script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}


