<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260302-201500

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\..\Standalone_OutputCsvPath') -ChildPath ("Results_SA-M0008-Set-ActiveDirectoryDistributionGroupMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$expectedGroupCategory = 'Distribution'


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

function Resolve-TargetAdGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IdentityType,

        [Parameter(Mandatory)]
        [string]$IdentityValue
    )

    $normalizedType = $IdentityType.Trim().ToLowerInvariant()
    switch ($normalizedType) {
        'samaccountname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADGroup -Filter "SamAccountName -eq '$escaped'" -Properties GroupCategory,DistinguishedName,Name,SamAccountName,ObjectGuid -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'name' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADGroup -Filter "Name -eq '$escaped'" -Properties GroupCategory,DistinguishedName,Name,SamAccountName,ObjectGuid -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'distinguishedname' {
            return Get-ADGroup -Identity $IdentityValue -Properties GroupCategory,DistinguishedName,Name,SamAccountName,ObjectGuid -ErrorAction SilentlyContinue
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            return Get-ADGroup -Identity $guid -Properties GroupCategory,DistinguishedName,Name,SamAccountName,ObjectGuid -ErrorAction SilentlyContinue
        }
        default {
            throw "GroupIdentityType '$IdentityType' is invalid. Use SamAccountName, Name, DistinguishedName, or ObjectGuid."
        }
    }
}

function Resolve-MemberObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IdentityType,

        [Parameter(Mandatory)]
        [string]$IdentityValue
    )

    $normalizedType = $IdentityType.Trim().ToLowerInvariant()
    switch ($normalizedType) {
        'samaccountname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue

            $user = Get-ADUser -Filter "SamAccountName -eq '$escaped'" -Properties DistinguishedName,ObjectGuid,SamAccountName,UserPrincipalName,Name -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($user) { return $user }

            $group = Get-ADGroup -Filter "SamAccountName -eq '$escaped'" -Properties DistinguishedName,ObjectGuid,SamAccountName,Name -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($group) { return $group }

            $computer = Get-ADComputer -Filter "SamAccountName -eq '$escaped'" -Properties DistinguishedName,ObjectGuid,SamAccountName,Name -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($computer) { return $computer }

            return $null
        }
        'userprincipalname' {
            $escaped = Escape-AdFilterValue -Value $IdentityValue
            return Get-ADUser -Filter "UserPrincipalName -eq '$escaped'" -Properties DistinguishedName,ObjectGuid,SamAccountName,UserPrincipalName,Name -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        'distinguishedname' {
            return Get-ADObject -Identity $IdentityValue -Properties DistinguishedName,ObjectGuid,samAccountName,userPrincipalName,Name,ObjectClass -ErrorAction SilentlyContinue
        }
        'objectguid' {
            $guid = [guid]$IdentityValue
            return Get-ADObject -Identity $guid -Properties DistinguishedName,ObjectGuid,samAccountName,userPrincipalName,Name,ObjectClass -ErrorAction SilentlyContinue
        }
        default {
            throw "MemberIdentityType '$IdentityType' is invalid. Use SamAccountName, UserPrincipalName, DistinguishedName, or ObjectGuid."
        }
    }
}

$requiredHeaders = @(
    'Action',
    'Notes',
    'GroupIdentityType',
    'GroupIdentityValue',
    'MemberIdentityType',
    'MemberIdentityValue',
    'MemberAction'
)

Write-Status -Message 'Starting Active Directory distribution group membership update script.'
Ensure-ActiveDirectoryConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $groupIdentityType = Get-TrimmedValue -Value $row.GroupIdentityType
    $groupIdentityValue = Get-TrimmedValue -Value $row.GroupIdentityValue
    $memberIdentityType = Get-TrimmedValue -Value $row.MemberIdentityType
    $memberIdentityValue = Get-TrimmedValue -Value $row.MemberIdentityValue
    $memberAction = (Get-TrimmedValue -Value $row.MemberAction).ToLowerInvariant()

    $primaryKey = "${groupIdentityType}:$groupIdentityValue|${memberIdentityType}:$memberIdentityValue"

    try {
        if ([string]::IsNullOrWhiteSpace($groupIdentityType) -or [string]::IsNullOrWhiteSpace($groupIdentityValue)) {
            throw 'GroupIdentityType and GroupIdentityValue are required.'
        }

        if ([string]::IsNullOrWhiteSpace($memberIdentityType) -or [string]::IsNullOrWhiteSpace($memberIdentityValue)) {
            throw 'MemberIdentityType and MemberIdentityValue are required.'
        }

        if ($memberAction -notin @('add', 'remove')) {
            throw "MemberAction '$($row.MemberAction)' is invalid. Use Add or Remove."
        }

        $targetGroup = Invoke-WithRetry -OperationName "Resolve AD group $groupIdentityValue" -ScriptBlock {
            Resolve-TargetAdGroup -IdentityType $groupIdentityType -IdentityValue $groupIdentityValue
        }

        if (-not $targetGroup) {
            throw 'Target group was not found.'
        }

        if ((Get-TrimmedValue -Value $targetGroup.GroupCategory) -ne $expectedGroupCategory) {
            throw "Target group category '$($targetGroup.GroupCategory)' does not match expected '$expectedGroupCategory'."
        }

        $memberObject = Invoke-WithRetry -OperationName "Resolve AD member $memberIdentityValue" -ScriptBlock {
            Resolve-MemberObject -IdentityType $memberIdentityType -IdentityValue $memberIdentityValue
        }

        if (-not $memberObject) {
            throw 'Member object was not found.'
        }

        $memberDn = Get-TrimmedValue -Value $memberObject.DistinguishedName
        if ([string]::IsNullOrWhiteSpace($memberDn)) {
            throw 'Resolved member object does not contain DistinguishedName.'
        }

        $groupState = Invoke-WithRetry -OperationName "Load group members for $($targetGroup.SamAccountName)" -ScriptBlock {
            Get-ADGroup -Identity $targetGroup.ObjectGuid -Properties Member -ErrorAction Stop
        }

        $existingMemberDns = @($groupState.Member)
        $memberExists = $existingMemberDns -contains $memberDn

        if ($memberAction -eq 'add') {
            if ($memberExists) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'Skipped' -Message 'Member already present in group.'))
            }
            elseif ($PSCmdlet.ShouldProcess($targetGroup.SamAccountName, "Add member $memberDn")) {
                Invoke-WithRetry -OperationName "Add member to AD distribution group $($targetGroup.SamAccountName)" -ScriptBlock {
                    Add-ADGroupMember -Identity $targetGroup.ObjectGuid -Members $memberDn -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'Completed' -Message 'Member added to group.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'WhatIf' -Message 'Add skipped due to WhatIf.'))
            }
        }
        else {
            if (-not $memberExists) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'Skipped' -Message 'Member is not currently in group.'))
            }
            elseif ($PSCmdlet.ShouldProcess($targetGroup.SamAccountName, "Remove member $memberDn")) {
                Invoke-WithRetry -OperationName "Remove member from AD distribution group $($targetGroup.SamAccountName)" -ScriptBlock {
                    Remove-ADGroupMember -Identity $targetGroup.ObjectGuid -Members $memberDn -Confirm:$false -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'Completed' -Message 'Member removed from group.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'WhatIf' -Message 'Remove skipped due to WhatIf.'))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetActiveDirectoryDistributionGroupMember' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Active Directory distribution group membership update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}


