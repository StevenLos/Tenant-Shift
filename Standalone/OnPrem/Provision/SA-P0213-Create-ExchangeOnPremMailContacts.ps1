<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-000100

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..\\..\\Standalone_OutputCsvPath') -ChildPath ("Results_SA-P0213-Create-ExchangeOnPremMailContacts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function ConvertTo-NormalizedSmtpAddress {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    $trimmed = $Value.Trim()
    if ($trimmed.StartsWith('SMTP:', [System.StringComparison]::OrdinalIgnoreCase)) {
        $trimmed = $trimmed.Substring(5)
    }

    return $trimmed.ToLowerInvariant()
}

function Get-OptionalColumnValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [psobject]$Row,

        [Parameter(Mandatory)]
        [string]$ColumnName
    )

    $property = $Row.PSObject.Properties[$ColumnName]
    if ($null -eq $property) {
        return ''
    }

    return Get-TrimmedValue -Value $property.Value
}

$requiredHeaders = @(
    'Name',
    'ExternalEmailAddress',
    'Alias',
    'DisplayName',
    'FirstName',
    'LastName',
    'PrimarySmtpAddress',
    'HiddenFromAddressListsEnabled'
)

Write-Status -Message 'Starting Exchange on-prem mail contact creation script.'
Ensure-ExchangeOnPremConnection

$newMailContactCommand = Get-Command -Name New-MailContact -ErrorAction Stop
$supportsOrganizationalUnit = $newMailContactCommand.Parameters.ContainsKey('OrganizationalUnit')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

Write-Status -Message 'Loading existing mail contacts for idempotent checks.'
$existingContacts = @(Invoke-WithRetry -OperationName 'Load existing mail contacts' -ScriptBlock {
    Get-MailContact -ResultSize Unlimited -ErrorAction Stop
})

$contactsByExternalEmail = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$contactsByPrimarySmtp = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$contactsByAlias = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($contact in $existingContacts) {
    $externalKey = ConvertTo-NormalizedSmtpAddress -Value ([string]$contact.ExternalEmailAddress)
    if (-not [string]::IsNullOrWhiteSpace($externalKey) -and -not $contactsByExternalEmail.ContainsKey($externalKey)) {
        $contactsByExternalEmail[$externalKey] = $contact
    }

    $primaryKey = ConvertTo-NormalizedSmtpAddress -Value ([string]$contact.PrimarySmtpAddress)
    if (-not [string]::IsNullOrWhiteSpace($primaryKey) -and -not $contactsByPrimarySmtp.ContainsKey($primaryKey)) {
        $contactsByPrimarySmtp[$primaryKey] = $contact
    }

    $aliasKey = ([string]$contact.Alias).Trim()
    if (-not [string]::IsNullOrWhiteSpace($aliasKey) -and -not $contactsByAlias.ContainsKey($aliasKey)) {
        $contactsByAlias[$aliasKey] = $contact
    }
}

$rowNumber = 1
foreach ($row in $rows) {
    $externalEmailAddress = Get-TrimmedValue -Value $row.ExternalEmailAddress
    $name = Get-TrimmedValue -Value $row.Name
    $alias = Get-TrimmedValue -Value $row.Alias

    try {
        if ([string]::IsNullOrWhiteSpace($externalEmailAddress)) {
            throw 'ExternalEmailAddress is required.'
        }

        if ([string]::IsNullOrWhiteSpace($name)) {
            $name = $externalEmailAddress.Split('@')[0]
        }

        if ([string]::IsNullOrWhiteSpace($alias)) {
            $alias = (($externalEmailAddress.Split('@')[0]) -replace '[^a-zA-Z0-9._-]', '')
        }

        if ([string]::IsNullOrWhiteSpace($alias)) {
            throw 'Alias is empty after sanitization. Provide an Alias value in the CSV.'
        }

        $externalKey = ConvertTo-NormalizedSmtpAddress -Value $externalEmailAddress
        $primarySmtpAddress = Get-TrimmedValue -Value $row.PrimarySmtpAddress
        $primarySmtpKey = ConvertTo-NormalizedSmtpAddress -Value $primarySmtpAddress

        $existingContact = $null
        if ($contactsByExternalEmail.ContainsKey($externalKey)) {
            $existingContact = $contactsByExternalEmail[$externalKey]
        }
        elseif (-not [string]::IsNullOrWhiteSpace($primarySmtpKey) -and $contactsByPrimarySmtp.ContainsKey($primarySmtpKey)) {
            $existingContact = $contactsByPrimarySmtp[$primarySmtpKey]
        }
        elseif ($contactsByAlias.ContainsKey($alias)) {
            $existingContact = $contactsByAlias[$alias]
        }

        if ($existingContact) {
            $existingIdentity = Get-TrimmedValue -Value $existingContact.Identity
            if ([string]::IsNullOrWhiteSpace($existingIdentity)) {
                $existingIdentity = Get-TrimmedValue -Value $existingContact.Name
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $externalEmailAddress -Action 'CreateMailContact' -Status 'Skipped' -Message "Mail contact already exists (identity '$existingIdentity')."))
            $rowNumber++
            continue
        }

        $params = @{
            Name                 = $name
            ExternalEmailAddress = $externalEmailAddress
            Alias                = $alias
        }

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $params.DisplayName = $displayName
        }

        $firstName = Get-TrimmedValue -Value $row.FirstName
        if (-not [string]::IsNullOrWhiteSpace($firstName)) {
            $params.FirstName = $firstName
        }

        $lastName = Get-TrimmedValue -Value $row.LastName
        if (-not [string]::IsNullOrWhiteSpace($lastName)) {
            $params.LastName = $lastName
        }

        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $params.PrimarySmtpAddress = $primarySmtpAddress
        }

        $organizationalUnit = Get-OptionalColumnValue -Row $row -ColumnName 'OrganizationalUnit'
        if ($supportsOrganizationalUnit -and -not [string]::IsNullOrWhiteSpace($organizationalUnit)) {
            $params.OrganizationalUnit = $organizationalUnit
        }

        $hiddenRaw = Get-TrimmedValue -Value $row.HiddenFromAddressListsEnabled
        $setHidden = -not [string]::IsNullOrWhiteSpace($hiddenRaw)

        if ($PSCmdlet.ShouldProcess($externalEmailAddress, 'Create Exchange on-prem mail contact')) {
            $createdContact = Invoke-WithRetry -OperationName "Create mail contact $externalEmailAddress" -ScriptBlock {
                New-MailContact @params -ErrorAction Stop
            }

            if ($setHidden) {
                $hiddenValue = ConvertTo-Bool -Value $hiddenRaw
                Invoke-WithRetry -OperationName "Set hidden from GAL for $externalEmailAddress" -ScriptBlock {
                    Set-MailContact -Identity $createdContact.Identity -HiddenFromAddressListsEnabled $hiddenValue -ErrorAction Stop
                }
            }

            $createdExternalKey = ConvertTo-NormalizedSmtpAddress -Value ([string]$createdContact.ExternalEmailAddress)
            if (-not [string]::IsNullOrWhiteSpace($createdExternalKey) -and -not $contactsByExternalEmail.ContainsKey($createdExternalKey)) {
                $contactsByExternalEmail[$createdExternalKey] = $createdContact
            }

            $createdPrimaryKey = ConvertTo-NormalizedSmtpAddress -Value ([string]$createdContact.PrimarySmtpAddress)
            if (-not [string]::IsNullOrWhiteSpace($createdPrimaryKey) -and -not $contactsByPrimarySmtp.ContainsKey($createdPrimaryKey)) {
                $contactsByPrimarySmtp[$createdPrimaryKey] = $createdContact
            }

            $createdAlias = Get-TrimmedValue -Value $createdContact.Alias
            if (-not [string]::IsNullOrWhiteSpace($createdAlias) -and -not $contactsByAlias.ContainsKey($createdAlias)) {
                $contactsByAlias[$createdAlias] = $createdContact
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $externalEmailAddress -Action 'CreateMailContact' -Status 'Created' -Message 'Mail contact created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $externalEmailAddress -Action 'CreateMailContact' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($externalEmailAddress) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $externalEmailAddress -Action 'CreateMailContact' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem mail contact creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}

