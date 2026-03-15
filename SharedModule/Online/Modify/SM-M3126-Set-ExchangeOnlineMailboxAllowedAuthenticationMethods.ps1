<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-153000

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3126-Set-ExchangeOnlineMailboxAllowedAuthenticationMethods_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

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

function ConvertTo-BooleanString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [bool]$Value
    )

    if ($Value) {
        return 'TRUE'
    }

    return 'FALSE'
}

$requiredHeaders = @(
    'MailboxIdentity',
    'AuthenticationProfile',
    'OWAEnabled',
    'MAPIEnabled',
    'IMAPEnabled',
    'POPEnabled',
    'OutlookMobileEnabled',
    'ActiveSyncEnabled',
    'EwsEnabled',
    'EwsAllowOutlook',
    'UniversalOutlookEnabled',
    'OneWinNativeOutlookEnabled',
    'SmtpClientAuthenticationDisabled',
    'Notes'
)

$profileValues = [System.Collections.Generic.Dictionary[string, hashtable]]::new([System.StringComparer]::OrdinalIgnoreCase)
$profileValues['EnableAll'] = [ordered]@{
    OWAEnabled                       = $true
    MAPIEnabled                      = $true
    IMAPEnabled                      = $true
    POPEnabled                       = $true
    OutlookMobileEnabled             = $true
    ActiveSyncEnabled                = $true
    EwsEnabled                       = $true
    EwsAllowOutlook                  = $true
    UniversalOutlookEnabled          = $true
    OneWinNativeOutlookEnabled       = $true
    SmtpClientAuthenticationDisabled = $false
}
$profileValues['DisableClientProtocols'] = [ordered]@{
    OWAEnabled                       = $false
    MAPIEnabled                      = $false
    IMAPEnabled                      = $false
    POPEnabled                       = $false
    OutlookMobileEnabled             = $false
    ActiveSyncEnabled                = $false
    EwsEnabled                       = $true
    EwsAllowOutlook                  = $false
    UniversalOutlookEnabled          = $false
    OneWinNativeOutlookEnabled       = $false
    SmtpClientAuthenticationDisabled = $true
}
$profileValues['DisableAll'] = [ordered]@{
    OWAEnabled                       = $false
    MAPIEnabled                      = $false
    IMAPEnabled                      = $false
    POPEnabled                       = $false
    OutlookMobileEnabled             = $false
    ActiveSyncEnabled                = $false
    EwsEnabled                       = $false
    EwsAllowOutlook                  = $false
    UniversalOutlookEnabled          = $false
    OneWinNativeOutlookEnabled       = $false
    SmtpClientAuthenticationDisabled = $true
}

$fieldToColumnMap = [ordered]@{
    OWAEnabled                       = 'OWAEnabled'
    MAPIEnabled                      = 'MAPIEnabled'
    IMAPEnabled                      = 'IMAPEnabled'
    POPEnabled                       = 'POPEnabled'
    OutlookMobileEnabled             = 'OutlookMobileEnabled'
    ActiveSyncEnabled                = 'ActiveSyncEnabled'
    EwsEnabled                       = 'EwsEnabled'
    EwsAllowOutlook                  = 'EwsAllowOutlook'
    UniversalOutlookEnabled          = 'UniversalOutlookEnabled'
    OneWinNativeOutlookEnabled       = 'OneWinNativeOutlookEnabled'
    SmtpClientAuthenticationDisabled = 'SmtpClientAuthenticationDisabled'
}

Write-Status -Message 'Starting Exchange Online mailbox authentication-method update script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$setCasMailboxCommand = Get-Command -Name Set-CASMailbox -ErrorAction Stop
$supportedSetCasParameters = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($parameterName in $setCasMailboxCommand.Parameters.Keys) {
    $null = $supportedSetCasParameters.Add($parameterName)
}

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
            Get-ExchangeOnlineMailbox -Identity $mailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetMailboxAllowedAuthenticationMethods' -Status 'NotFound' -Message 'Mailbox not found.'))
            $rowNumber++
            continue
        }

        $casMailbox = Invoke-WithRetry -OperationName "Lookup CAS mailbox $mailboxIdentity" -ScriptBlock {
            Get-ExchangeOnlineCasMailbox -Identity $mailbox.Identity -ErrorAction Stop
        }

        $desiredValues = [ordered]@{}
        $profileName = Get-TrimmedValue -Value $row.AuthenticationProfile
        if (-not [string]::IsNullOrWhiteSpace($profileName)) {
            if (-not $profileValues.ContainsKey($profileName)) {
                throw "AuthenticationProfile '$profileName' is invalid. Use EnableAll, DisableClientProtocols, or DisableAll."
            }

            foreach ($key in $profileValues[$profileName].Keys) {
                $desiredValues[$key] = [bool]$profileValues[$profileName][$key]
            }
        }

        foreach ($fieldName in $fieldToColumnMap.Keys) {
            $columnName = [string]$fieldToColumnMap[$fieldName]
            $nullableBool = Get-NullableBool -Value $row.$columnName
            if ($null -ne $nullableBool) {
                $desiredValues[$fieldName] = [bool]$nullableBool
            }
        }

        if ($desiredValues.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetMailboxAllowedAuthenticationMethods' -Status 'Skipped' -Message 'No authentication-method updates were requested.'))
            $rowNumber++
            continue
        }

        $setParams = @{
            Identity = $mailbox.Identity
        }

        $warnings = [System.Collections.Generic.List[string]]::new()
        foreach ($fieldName in $desiredValues.Keys) {
            if (-not $supportedSetCasParameters.Contains($fieldName)) {
                $warnings.Add("Set-CASMailbox does not support parameter '$fieldName' in this session.")
                continue
            }

            $desiredValue = [bool]$desiredValues[$fieldName]
            $currentValue = [bool]$casMailbox.$fieldName

            if ($currentValue -ne $desiredValue) {
                $setParams[$fieldName] = $desiredValue
            }
        }

        if ($setParams.Count -eq 1) {
            $message = 'Mailbox authentication methods are already in the requested state.'
            if ($warnings.Count -gt 0) {
                $message = "$message Warnings: $($warnings -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetMailboxAllowedAuthenticationMethods' -Status 'Skipped' -Message $message))
            $rowNumber++
            continue
        }

        $changeSummary = [System.Collections.Generic.List[string]]::new()
        foreach ($fieldName in @($setParams.Keys | Where-Object { $_ -ne 'Identity' } | Sort-Object)) {
            $desiredText = ConvertTo-BooleanString -Value ([bool]$setParams[$fieldName])
            $changeSummary.Add("$fieldName=$desiredText")
        }

        if ($PSCmdlet.ShouldProcess($mailboxIdentity, 'Update Exchange Online mailbox authentication methods')) {
            Invoke-WithRetry -OperationName "Update mailbox authentication methods $mailboxIdentity" -ScriptBlock {
                Set-CASMailbox @setParams -ErrorAction Stop
            }

            $message = "Mailbox authentication methods updated: $($changeSummary -join '; ')."
            if ($warnings.Count -gt 0) {
                $message = "$message Warnings: $($warnings -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetMailboxAllowedAuthenticationMethods' -Status 'Updated' -Message $message))
        }
        else {
            $message = "Update skipped due to WhatIf. Requested changes: $($changeSummary -join '; ')."
            if ($warnings.Count -gt 0) {
                $message = "$message Warnings: $($warnings -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetMailboxAllowedAuthenticationMethods' -Status 'WhatIf' -Message $message))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetMailboxAllowedAuthenticationMethods' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mailbox authentication-method update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
