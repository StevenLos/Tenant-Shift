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

.SYNOPSIS
    Gets ExchangeOnPremSharedMailboxes and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnPremSharedMailboxes from Active Directory and writes the results to a CSV file.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all objects in scope (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all objects in scope rather than processing from an input CSV file. Uses the DiscoverAll parameter set.

.PARAMETER SearchBase
    Distinguished name of the Active Directory OU to scope the discovery. If omitted, searches the entire domain.

.PARAMETER Server
    Active Directory domain controller to target. If omitted, uses the default DC for the current domain.

.PARAMETER MaxObjects
    Maximum number of objects to retrieve. 0 (default) means no limit.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-IR0216-Get-ExchangeOnPremSharedMailboxes.ps1 -InputCsvPath .\0216.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR0216-Get-ExchangeOnPremSharedMailboxes.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Exchange Management Shell cmdlets (session)
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR0216-Get-ExchangeOnPremSharedMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

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

function Resolve-SharedMailboxesByScope {
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
            RecipientTypeDetails = 'SharedMailbox'
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
    if ($mailbox -and (Get-TrimmedValue -Value $mailbox.RecipientTypeDetails) -eq 'SharedMailbox') {
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

$requiredHeaders = @(
    'SharedMailboxIdentity'
)

Write-Status -Message 'Starting Exchange on-prem shared mailbox inventory script.'
Ensure-ExchangeOnPremConnection

$scopeMode = 'Csv'
$resolvedServer = Get-TrimmedValue -Value $Server
$resolvedSearchBase = ''
$runWasTruncated = $false

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    $resolvedSearchBase = Get-TrimmedValue -Value $SearchBase
    Write-Status -Message "DiscoverAll enabled for Exchange on-prem shared mailboxes. SearchBase='$resolvedSearchBase'." -Level WARN
    $rows = @([PSCustomObject]@{
            SharedMailboxIdentity = '*'
        })
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $sharedMailboxIdentity = Get-TrimmedValue -Value $row.SharedMailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($sharedMailboxIdentity)) {
            throw 'SharedMailboxIdentity is required. Use * to inventory all shared mailboxes.'
        }

        $effectiveSearchBase = if ($scopeMode -eq 'DiscoverAll') { $resolvedSearchBase } else { '' }
        $mailboxes = @(Invoke-WithRetry -OperationName "Load shared mailboxes for $sharedMailboxIdentity" -ScriptBlock {
            Resolve-SharedMailboxesByScope -Identity $sharedMailboxIdentity -SearchBase $effectiveSearchBase -Server $resolvedServer
        })

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $mailboxes.Count -gt $MaxObjects) {
            $mailboxes = @($mailboxes | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($mailboxes.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentity -Action 'GetExchangeSharedMailbox' -Status 'NotFound' -Message 'No matching shared mailboxes were found.' -Data ([ordered]@{
                        SharedMailboxIdentity             = $sharedMailboxIdentity
                        DisplayName                       = ''
                        Alias                             = ''
                        PrimarySmtpAddress                = ''
                        UserPrincipalName                 = ''
                        HiddenFromAddressListsEnabled     = ''
                        GrantSendOnBehalfTo               = ''
                        MessageCopyForSentAsEnabled       = ''
                        MessageCopyForSendOnBehalfEnabled = ''
                        ForwardingSmtpAddress             = ''
                        ForwardingAddress                 = ''
                        DeliverToMailboxAndForward        = ''
                        AuditEnabled                      = ''
                        LitigationHoldEnabled             = ''
                        MailTip                           = ''
                        WhenCreatedUTC                    = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $identity = Get-TrimmedValue -Value $mailbox.Identity
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeSharedMailbox' -Status 'Completed' -Message 'Shared mailbox exported.' -Data ([ordered]@{
                        SharedMailboxIdentity             = $identity
                        DisplayName                       = Get-TrimmedValue -Value $mailbox.DisplayName
                        Alias                             = Get-TrimmedValue -Value $mailbox.Alias
                        PrimarySmtpAddress                = Get-TrimmedValue -Value $mailbox.PrimarySmtpAddress
                        UserPrincipalName                 = Get-TrimmedValue -Value $mailbox.UserPrincipalName
                        HiddenFromAddressListsEnabled     = [string]$mailbox.HiddenFromAddressListsEnabled
                        GrantSendOnBehalfTo               = Convert-MultiValueToString -Value $mailbox.GrantSendOnBehalfTo
                        MessageCopyForSentAsEnabled       = [string]$mailbox.MessageCopyForSentAsEnabled
                        MessageCopyForSendOnBehalfEnabled = [string]$mailbox.MessageCopyForSendOnBehalfEnabled
                        ForwardingSmtpAddress             = Get-TrimmedValue -Value $mailbox.ForwardingSmtpAddress
                        ForwardingAddress                 = Get-TrimmedValue -Value $mailbox.ForwardingAddress
                        DeliverToMailboxAndForward        = [string]$mailbox.DeliverToMailboxAndForward
                        AuditEnabled                      = [string]$mailbox.AuditEnabled
                        LitigationHoldEnabled             = [string]$mailbox.LitigationHoldEnabled
                        MailTip                           = Get-TrimmedValue -Value $mailbox.MailTip
                        WhenCreatedUTC                    = [string]$mailbox.WhenCreatedUTC
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($sharedMailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentity -Action 'GetExchangeSharedMailbox' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    SharedMailboxIdentity             = $sharedMailboxIdentity
                    DisplayName                       = ''
                    Alias                             = ''
                    PrimarySmtpAddress                = ''
                    UserPrincipalName                 = ''
                    HiddenFromAddressListsEnabled     = ''
                    GrantSendOnBehalfTo               = ''
                    MessageCopyForSentAsEnabled       = ''
                    MessageCopyForSendOnBehalfEnabled = ''
                    ForwardingSmtpAddress             = ''
                    ForwardingAddress                 = ''
                    DeliverToMailboxAndForward        = ''
                    AuditEnabled                      = ''
                    LitigationHoldEnabled             = ''
                    MailTip                           = ''
                    WhenCreatedUTC                    = ''
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
Write-Status -Message 'Exchange on-prem shared mailbox inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
