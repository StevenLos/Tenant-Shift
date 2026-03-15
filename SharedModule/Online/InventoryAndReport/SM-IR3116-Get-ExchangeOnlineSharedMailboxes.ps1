<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-005957

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR3116-Get-ExchangeOnlineSharedMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

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

Write-Status -Message 'Starting Exchange Online shared mailbox inventory script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$scopeMode = 'Csv'
if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. CSV input is bypassed.' -Level WARN

    $discoverRow = [ordered]@{}
    foreach ($header in $requiredHeaders) {
        $discoverRow[$header] = '*'
    }

    $rows = @([PSCustomObject]$discoverRow)
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $sharedMailboxIdentity = ([string]$row.SharedMailboxIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($sharedMailboxIdentity)) {
            throw 'SharedMailboxIdentity is required. Use * to inventory all shared mailboxes.'
        }

        $mailboxes = @()
        if ($sharedMailboxIdentity -eq '*') {
            $mailboxes = @(Invoke-WithRetry -OperationName 'Load all shared mailboxes' -ScriptBlock {
                Get-ExchangeOnlineMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop
            })
        }
        else {
            $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $sharedMailboxIdentity" -ScriptBlock {
                Get-ExchangeOnlineMailbox -Identity $sharedMailboxIdentity -ErrorAction SilentlyContinue
            }

            if ($mailbox -and ([string]$mailbox.RecipientTypeDetails).Trim() -eq 'SharedMailbox') {
                $mailboxes = @($mailbox)
            }
        }

        if ($mailboxes.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentity -Action 'GetExchangeSharedMailbox' -Status 'NotFound' -Message 'No matching shared mailboxes were found.' -Data ([ordered]@{
                        SharedMailboxIdentity                     = $sharedMailboxIdentity
                        DisplayName                               = ''
                        Alias                                     = ''
                        PrimarySmtpAddress                        = ''
                        UserPrincipalName                         = ''
                        HiddenFromAddressListsEnabled             = ''
                        GrantSendOnBehalfTo                       = ''
                        MessageCopyForSentAsEnabled               = ''
                        MessageCopyForSendOnBehalfEnabled         = ''
                        ForwardingSmtpAddress                     = ''
                        DeliverToMailboxAndForward                = ''
                        AuditEnabled                              = ''
                        LitigationHoldEnabled                     = ''
                        WhenCreatedUTC                            = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $identity = ([string]$mailbox.Identity).Trim()
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeSharedMailbox' -Status 'Completed' -Message 'Shared mailbox exported.' -Data ([ordered]@{
                        SharedMailboxIdentity                     = $identity
                        DisplayName                               = ([string]$mailbox.DisplayName).Trim()
                        Alias                                     = ([string]$mailbox.Alias).Trim()
                        PrimarySmtpAddress                        = ([string]$mailbox.PrimarySmtpAddress).Trim()
                        UserPrincipalName                         = ([string]$mailbox.UserPrincipalName).Trim()
                        HiddenFromAddressListsEnabled             = [string]$mailbox.HiddenFromAddressListsEnabled
                        GrantSendOnBehalfTo                       = Convert-MultiValueToString -Value $mailbox.GrantSendOnBehalfTo
                        MessageCopyForSentAsEnabled               = [string]$mailbox.MessageCopyForSentAsEnabled
                        MessageCopyForSendOnBehalfEnabled         = [string]$mailbox.MessageCopyForSendOnBehalfEnabled
                        ForwardingSmtpAddress                     = ([string]$mailbox.ForwardingSmtpAddress).Trim()
                        DeliverToMailboxAndForward                = [string]$mailbox.DeliverToMailboxAndForward
                        AuditEnabled                              = [string]$mailbox.AuditEnabled
                        LitigationHoldEnabled                     = [string]$mailbox.LitigationHoldEnabled
                        WhenCreatedUTC                            = [string]$mailbox.WhenCreatedUTC
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($sharedMailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentity -Action 'GetExchangeSharedMailbox' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    SharedMailboxIdentity                     = $sharedMailboxIdentity
                    DisplayName                               = ''
                    Alias                                     = ''
                    PrimarySmtpAddress                        = ''
                    UserPrincipalName                         = ''
                    HiddenFromAddressListsEnabled             = ''
                    GrantSendOnBehalfTo                       = ''
                    MessageCopyForSentAsEnabled               = ''
                    MessageCopyForSendOnBehalfEnabled         = ''
                    ForwardingSmtpAddress                     = ''
                    DeliverToMailboxAndForward                = ''
                    AuditEnabled                              = ''
                    LitigationHoldEnabled                     = ''
                    WhenCreatedUTC                            = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online shared mailbox inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}










