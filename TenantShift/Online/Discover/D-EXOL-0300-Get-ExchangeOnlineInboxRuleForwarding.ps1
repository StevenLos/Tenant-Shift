<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260406-000000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Exports inbox rules with forwarding actions across Exchange Online mailboxes.

.DESCRIPTION
    Exports all inbox rules across user mailboxes that have ForwardTo, ForwardAsAttachmentTo,
    or RedirectTo actions configured. For each matching rule, outputs the mailbox owner,
    rule name, rule state, and resolved recipient address.
    Supports progress reporting and a -BatchSize parameter to manage throughput on large tenants.
    All results — including mailboxes that could not be queried — are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include UserPrincipalName.
    See the companion .input.csv template for the full column list.

.PARAMETER DiscoverAll
    Enumerate all user mailboxes in the tenant rather than processing from an input CSV file.

.PARAMETER BatchSize
    Number of mailboxes to process per batch when writing progress. Default is 50.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-EXOL-0300-Get-ExchangeOnlineInboxRuleForwarding.ps1 -InputCsvPath .\D-EXOL-0300-Get-ExchangeOnlineInboxRuleForwarding.input.csv

    Export inbox rule forwarding data for mailboxes listed in the input CSV.

.EXAMPLE
    .\D-EXOL-0300-Get-ExchangeOnlineInboxRuleForwarding.ps1 -DiscoverAll -BatchSize 25

    Export inbox rule forwarding data for all user mailboxes in the tenant.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator
    Limitations:      Can be slow on large tenants. Get-InboxRule runs per mailbox.
                      Use -BatchSize to tune throughput vs. connection stability.

    CSV Fields:
    Column              Type      Required  Description
    ------------------  --------  --------  -----------
    UserPrincipalName   String    Yes       UPN of the mailbox to inspect
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(ParameterSetName = 'FromCsv')]
    [Parameter(ParameterSetName = 'DiscoverAll')]
    [ValidateRange(1, 500)]
    [int]$BatchSize = 50,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-EXOL-0300-Get-ExchangeOnlineInboxRuleForwarding_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
        [Parameter(Mandatory)][int]$RowNumber,
        [Parameter(Mandatory)][string]$PrimaryKey,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$Message,
        [Parameter(Mandatory)][hashtable]$Data
    )

    $base    = New-ResultObject -RowNumber $RowNumber -PrimaryKey $PrimaryKey -Action $Action -Status $Status -Message $Message
    $ordered = [ordered]@{}

    foreach ($prop in $base.PSObject.Properties.Name) {
        $ordered[$prop] = $base.$prop
    }

    foreach ($key in $Data.Keys) {
        $ordered[$key] = $Data[$key]
    }

    return [PSCustomObject]$ordered
}

function New-EmptyForwardingRuleData {
    [CmdletBinding()]
    param([string]$UserPrincipalName = '')

    return [ordered]@{
        UserPrincipalName  = $UserPrincipalName
        MailboxDisplayName = ''
        RuleName           = ''
        RuleEnabled        = ''
        ForwardTo          = ''
        ForwardAsAttachmentTo = ''
        RedirectTo         = ''
        ForwardingActionCount = ''
    }
}

function Resolve-AddressList {
    [CmdletBinding()]
    param([object]$AddressCollection)

    if ($null -eq $AddressCollection) { return '' }
    $items = @($AddressCollection | ForEach-Object { Get-TrimmedValue -Value ([string]$_) } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    return $items -join '; '
}

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'UserPrincipalName',
    'MailboxDisplayName',
    'RuleName',
    'RuleEnabled',
    'ForwardTo',
    'ForwardAsAttachmentTo',
    'RedirectTo',
    'ForwardingActionCount'
)

$requiredHeaders = @('UserPrincipalName')

Write-Status -Message 'Starting Exchange Online inbox rule forwarding audit script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$scopeMode = 'Csv'

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. Fetching all user mailboxes.' -Level WARN

    $allMailboxes = Invoke-WithRetry -OperationName 'Get all user mailboxes' -ScriptBlock {
        Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -ErrorAction Stop |
            Select-Object UserPrincipalName, DisplayName
    }

    Write-Status -Message "Found $($allMailboxes.Count) user mailboxes."
    $rows = @($allMailboxes | ForEach-Object {
        [PSCustomObject]@{ UserPrincipalName = [string]$_.UserPrincipalName }
    })
} else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results    = [System.Collections.Generic.List[object]]::new()
$rowNumber  = 1
$processed  = 0
$totalRows  = $rows.Count

foreach ($row in $rows) {
    $upn        = Get-TrimmedValue -Value $row.UserPrincipalName
    $primaryKey = $upn

    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Status -Message "Row $rowNumber skipped: UserPrincipalName is empty." -Level WARN
        $rowNumber++
        continue
    }

    try {
        # Get mailbox display name.
        $mailbox = Invoke-WithRetry -OperationName "Get mailbox $upn" -ScriptBlock {
            Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetInboxRuleForwarding' -Status 'NotFound' -Message 'Mailbox not found.' -Data (New-EmptyForwardingRuleData -UserPrincipalName $upn)))
            $rowNumber++
            $processed++
            continue
        }

        $displayName = Get-TrimmedValue -Value $mailbox.DisplayName

        # Get all inbox rules for this mailbox.
        $inboxRules = Invoke-WithRetry -OperationName "Get inbox rules for $upn" -ScriptBlock {
            Get-InboxRule -Mailbox $upn -ErrorAction Stop
        }

        # Filter to rules with forwarding/redirect actions.
        $forwardingRules = @($inboxRules | Where-Object {
            ($null -ne $_.ForwardTo -and @($_.ForwardTo).Count -gt 0) -or
            ($null -ne $_.ForwardAsAttachmentTo -and @($_.ForwardAsAttachmentTo).Count -gt 0) -or
            ($null -ne $_.RedirectTo -and @($_.RedirectTo).Count -gt 0)
        })

        if ($forwardingRules.Count -eq 0) {
            # Mailbox has no forwarding inbox rules — output a single NoForwardingRules row.
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetInboxRuleForwarding' -Status 'Completed' -Message 'No inbox rules with forwarding actions found.' -Data ([ordered]@{
                UserPrincipalName     = $upn
                MailboxDisplayName    = $displayName
                RuleName              = ''
                RuleEnabled           = ''
                ForwardTo             = ''
                ForwardAsAttachmentTo = ''
                RedirectTo            = ''
                ForwardingActionCount = '0'
            })))
        } else {
            foreach ($rule in ($forwardingRules | Sort-Object -Property Name)) {
                $forwardTo           = Resolve-AddressList -AddressCollection $rule.ForwardTo
                $forwardAsAttachment = Resolve-AddressList -AddressCollection $rule.ForwardAsAttachmentTo
                $redirectTo          = Resolve-AddressList -AddressCollection $rule.RedirectTo
                $actionCount         = (@($rule.ForwardTo).Count + @($rule.ForwardAsAttachmentTo).Count + @($rule.RedirectTo).Count)

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "${upn}|$([string]$rule.Name)" -Action 'GetInboxRuleForwarding' -Status 'Completed' -Message 'Forwarding inbox rule exported.' -Data ([ordered]@{
                    UserPrincipalName     = $upn
                    MailboxDisplayName    = $displayName
                    RuleName              = Get-TrimmedValue -Value $rule.Name
                    RuleEnabled           = [string]$rule.Enabled
                    ForwardTo             = $forwardTo
                    ForwardAsAttachmentTo = $forwardAsAttachment
                    RedirectTo            = $redirectTo
                    ForwardingActionCount = [string]$actionCount
                })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetInboxRuleForwarding' -Status 'Failed' -Message $_.Exception.Message -Data (New-EmptyForwardingRuleData -UserPrincipalName $upn)))
    }

    $rowNumber++
    $processed++

    if ($processed % $BatchSize -eq 0) {
        Write-Status -Message "Progress: $processed / $totalRows mailboxes processed."
    }
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online inbox rule forwarding audit script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
