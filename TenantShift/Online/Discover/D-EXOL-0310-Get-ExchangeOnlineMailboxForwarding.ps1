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
    Exports mailboxes with ForwardingAddress or ForwardingSmtpAddress configured.

.DESCRIPTION
    Exports all mailboxes in the tenant that have ForwardingAddress or ForwardingSmtpAddress
    set at the mailbox level (as opposed to inbox rules, which are covered by D-EXOL-0300).
    Includes the DeliverToMailboxAndForward flag indicating whether mail is both delivered
    to the mailbox and forwarded, or forwarded only.
    All results — including mailboxes that could not be queried — are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include UserPrincipalName.
    See the companion .input.csv template for the full column list.

.PARAMETER DiscoverAll
    Enumerate all user and shared mailboxes in the tenant rather than processing from an input CSV file.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-EXOL-0310-Get-ExchangeOnlineMailboxForwarding.ps1 -InputCsvPath .\D-EXOL-0310-Get-ExchangeOnlineMailboxForwarding.input.csv

    Export mailbox forwarding settings for mailboxes listed in the input CSV.

.EXAMPLE
    .\D-EXOL-0310-Get-ExchangeOnlineMailboxForwarding.ps1 -DiscoverAll

    Export mailbox forwarding settings for all user and shared mailboxes in the tenant.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator
    Limitations:      Covers mailbox-level forwarding only. Inbox rule forwarding is covered by D-EXOL-0300.
                      Transport rule forwarding is covered by D-EXOL-0320.

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-EXOL-0310-Get-ExchangeOnlineMailboxForwarding_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'UserPrincipalName',
    'DisplayName',
    'RecipientTypeDetails',
    'ForwardingAddress',
    'ForwardingSmtpAddress',
    'DeliverToMailboxAndForward',
    'HasForwarding'
)

$mailboxProperties = @(
    'ForwardingAddress',
    'ForwardingSmtpAddress',
    'DeliverToMailboxAndForward',
    'DisplayName',
    'RecipientTypeDetails',
    'UserPrincipalName'
)

$requiredHeaders = @('UserPrincipalName')

Write-Status -Message 'Starting Exchange Online mailbox forwarding audit script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$scopeMode = 'Csv'

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. Fetching all mailboxes with forwarding configured.' -Level WARN

    $forwardingMailboxes = Invoke-WithRetry -OperationName 'Get mailboxes with forwarding' -ScriptBlock {
        Get-Mailbox -RecipientTypeDetails UserMailbox, SharedMailbox -Filter {
            ForwardingAddress -ne $null -or ForwardingSmtpAddress -ne $null
        } -ResultSize Unlimited -Properties $mailboxProperties -ErrorAction Stop
    }

    Write-Status -Message "Found $($forwardingMailboxes.Count) mailboxes with forwarding configured."
    $rows = @($forwardingMailboxes | ForEach-Object {
        [PSCustomObject]@{ UserPrincipalName = [string]$_.UserPrincipalName }
    })

    if ($rows.Count -eq 0) {
        Write-Status -Message 'No mailboxes with forwarding configured were found. Output will contain a single informational row.'
    }
} else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $upn        = Get-TrimmedValue -Value $row.UserPrincipalName
    $primaryKey = $upn

    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Status -Message "Row $rowNumber skipped: UserPrincipalName is empty." -Level WARN
        $rowNumber++
        continue
    }

    try {
        $mailbox = Invoke-WithRetry -OperationName "Get mailbox $upn" -ScriptBlock {
            Get-Mailbox -Identity $upn -Properties $mailboxProperties -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetMailboxForwarding' -Status 'NotFound' -Message 'Mailbox not found.' -Data ([ordered]@{
                UserPrincipalName          = $upn
                DisplayName                = ''
                RecipientTypeDetails       = ''
                ForwardingAddress          = ''
                ForwardingSmtpAddress      = ''
                DeliverToMailboxAndForward = ''
                HasForwarding              = ''
            })))
            $rowNumber++
            continue
        }

        $fwdAddress    = Get-TrimmedValue -Value $mailbox.ForwardingAddress
        $fwdSmtp       = Get-TrimmedValue -Value $mailbox.ForwardingSmtpAddress
        $hasForwarding = (-not [string]::IsNullOrWhiteSpace($fwdAddress)) -or (-not [string]::IsNullOrWhiteSpace($fwdSmtp))

        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetMailboxForwarding' -Status 'Completed' -Message $(if ($hasForwarding) { 'Mailbox forwarding configured.' } else { 'No mailbox-level forwarding configured.' }) -Data ([ordered]@{
            UserPrincipalName          = Get-TrimmedValue -Value $mailbox.UserPrincipalName
            DisplayName                = Get-TrimmedValue -Value $mailbox.DisplayName
            RecipientTypeDetails       = [string]$mailbox.RecipientTypeDetails
            ForwardingAddress          = $fwdAddress
            ForwardingSmtpAddress      = $fwdSmtp
            DeliverToMailboxAndForward = [string]$mailbox.DeliverToMailboxAndForward
            HasForwarding              = [string]$hasForwarding
        })))
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetMailboxForwarding' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            UserPrincipalName          = $upn
            DisplayName                = ''
            RecipientTypeDetails       = ''
            ForwardingAddress          = ''
            ForwardingSmtpAddress      = ''
            DeliverToMailboxAndForward = ''
            HasForwarding              = ''
        })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mailbox forwarding audit script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
