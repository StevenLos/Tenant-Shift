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

.SYNOPSIS
    Provisions ExchangeOnlineSharedMailboxes in Microsoft 365.

.DESCRIPTION
    Creates ExchangeOnlineSharedMailboxes in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P3116-Create-ExchangeOnlineSharedMailboxes.ps1 -InputCsvPath .\3116.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P3116-Create-ExchangeOnlineSharedMailboxes.ps1 -InputCsvPath .\3116.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator
    Limitations:      None known.

    CSV Fields:
    Column                             Type      Required  Description
    ---------------------------------  ----      --------  -----------
    Name                               String    Yes       <fill in description>
    Alias                              String    Yes       <fill in description>
    DisplayName                        String    Yes       <fill in description>
    UserPrincipalName                  String    Yes       <fill in description>
    PrimarySmtpAddress                 String    Yes       <fill in description>
    HiddenFromAddressListsEnabled      String    Yes       <fill in description>
    GrantSendOnBehalfTo                String    Yes       <fill in description>
    MessageCopyForSentAsEnabled        String    Yes       <fill in description>
    MessageCopyForSendOnBehalfEnabled  String    Yes       <fill in description>
    ForwardingSmtpAddress              String    Yes       <fill in description>
    DeliverToMailboxAndForward         String    Yes       <fill in description>
    AuditEnabled                       String    Yes       <fill in description>
    LitigationHoldEnabled              String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P3116-Create-ExchangeOnlineSharedMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'Name',
    'Alias',
    'DisplayName',
    'UserPrincipalName',
    'PrimarySmtpAddress',
    'HiddenFromAddressListsEnabled',
    'GrantSendOnBehalfTo',
    'MessageCopyForSentAsEnabled',
    'MessageCopyForSendOnBehalfEnabled',
    'ForwardingSmtpAddress',
    'DeliverToMailboxAndForward',
    'AuditEnabled',
    'LitigationHoldEnabled'
)

Write-Status -Message 'Starting Exchange Online shared mailbox creation script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$newMailboxCommand = Get-Command -Name New-Mailbox -ErrorAction Stop
$newMailboxSupportsUserPrincipalName = $newMailboxCommand.Parameters.ContainsKey('UserPrincipalName')
$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop

$supports = @{
    HiddenFromAddressListsEnabled      = $setMailboxCommand.Parameters.ContainsKey('HiddenFromAddressListsEnabled')
    GrantSendOnBehalfTo                = $setMailboxCommand.Parameters.ContainsKey('GrantSendOnBehalfTo')
    MessageCopyForSentAsEnabled        = $setMailboxCommand.Parameters.ContainsKey('MessageCopyForSentAsEnabled')
    MessageCopyForSendOnBehalfEnabled  = $setMailboxCommand.Parameters.ContainsKey('MessageCopyForSendOnBehalfEnabled')
    ForwardingSmtpAddress              = $setMailboxCommand.Parameters.ContainsKey('ForwardingSmtpAddress')
    DeliverToMailboxAndForward         = $setMailboxCommand.Parameters.ContainsKey('DeliverToMailboxAndForward')
    AuditEnabled                       = $setMailboxCommand.Parameters.ContainsKey('AuditEnabled')
    LitigationHoldEnabled              = $setMailboxCommand.Parameters.ContainsKey('LitigationHoldEnabled')
}

if (-not $newMailboxSupportsUserPrincipalName) {
    Write-Status -Message "New-Mailbox in this session does not support -UserPrincipalName. The 'UserPrincipalName' CSV value will be ignored." -Level WARN
}

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $name = ([string]$row.Name).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Name is required.'
        }

        $alias = ([string]$row.Alias).Trim()
        $displayName = ([string]$row.DisplayName).Trim()
        $userPrincipalName = ([string]$row.UserPrincipalName).Trim()
        $primarySmtpAddress = ([string]$row.PrimarySmtpAddress).Trim()

        $lookupIdentity = if (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) {
            $userPrincipalName
        }
        elseif (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $primarySmtpAddress
        }
        elseif (-not [string]::IsNullOrWhiteSpace($alias)) {
            $alias
        }
        else {
            $name
        }

        $existingMailbox = Invoke-WithRetry -OperationName "Lookup shared mailbox $lookupIdentity" -ScriptBlock {
            Get-ExchangeOnlineMailbox -Identity $lookupIdentity -ErrorAction SilentlyContinue
        }
        if ($existingMailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateSharedMailbox' -Status 'Skipped' -Message 'Shared mailbox already exists.'))
            $rowNumber++
            continue
        }

        $params = @{
            Shared = $true
            Name   = $name
        }

        if (-not [string]::IsNullOrWhiteSpace($alias)) {
            $params.Alias = $alias
        }

        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $params.DisplayName = $displayName
        }

        $upnIgnored = $false
        if (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) {
            if ($newMailboxSupportsUserPrincipalName) {
                $params.UserPrincipalName = $userPrincipalName
            }
            else {
                $upnIgnored = $true
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $params.PrimarySmtpAddress = $primarySmtpAddress
        }

        if ($PSCmdlet.ShouldProcess($lookupIdentity, 'Create Exchange Online shared mailbox')) {
            $createdMailbox = Invoke-WithRetry -OperationName "Create shared mailbox $lookupIdentity" -ScriptBlock {
                New-Mailbox @params -ErrorAction Stop
            }

            $setParams = @{
                Identity = $createdMailbox.Identity
            }
            $setMessages = [System.Collections.Generic.List[string]]::new()

            $hiddenRaw = ([string]$row.HiddenFromAddressListsEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
                if ($supports.HiddenFromAddressListsEnabled) {
                    $setParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
                }
                else {
                    $setMessages.Add('HiddenFromAddressListsEnabled ignored (unsupported parameter).')
                }
            }

            $sendOnBehalfList = ConvertTo-Array -Value ([string]$row.GrantSendOnBehalfTo)
            if ($sendOnBehalfList.Count -gt 0) {
                if ($supports.GrantSendOnBehalfTo) {
                    $setParams.GrantSendOnBehalfTo = @{ Add = $sendOnBehalfList }
                }
                else {
                    $setMessages.Add('GrantSendOnBehalfTo ignored (unsupported parameter).')
                }
            }

            $sentAsCopyRaw = ([string]$row.MessageCopyForSentAsEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($sentAsCopyRaw)) {
                if ($supports.MessageCopyForSentAsEnabled) {
                    $setParams.MessageCopyForSentAsEnabled = ConvertTo-Bool -Value $sentAsCopyRaw
                }
                else {
                    $setMessages.Add('MessageCopyForSentAsEnabled ignored (unsupported parameter).')
                }
            }

            $sendOnBehalfCopyRaw = ([string]$row.MessageCopyForSendOnBehalfEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($sendOnBehalfCopyRaw)) {
                if ($supports.MessageCopyForSendOnBehalfEnabled) {
                    $setParams.MessageCopyForSendOnBehalfEnabled = ConvertTo-Bool -Value $sendOnBehalfCopyRaw
                }
                else {
                    $setMessages.Add('MessageCopyForSendOnBehalfEnabled ignored (unsupported parameter).')
                }
            }

            $forwardingSmtpAddress = ([string]$row.ForwardingSmtpAddress).Trim()
            if (-not [string]::IsNullOrWhiteSpace($forwardingSmtpAddress)) {
                if ($supports.ForwardingSmtpAddress) {
                    $setParams.ForwardingSmtpAddress = $forwardingSmtpAddress
                }
                else {
                    $setMessages.Add('ForwardingSmtpAddress ignored (unsupported parameter).')
                }
            }

            $deliverAndForwardRaw = ([string]$row.DeliverToMailboxAndForward).Trim()
            if (-not [string]::IsNullOrWhiteSpace($deliverAndForwardRaw)) {
                if ($supports.DeliverToMailboxAndForward) {
                    $setParams.DeliverToMailboxAndForward = ConvertTo-Bool -Value $deliverAndForwardRaw
                }
                else {
                    $setMessages.Add('DeliverToMailboxAndForward ignored (unsupported parameter).')
                }
            }

            $auditEnabledRaw = ([string]$row.AuditEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($auditEnabledRaw)) {
                if ($supports.AuditEnabled) {
                    $setParams.AuditEnabled = ConvertTo-Bool -Value $auditEnabledRaw
                }
                else {
                    $setMessages.Add('AuditEnabled ignored (unsupported parameter).')
                }
            }

            $litigationHoldRaw = ([string]$row.LitigationHoldEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($litigationHoldRaw)) {
                if ($supports.LitigationHoldEnabled) {
                    $setParams.LitigationHoldEnabled = ConvertTo-Bool -Value $litigationHoldRaw
                }
                else {
                    $setMessages.Add('LitigationHoldEnabled ignored (unsupported parameter).')
                }
            }

            if ($setParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Set shared mailbox options $lookupIdentity" -ScriptBlock {
                    Set-Mailbox @setParams -ErrorAction Stop
                }
            }

            $successMessage = 'Shared mailbox created successfully.'
            if ($upnIgnored) {
                $successMessage = "$successMessage UserPrincipalName was provided but ignored because this New-Mailbox session does not support -UserPrincipalName."
            }
            if ($setMessages.Count -gt 0) {
                $successMessage = "$successMessage $($setMessages -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateSharedMailbox' -Status 'Created' -Message $successMessage))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateSharedMailbox' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($name) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $name -Action 'CreateSharedMailbox' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online shared mailbox creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}






