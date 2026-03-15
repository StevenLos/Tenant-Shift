<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-013000

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3124-Set-ExchangeOnlineUserMailboxForwarding_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

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

function Resolve-RecipientByIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,

        [Parameter(Mandatory)]
        [string]$OperationName
    )

    return Invoke-WithRetry -OperationName $OperationName -ScriptBlock {
        Get-ExchangeOnlineRecipient -Identity $Identity -ErrorAction SilentlyContinue
    }
}

$requiredHeaders = @(
    'MailboxIdentity',
    'ForwardingMode',
    'ForwardingSmtpAddress',
    'ForwardingRecipientIdentity',
    'DeliverToMailboxAndForward',
    'ExpectedPrimarySmtpAddress',
    'Notes'
)

Write-Status -Message 'Starting Exchange Online user mailbox forwarding script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop
$supports = @{
    ForwardingAddress          = $setMailboxCommand.Parameters.ContainsKey('ForwardingAddress')
    ForwardingSmtpAddress      = $setMailboxCommand.Parameters.ContainsKey('ForwardingSmtpAddress')
    DeliverToMailboxAndForward = $setMailboxCommand.Parameters.ContainsKey('DeliverToMailboxAndForward')
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

        $forwardingModeRaw = Get-TrimmedValue -Value $row.ForwardingMode
        if ([string]::IsNullOrWhiteSpace($forwardingModeRaw)) {
            throw 'ForwardingMode is required. Use Smtp, Recipient, or Clear.'
        }

        $forwardingMode = $forwardingModeRaw.ToLowerInvariant()
        if ($forwardingMode -notin @('smtp', 'recipient', 'clear')) {
            throw "ForwardingMode '$forwardingModeRaw' is invalid. Use Smtp, Recipient, or Clear."
        }

        $forwardingSmtpAddress = Get-TrimmedValue -Value $row.ForwardingSmtpAddress
        $forwardingRecipientIdentity = Get-TrimmedValue -Value $row.ForwardingRecipientIdentity

        if ($forwardingMode -eq 'smtp' -and [string]::IsNullOrWhiteSpace($forwardingSmtpAddress)) {
            throw 'ForwardingSmtpAddress is required when ForwardingMode is Smtp.'
        }

        if ($forwardingMode -eq 'recipient' -and [string]::IsNullOrWhiteSpace($forwardingRecipientIdentity)) {
            throw 'ForwardingRecipientIdentity is required when ForwardingMode is Recipient.'
        }

        $deliverRaw = Get-TrimmedValue -Value $row.DeliverToMailboxAndForward
        $deliverIsSpecified = -not [string]::IsNullOrWhiteSpace($deliverRaw)
        $desiredDeliver = $null
        if ($deliverIsSpecified) {
            $desiredDeliver = ConvertTo-Bool -Value $deliverRaw
        }
        elseif ($forwardingMode -eq 'clear') {
            $desiredDeliver = $false
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup user mailbox $mailboxIdentity" -ScriptBlock {
            Get-ExchangeOnlineMailbox -Identity $mailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox -or (Get-TrimmedValue -Value $mailbox.RecipientTypeDetails) -ne 'UserMailbox') {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetUserMailboxForwarding' -Status 'NotFound' -Message 'User mailbox not found.'))
            $rowNumber++
            continue
        }

        $expectedPrimarySmtpAddress = ConvertTo-NormalizedSmtpAddress -Value (Get-TrimmedValue -Value $row.ExpectedPrimarySmtpAddress)
        if (-not [string]::IsNullOrWhiteSpace($expectedPrimarySmtpAddress)) {
            $currentPrimary = ConvertTo-NormalizedSmtpAddress -Value (Get-TrimmedValue -Value $mailbox.PrimarySmtpAddress)
            if ($currentPrimary -ne $expectedPrimarySmtpAddress) {
                throw "ExpectedPrimarySmtpAddress mismatch. Current '$currentPrimary' does not match expected '$expectedPrimarySmtpAddress'."
            }
        }

        $currentForwardingSmtpAddress = ConvertTo-NormalizedSmtpAddress -Value (Get-TrimmedValue -Value $mailbox.ForwardingSmtpAddress)
        $currentForwardingAddressRaw = Get-TrimmedValue -Value $mailbox.ForwardingAddress

        $currentForwardingRecipient = $null
        $currentForwardingRecipientPrimary = ''
        $currentForwardingRecipientDn = ''
        if (-not [string]::IsNullOrWhiteSpace($currentForwardingAddressRaw)) {
            $currentForwardingRecipient = Resolve-RecipientByIdentity -Identity $currentForwardingAddressRaw -OperationName "Resolve current forwarding recipient for $mailboxIdentity"
            if ($currentForwardingRecipient) {
                $currentForwardingRecipientPrimary = ConvertTo-NormalizedSmtpAddress -Value (Get-TrimmedValue -Value $currentForwardingRecipient.PrimarySmtpAddress)
                $currentForwardingRecipientDn = Get-TrimmedValue -Value $currentForwardingRecipient.DistinguishedName
            }
        }

        $setParams = @{
            Identity = $mailbox.Identity
        }
        $warnings = [System.Collections.Generic.List[string]]::new()

        switch ($forwardingMode) {
            'smtp' {
                $desiredForwardingSmtpAddress = ConvertTo-NormalizedSmtpAddress -Value $forwardingSmtpAddress
                if ($currentForwardingSmtpAddress -ne $desiredForwardingSmtpAddress) {
                    if ($supports.ForwardingSmtpAddress) {
                        $setParams.ForwardingSmtpAddress = $forwardingSmtpAddress
                    }
                    else {
                        $warnings.Add('ForwardingSmtpAddress ignored (unsupported parameter).')
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($currentForwardingAddressRaw)) {
                    if ($supports.ForwardingAddress) {
                        $setParams.ForwardingAddress = $null
                    }
                    else {
                        $warnings.Add('ForwardingAddress clear ignored (unsupported parameter).')
                    }
                }
            }
            'recipient' {
                $targetRecipient = Resolve-RecipientByIdentity -Identity $forwardingRecipientIdentity -OperationName "Resolve target forwarding recipient $forwardingRecipientIdentity"
                if (-not $targetRecipient) {
                    throw "Forwarding recipient '$forwardingRecipientIdentity' was not found."
                }

                $targetRecipientPrimary = ConvertTo-NormalizedSmtpAddress -Value (Get-TrimmedValue -Value $targetRecipient.PrimarySmtpAddress)
                $targetRecipientDn = Get-TrimmedValue -Value $targetRecipient.DistinguishedName

                $recipientMatches = $false
                if (-not [string]::IsNullOrWhiteSpace($currentForwardingRecipientPrimary) -and -not [string]::IsNullOrWhiteSpace($targetRecipientPrimary) -and $currentForwardingRecipientPrimary -eq $targetRecipientPrimary) {
                    $recipientMatches = $true
                }
                elseif (-not [string]::IsNullOrWhiteSpace($currentForwardingRecipientDn) -and -not [string]::IsNullOrWhiteSpace($targetRecipientDn) -and $currentForwardingRecipientDn -eq $targetRecipientDn) {
                    $recipientMatches = $true
                }
                elseif (-not [string]::IsNullOrWhiteSpace($currentForwardingAddressRaw) -and -not [string]::IsNullOrWhiteSpace($targetRecipientDn) -and $currentForwardingAddressRaw -eq $targetRecipientDn) {
                    $recipientMatches = $true
                }

                if (-not $recipientMatches) {
                    if ($supports.ForwardingAddress) {
                        $setParams.ForwardingAddress = $targetRecipient.Identity
                    }
                    else {
                        $warnings.Add('ForwardingAddress ignored (unsupported parameter).')
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($currentForwardingSmtpAddress)) {
                    if ($supports.ForwardingSmtpAddress) {
                        $setParams.ForwardingSmtpAddress = $null
                    }
                    else {
                        $warnings.Add('ForwardingSmtpAddress clear ignored (unsupported parameter).')
                    }
                }
            }
            'clear' {
                if (-not [string]::IsNullOrWhiteSpace($currentForwardingAddressRaw)) {
                    if ($supports.ForwardingAddress) {
                        $setParams.ForwardingAddress = $null
                    }
                    else {
                        $warnings.Add('ForwardingAddress clear ignored (unsupported parameter).')
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($currentForwardingSmtpAddress)) {
                    if ($supports.ForwardingSmtpAddress) {
                        $setParams.ForwardingSmtpAddress = $null
                    }
                    else {
                        $warnings.Add('ForwardingSmtpAddress clear ignored (unsupported parameter).')
                    }
                }
            }
        }

        if ($null -ne $desiredDeliver) {
            $currentDeliver = [bool]$mailbox.DeliverToMailboxAndForward
            if ($currentDeliver -ne [bool]$desiredDeliver) {
                if ($supports.DeliverToMailboxAndForward) {
                    $setParams.DeliverToMailboxAndForward = [bool]$desiredDeliver
                }
                else {
                    $warnings.Add('DeliverToMailboxAndForward ignored (unsupported parameter).')
                }
            }
        }

        if ($setParams.Count -eq 1) {
            $message = 'No forwarding updates required.'
            if ($warnings.Count -gt 0) {
                $message = "$message $($warnings -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetUserMailboxForwarding' -Status 'Skipped' -Message $message))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($mailboxIdentity, 'Set Exchange Online user mailbox forwarding')) {
            Invoke-WithRetry -OperationName "Set mailbox forwarding for $mailboxIdentity" -ScriptBlock {
                Set-Mailbox @setParams -ErrorAction Stop
            }

            $message = 'User mailbox forwarding updated successfully.'
            if ($warnings.Count -gt 0) {
                $message = "$message $($warnings -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetUserMailboxForwarding' -Status 'Updated' -Message $message))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetUserMailboxForwarding' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetUserMailboxForwarding' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online user mailbox forwarding script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
