<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260305-081500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)

.SYNOPSIS
    Modifies ExchangeOnPremMailboxToMailEnabledUser in Active Directory.

.DESCRIPTION
    Updates ExchangeOnPremMailboxToMailEnabledUser in Active Directory based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M0225-Convert-ExchangeOnPremMailboxToMailEnabledUser.ps1 -InputCsvPath .\0225.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M0225-Convert-ExchangeOnPremMailboxToMailEnabledUser.ps1 -InputCsvPath .\0225.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Exchange Management Shell cmdlets (session)
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    Column                          Type      Required  Description
    ------------------------------  ----      --------  -----------
    MailboxIdentity                 String    Yes       <fill in description>
    ExternalEmailAddress            String    Yes       <fill in description>
    PreserveLegacyExchangeDnAsX500  String    Yes       <fill in description>
    PreserveExistingProxyAddresses  String    Yes       <fill in description>
    DisableEmailAddressPolicy       String    Yes       <fill in description>
    ExpectedPrimarySmtpAddress      String    Yes       <fill in description>
    TargetAlias                     String    Yes       <fill in description>
    Notes                           String    Yes       <fill in description>
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M0225-Convert-ExchangeOnPremMailboxToMailEnabledUser_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function ConvertTo-NormalizedSmtpAddress {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    if ($text.StartsWith('SMTP:', [System.StringComparison]::OrdinalIgnoreCase)) {
        $text = $text.Substring(5)
    }

    return $text.ToLowerInvariant()
}

function Test-IsMailboxRecipientType {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$RecipientTypeDetails
    )

    $type = Get-TrimmedValue -Value $RecipientTypeDetails
    return ($type -in @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox'))
}

function Get-FirstNonEmptyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Values
    )

    foreach ($value in $Values) {
        $text = Get-TrimmedValue -Value $value
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            return $text
        }
    }

    return ''
}

$requiredHeaders = @(
    'MailboxIdentity',
    'ExternalEmailAddress',
    'PreserveLegacyExchangeDnAsX500',
    'PreserveExistingProxyAddresses',
    'DisableEmailAddressPolicy',
    'ExpectedPrimarySmtpAddress',
    'TargetAlias',
    'Notes'
)

Write-Status -Message 'Starting Exchange on-prem mailbox-to-mail-user conversion script.'
Ensure-ExchangeOnPremConnection

$setMailUserCommand = Get-Command -Name Set-MailUser -ErrorAction Stop
$supports = @{
    EmailAddresses           = $setMailUserCommand.Parameters.ContainsKey('EmailAddresses')
    ExternalEmailAddress     = $setMailUserCommand.Parameters.ContainsKey('ExternalEmailAddress')
    EmailAddressPolicyEnabled = $setMailUserCommand.Parameters.ContainsKey('EmailAddressPolicyEnabled')
    Alias                    = $setMailUserCommand.Parameters.ContainsKey('Alias')
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

        $externalEmailAddress = Get-TrimmedValue -Value $row.ExternalEmailAddress
        if ([string]::IsNullOrWhiteSpace($externalEmailAddress)) {
            throw 'ExternalEmailAddress is required.'
        }

        $preserveLegacyExchangeDn = ConvertTo-Bool -Value (Get-TrimmedValue -Value $row.PreserveLegacyExchangeDnAsX500) -Default $true
        $preserveExistingProxyAddresses = ConvertTo-Bool -Value (Get-TrimmedValue -Value $row.PreserveExistingProxyAddresses) -Default $true
        $disableEmailAddressPolicy = ConvertTo-Bool -Value (Get-TrimmedValue -Value $row.DisableEmailAddressPolicy) -Default $true
        $expectedPrimarySmtpAddress = ConvertTo-NormalizedSmtpAddress -Value (Get-TrimmedValue -Value $row.ExpectedPrimarySmtpAddress)
        $targetAlias = Get-TrimmedValue -Value $row.TargetAlias

        $recipient = Invoke-WithRetry -OperationName "Lookup recipient $mailboxIdentity" -ScriptBlock {
            Get-Recipient -Identity $mailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $recipient) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'ConvertMailboxToMailUser' -Status 'NotFound' -Message 'Mailbox or mail user was not found.'))
            $rowNumber++
            continue
        }

        $recipientTypeDetails = Get-TrimmedValue -Value $recipient.RecipientTypeDetails

        if ($recipientTypeDetails -eq 'MailUser') {
            $mailUser = Invoke-WithRetry -OperationName "Lookup mail user $mailboxIdentity" -ScriptBlock {
                Get-MailUser -Identity $mailboxIdentity -ErrorAction SilentlyContinue
            }

            if (-not $mailUser) {
                throw "Recipient '$mailboxIdentity' is already MailUser, but Get-MailUser returned no object."
            }

            $setParams = @{
                Identity = $mailUser.Identity
            }

            $desiredExternalNormalized = ConvertTo-NormalizedSmtpAddress -Value $externalEmailAddress
            $currentExternalNormalized = ConvertTo-NormalizedSmtpAddress -Value (Get-TrimmedValue -Value $mailUser.ExternalEmailAddress)
            if ($currentExternalNormalized -ne $desiredExternalNormalized -and $supports.ExternalEmailAddress) {
                $setParams.ExternalEmailAddress = $externalEmailAddress
            }

            if (-not [string]::IsNullOrWhiteSpace($targetAlias) -and $supports.Alias) {
                $currentAlias = Get-TrimmedValue -Value $mailUser.Alias
                if (-not $currentAlias.Equals($targetAlias, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $setParams.Alias = $targetAlias
                }
            }

            if ($supports.EmailAddressPolicyEnabled) {
                $desiredPolicyEnabled = -not $disableEmailAddressPolicy
                if ([bool]$mailUser.EmailAddressPolicyEnabled -ne [bool]$desiredPolicyEnabled) {
                    $setParams.EmailAddressPolicyEnabled = [bool]$desiredPolicyEnabled
                }
            }

            if ($setParams.Count -eq 1) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'ConvertMailboxToMailUser' -Status 'Skipped' -Message 'Recipient is already a mail-enabled user with requested settings.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess($mailboxIdentity, 'Update existing mail-enabled user settings')) {
                Invoke-WithRetry -OperationName "Update mail user $mailboxIdentity" -ScriptBlock {
                    Set-MailUser @setParams -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'ConvertMailboxToMailUser' -Status 'Updated' -Message 'Existing mail-enabled user updated.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'ConvertMailboxToMailUser' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
            }

            $rowNumber++
            continue
        }

        if (-not (Test-IsMailboxRecipientType -RecipientTypeDetails $recipientTypeDetails)) {
            throw "Recipient '$mailboxIdentity' is '$recipientTypeDetails'. Expected mailbox recipient type or MailUser."
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $mailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            throw "Mailbox '$mailboxIdentity' was not found."
        }

        if (-not [string]::IsNullOrWhiteSpace($expectedPrimarySmtpAddress)) {
            $currentPrimarySmtp = ConvertTo-NormalizedSmtpAddress -Value (Get-TrimmedValue -Value $mailbox.PrimarySmtpAddress)
            if ($currentPrimarySmtp -ne $expectedPrimarySmtpAddress) {
                throw "ExpectedPrimarySmtpAddress mismatch. Current '$currentPrimarySmtp' does not match expected '$expectedPrimarySmtpAddress'."
            }
        }

        $emailAddressSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        if ($preserveExistingProxyAddresses) {
            foreach ($address in @($mailbox.EmailAddresses)) {
                $addressText = Get-TrimmedValue -Value $address
                if (-not [string]::IsNullOrWhiteSpace($addressText)) {
                    $null = $emailAddressSet.Add($addressText)
                }
            }
        }

        if ($preserveLegacyExchangeDn) {
            $legacyExchangeDn = Get-TrimmedValue -Value $mailbox.LegacyExchangeDn
            if (-not [string]::IsNullOrWhiteSpace($legacyExchangeDn)) {
                $null = $emailAddressSet.Add("X500:$legacyExchangeDn")
            }
        }

        $conversionIdentity = Get-FirstNonEmptyValue -Values @(
            $mailbox.SamAccountName,
            $mailbox.Alias,
            $mailbox.UserPrincipalName,
            $mailbox.Identity
        )

        if ([string]::IsNullOrWhiteSpace($conversionIdentity)) {
            throw "Unable to resolve a usable identity for mailbox '$mailboxIdentity'."
        }

        if ($PSCmdlet.ShouldProcess($mailboxIdentity, 'Convert mailbox to mail-enabled user')) {
            Invoke-WithRetry -OperationName "Disable mailbox $mailboxIdentity" -ScriptBlock {
                Disable-Mailbox -Identity $mailbox.Identity -Confirm:$false -ErrorAction Stop
            }

            Invoke-WithRetry -OperationName "Enable mail user $mailboxIdentity" -ScriptBlock {
                Enable-MailUser -Identity $conversionIdentity -ExternalEmailAddress $externalEmailAddress -ErrorAction Stop
            }

            $setParams = @{
                Identity = $conversionIdentity
            }

            if ($supports.ExternalEmailAddress) {
                $setParams.ExternalEmailAddress = $externalEmailAddress
            }

            if ($supports.EmailAddressPolicyEnabled) {
                $setParams.EmailAddressPolicyEnabled = -not $disableEmailAddressPolicy
            }

            if (-not [string]::IsNullOrWhiteSpace($targetAlias) -and $supports.Alias) {
                $setParams.Alias = $targetAlias
            }

            if ($supports.EmailAddresses -and $emailAddressSet.Count -gt 0) {
                $setParams.EmailAddresses = @($emailAddressSet)
            }

            if ($setParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Finalize mail user settings $mailboxIdentity" -ScriptBlock {
                    Set-MailUser @setParams -ErrorAction Stop
                }
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'ConvertMailboxToMailUser' -Status 'Converted' -Message 'Mailbox converted to mail-enabled user.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'ConvertMailboxToMailUser' -Status 'WhatIf' -Message 'Conversion skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'ConvertMailboxToMailUser' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem mailbox-to-mail-user conversion script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
