<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-220000

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)

.SYNOPSIS
    Provisions ExchangeOnPremSharedMailboxes in Active Directory.

.DESCRIPTION
    Creates ExchangeOnPremSharedMailboxes in Active Directory based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P0216-Create-ExchangeOnPremSharedMailboxes.ps1 -InputCsvPath .\0216.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P0216-Create-ExchangeOnPremSharedMailboxes.ps1 -InputCsvPath .\0216.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Exchange Management Shell cmdlets (session)
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    Column                             Type      Required  Description
    ---------------------------------  ----      --------  -----------
    Name                               String    Yes       <fill in description>
    Alias                              String    Yes       <fill in description>
    DisplayName                        String    Yes       <fill in description>
    PrimarySmtpAddress                 String    Yes       <fill in description>
    HiddenFromAddressListsEnabled      String    Yes       <fill in description>
    MessageCopyForSentAsEnabled        String    Yes       <fill in description>
    MessageCopyForSendOnBehalfEnabled  String    Yes       <fill in description>
    MailTip                            String    Yes       <fill in description>
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P0216-Create-ExchangeOnPremSharedMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

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
    'Alias',
    'DisplayName',
    'PrimarySmtpAddress',
    'HiddenFromAddressListsEnabled',
    'MessageCopyForSentAsEnabled',
    'MessageCopyForSendOnBehalfEnabled',
    'MailTip'
)

Write-Status -Message 'Starting Exchange on-prem shared mailbox creation script.'
Ensure-ExchangeOnPremConnection

$newMailboxCommand = Get-Command -Name New-Mailbox -ErrorAction Stop
$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop

$supports = @{
    OrganizationalUnit                  = $newMailboxCommand.Parameters.ContainsKey('OrganizationalUnit')
    HiddenFromAddressListsEnabled       = $setMailboxCommand.Parameters.ContainsKey('HiddenFromAddressListsEnabled')
    MessageCopyForSentAsEnabled         = $setMailboxCommand.Parameters.ContainsKey('MessageCopyForSentAsEnabled')
    MessageCopyForSendOnBehalfEnabled   = $setMailboxCommand.Parameters.ContainsKey('MessageCopyForSendOnBehalfEnabled')
    MailTip                             = $setMailboxCommand.Parameters.ContainsKey('MailTip')
}

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $name = Get-TrimmedValue -Value $row.Name

    try {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Name is required.'
        }

        $alias = Get-TrimmedValue -Value $row.Alias
        if ([string]::IsNullOrWhiteSpace($alias)) {
            $alias = ($name -replace '[^a-zA-Z0-9]', '')
            if ([string]::IsNullOrWhiteSpace($alias)) {
                throw 'Alias is empty after sanitization. Provide an Alias value in the CSV.'
            }
        }

        $primarySmtpAddress = Get-TrimmedValue -Value $row.PrimarySmtpAddress
        $identityToCheck = if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) { $primarySmtpAddress } else { $name }

        $existingMailbox = Invoke-WithRetry -OperationName "Lookup shared mailbox $identityToCheck" -ScriptBlock {
            Get-Mailbox -Identity $identityToCheck -ErrorAction SilentlyContinue
        }

        if ($existingMailbox) {
            $recipientTypeDetails = Get-TrimmedValue -Value $existingMailbox.RecipientTypeDetails
            if ($recipientTypeDetails -eq 'SharedMailbox') {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateSharedMailbox' -Status 'Skipped' -Message 'Shared mailbox already exists.'))
                $rowNumber++
                continue
            }

            throw "Recipient '$identityToCheck' already exists with type '$recipientTypeDetails'."
        }

        $createParams = @{
            Name  = $name
            Alias = $alias
            Shared = $true
        }

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $createParams.DisplayName = $displayName
        }

        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $createParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $organizationalUnit = Get-OptionalColumnValue -Row $row -ColumnName 'OrganizationalUnit'
        if ($supports.OrganizationalUnit -and -not [string]::IsNullOrWhiteSpace($organizationalUnit)) {
            $createParams.OrganizationalUnit = $organizationalUnit
        }

        if ($PSCmdlet.ShouldProcess($identityToCheck, 'Create Exchange on-prem shared mailbox')) {
            $createdMailbox = Invoke-WithRetry -OperationName "Create shared mailbox $identityToCheck" -ScriptBlock {
                New-Mailbox @createParams -ErrorAction Stop
            }

            $setParams = @{ Identity = $createdMailbox.Identity }
            $warnings = [System.Collections.Generic.List[string]]::new()

            $hiddenRaw = Get-TrimmedValue -Value $row.HiddenFromAddressListsEnabled
            if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
                if ($supports.HiddenFromAddressListsEnabled) {
                    $setParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
                }
                else {
                    $warnings.Add('HiddenFromAddressListsEnabled ignored (unsupported parameter).')
                }
            }

            $copySentAsRaw = Get-TrimmedValue -Value $row.MessageCopyForSentAsEnabled
            if (-not [string]::IsNullOrWhiteSpace($copySentAsRaw)) {
                if ($supports.MessageCopyForSentAsEnabled) {
                    $setParams.MessageCopyForSentAsEnabled = ConvertTo-Bool -Value $copySentAsRaw
                }
                else {
                    $warnings.Add('MessageCopyForSentAsEnabled ignored (unsupported parameter).')
                }
            }

            $copySobRaw = Get-TrimmedValue -Value $row.MessageCopyForSendOnBehalfEnabled
            if (-not [string]::IsNullOrWhiteSpace($copySobRaw)) {
                if ($supports.MessageCopyForSendOnBehalfEnabled) {
                    $setParams.MessageCopyForSendOnBehalfEnabled = ConvertTo-Bool -Value $copySobRaw
                }
                else {
                    $warnings.Add('MessageCopyForSendOnBehalfEnabled ignored (unsupported parameter).')
                }
            }

            $mailTip = Get-TrimmedValue -Value $row.MailTip
            if (-not [string]::IsNullOrWhiteSpace($mailTip)) {
                if ($supports.MailTip) {
                    $setParams.MailTip = $mailTip
                }
                else {
                    $warnings.Add('MailTip ignored (unsupported parameter).')
                }
            }

            if ($setParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Set shared mailbox options $identityToCheck" -ScriptBlock {
                    Set-Mailbox @setParams -ErrorAction Stop
                }
            }

            $message = 'Shared mailbox created successfully.'
            if ($warnings.Count -gt 0) {
                $message = "$message $($warnings -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateSharedMailbox' -Status 'Created' -Message $message))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateSharedMailbox' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($name) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $name -Action 'CreateSharedMailbox' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem shared mailbox creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
