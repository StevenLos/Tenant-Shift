<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-170500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineSharedMailboxSmtpAddresses and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineSharedMailboxSmtpAddresses from Microsoft 365 and writes the results to a CSV file.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all objects in scope (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all objects in scope rather than processing from an input CSV file. Uses the DiscoverAll parameter set.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-IR3132-Get-ExchangeOnlineSharedMailboxSmtpAddresses.ps1 -InputCsvPath .\3132.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3132-Get-ExchangeOnlineSharedMailboxSmtpAddresses.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator
    Limitations:      None known.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-EXOL-0270-Get-ExchangeOnlineSharedMailboxSmtpAddresses_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-ObjectPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($property) {
        return $property.Value
    }

    return $null
}

function Get-StringPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    return Get-TrimmedValue -Value (Get-ObjectPropertyValue -InputObject $InputObject -PropertyName $PropertyName)
}

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

    return $text.Trim().ToLowerInvariant()
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

function Get-SmtpAddressEntries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox
    )

    $entries = [System.Collections.Generic.List[object]]::new()
    $seenAddresses = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $primarySmtpAddress = ConvertTo-NormalizedSmtpAddress -Value (Get-ObjectPropertyValue -InputObject $Mailbox -PropertyName 'PrimarySmtpAddress')
    $emailAddresses = @(Get-ObjectPropertyValue -InputObject $Mailbox -PropertyName 'EmailAddresses')

    foreach ($rawAddress in $emailAddresses) {
        $rawText = Get-TrimmedValue -Value $rawAddress
        if ([string]::IsNullOrWhiteSpace($rawText)) {
            continue
        }

        if (-not $rawText.StartsWith('SMTP:', [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        $smtpAddress = ConvertTo-NormalizedSmtpAddress -Value $rawText
        if ([string]::IsNullOrWhiteSpace($smtpAddress)) {
            continue
        }

        if (-not $seenAddresses.Add($smtpAddress)) {
            continue
        }

        $entries.Add([PSCustomObject]@{
                EmailAddress    = $smtpAddress
                EmailAddressRaw = $rawText
                AddressRole     = if ($smtpAddress -eq $primarySmtpAddress) { 'Primary' } else { 'Secondary' }
                AddressSource   = 'EmailAddresses'
            })
    }

    if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress) -and -not $seenAddresses.Contains($primarySmtpAddress)) {
        $entries.Add([PSCustomObject]@{
                EmailAddress    = $primarySmtpAddress
                EmailAddressRaw = "SMTP:$primarySmtpAddress"
                AddressRole     = 'Primary'
                AddressSource   = 'PrimarySmtpAddressFallback'
            })
    }

    return @(
        $entries |
            Sort-Object -Property @{ Expression = { if ($_.AddressRole -eq 'Primary') { 0 } else { 1 } } }, @{ Expression = { $_.EmailAddress } }
    )
}

$requiredHeaders = @(
    'SharedMailboxIdentity'
)

$mailboxProperties = @(
    'DisplayName',
    'Alias',
    'UserPrincipalName',
    'PrimarySmtpAddress',
    'EmailAddresses',
    'RecipientTypeDetails',
    'HiddenFromAddressListsEnabled',
    'WhenCreatedUTC'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'SharedMailboxIdentity',
    'DisplayName',
    'Alias',
    'UserPrincipalName',
    'RecipientTypeDetails',
    'PrimarySmtpAddress',
    'EmailAddress',
    'EmailAddressRaw',
    'AddressRole',
    'AddressSource',
    'AddressOrdinal',
    'TotalSmtpAddressCount',
    'HiddenFromAddressListsEnabled',
    'WhenCreatedUTC'
)

Write-Status -Message 'Starting Exchange Online shared mailbox SMTP-address inventory script.'
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
    $sharedMailboxIdentityInput = Get-TrimmedValue -Value $row.SharedMailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($sharedMailboxIdentityInput)) {
            throw 'SharedMailboxIdentity is required. Use * to inventory SMTP addresses for all shared mailboxes.'
        }

        $mailboxes = @()
        if ($sharedMailboxIdentityInput -eq '*') {
            $mailboxes = @(Invoke-WithRetry -OperationName 'Load all shared mailboxes for SMTP-address inventory' -ScriptBlock {
                Get-ExchangeOnlineMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -Properties $mailboxProperties -ErrorAction Stop
            })
        }
        else {
            $mailbox = Invoke-WithRetry -OperationName "Lookup shared mailbox $sharedMailboxIdentityInput" -ScriptBlock {
                Get-ExchangeOnlineMailbox -Identity $sharedMailboxIdentityInput -Properties $mailboxProperties -ErrorAction SilentlyContinue
            }

            if ($mailbox -and (Get-StringPropertyValue -InputObject $mailbox -PropertyName 'RecipientTypeDetails') -eq 'SharedMailbox') {
                $mailboxes = @($mailbox)
            }
        }

        if ($mailboxes.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentityInput -Action 'GetExchangeSharedMailboxSmtpAddress' -Status 'NotFound' -Message 'No matching shared mailboxes were found.' -Data ([ordered]@{
                            SharedMailboxIdentity          = $sharedMailboxIdentityInput
                            DisplayName                    = ''
                            Alias                          = ''
                            UserPrincipalName              = ''
                            RecipientTypeDetails           = ''
                            PrimarySmtpAddress             = ''
                            EmailAddress                   = ''
                            EmailAddressRaw                = ''
                            AddressRole                    = ''
                            AddressSource                  = ''
                            AddressOrdinal                 = ''
                            TotalSmtpAddressCount          = ''
                            HiddenFromAddressListsEnabled  = ''
                            WhenCreatedUTC                 = ''
                        })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $sharedMailboxIdentityResolved = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'Identity'
            $smtpAddressEntries = @(Get-SmtpAddressEntries -Mailbox $mailbox)

            if ($smtpAddressEntries.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentityResolved -Action 'GetExchangeSharedMailboxSmtpAddress' -Status 'Completed' -Message 'Shared mailbox exported. No SMTP addresses were returned.' -Data ([ordered]@{
                                SharedMailboxIdentity          = $sharedMailboxIdentityResolved
                                DisplayName                    = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'DisplayName'
                                Alias                          = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'Alias'
                                UserPrincipalName              = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'UserPrincipalName'
                                RecipientTypeDetails           = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'RecipientTypeDetails'
                                PrimarySmtpAddress             = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'PrimarySmtpAddress'
                                EmailAddress                   = ''
                                EmailAddressRaw                = ''
                                AddressRole                    = ''
                                AddressSource                  = ''
                                AddressOrdinal                 = ''
                                TotalSmtpAddressCount          = '0'
                                HiddenFromAddressListsEnabled  = [string](Get-ObjectPropertyValue -InputObject $mailbox -PropertyName 'HiddenFromAddressListsEnabled')
                                WhenCreatedUTC                 = [string](Get-ObjectPropertyValue -InputObject $mailbox -PropertyName 'WhenCreatedUTC')
                            })))
                continue
            }

            $totalSmtpAddressCount = [string]$smtpAddressEntries.Count
            for ($addressIndex = 0; $addressIndex -lt $smtpAddressEntries.Count; $addressIndex++) {
                $smtpAddressEntry = $smtpAddressEntries[$addressIndex]
                $smtpAddress = Get-TrimmedValue -Value $smtpAddressEntry.EmailAddress
                $primaryKey = if (-not [string]::IsNullOrWhiteSpace($smtpAddress)) {
                    '{0}|{1}' -f $sharedMailboxIdentityResolved, $smtpAddress
                }
                else {
                    $sharedMailboxIdentityResolved
                }

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetExchangeSharedMailboxSmtpAddress' -Status 'Completed' -Message 'Shared mailbox SMTP address exported.' -Data ([ordered]@{
                                SharedMailboxIdentity          = $sharedMailboxIdentityResolved
                                DisplayName                    = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'DisplayName'
                                Alias                          = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'Alias'
                                UserPrincipalName              = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'UserPrincipalName'
                                RecipientTypeDetails           = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'RecipientTypeDetails'
                                PrimarySmtpAddress             = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'PrimarySmtpAddress'
                                EmailAddress                   = $smtpAddress
                                EmailAddressRaw                = Get-TrimmedValue -Value $smtpAddressEntry.EmailAddressRaw
                                AddressRole                    = Get-TrimmedValue -Value $smtpAddressEntry.AddressRole
                                AddressSource                  = Get-TrimmedValue -Value $smtpAddressEntry.AddressSource
                                AddressOrdinal                 = [string]($addressIndex + 1)
                                TotalSmtpAddressCount          = $totalSmtpAddressCount
                                HiddenFromAddressListsEnabled  = [string](Get-ObjectPropertyValue -InputObject $mailbox -PropertyName 'HiddenFromAddressListsEnabled')
                                WhenCreatedUTC                 = [string](Get-ObjectPropertyValue -InputObject $mailbox -PropertyName 'WhenCreatedUTC')
                            })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($sharedMailboxIdentityInput) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $sharedMailboxIdentityInput -Action 'GetExchangeSharedMailboxSmtpAddress' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                        SharedMailboxIdentity          = $sharedMailboxIdentityInput
                        DisplayName                    = ''
                        Alias                          = ''
                        UserPrincipalName              = ''
                        RecipientTypeDetails           = ''
                        PrimarySmtpAddress             = ''
                        EmailAddress                   = ''
                        EmailAddressRaw                = ''
                        AddressRole                    = ''
                        AddressSource                  = ''
                        AddressOrdinal                 = ''
                        TotalSmtpAddressCount          = ''
                        HiddenFromAddressListsEnabled  = ''
                        WhenCreatedUTC                 = ''
                    })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

$orderedResults = foreach ($result in $results) {
    $ordered = [ordered]@{}

    foreach ($propertyName in $reportPropertyOrder) {
        $property = $result.PSObject.Properties[$propertyName]
        if ($property) {
            $ordered[$propertyName] = $property.Value
        }
        else {
            $ordered[$propertyName] = ''
        }
    }

    foreach ($property in $result.PSObject.Properties) {
        if (-not $ordered.Contains($property.Name)) {
            $ordered[$property.Name] = $property.Value
        }
    }

    [PSCustomObject]$ordered
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online shared mailbox SMTP-address inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
