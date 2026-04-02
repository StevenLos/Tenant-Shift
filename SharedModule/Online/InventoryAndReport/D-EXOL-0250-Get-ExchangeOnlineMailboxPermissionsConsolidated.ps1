<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-183000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineMailboxPermissionsConsolidated and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineMailboxPermissionsConsolidated from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-IR3129-Get-ExchangeOnlineMailboxPermissionsConsolidated.ps1 -InputCsvPath .\3129.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3129-Get-ExchangeOnlineMailboxPermissionsConsolidated.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-EXOL-0250-Get-ExchangeOnlineMailboxPermissionsConsolidated_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Normalize-TrusteeKey {
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

    return $text.ToLowerInvariant()
}

function Convert-HashSetToSemicolonString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Set
    )

    if ($null -eq $Set -or $Set.Count -eq 0) {
        return ''
    }

    return (@($Set | Sort-Object) -join ';')
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

function Get-MailboxLookupIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox
    )

    $primarySmtpAddress = Get-StringPropertyValue -InputObject $Mailbox -PropertyName 'PrimarySmtpAddress'
    if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
        return $primarySmtpAddress
    }

    return Get-StringPropertyValue -InputObject $Mailbox -PropertyName 'Identity'
}

$requiredHeaders = @(
    'MailboxIdentity'
)

$mailboxProperties = @(
    'GrantSendOnBehalfTo'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'MailboxIdentity',
    'MailboxRecipientTypeDetails',
    'FullAccessTrusteeCount',
    'ReadOnlyTrusteeCount',
    'SendAsTrusteeCount',
    'SendOnBehalfTrusteeCount',
    'AllDelegatedTrusteeCount',
    'FullAccessTrustees',
    'ReadOnlyTrustees',
    'SendAsTrustees',
    'SendOnBehalfTrustees',
    'AllDelegatedTrustees',
    'AllDelegatedTrusteePrimarySmtpAddresses'
)

Write-Status -Message 'Starting Exchange Online consolidated mailbox-permission inventory script.'
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
$recipientSummaryByKey = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

$resolveRecipientSummary = {
    param(
        [Parameter(Mandatory)]
        [string]$IdentityHint
    )

    $normalized = Normalize-TrusteeKey -Value $IdentityHint
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return [PSCustomObject]@{
            TrusteeIdentity           = ''
            TrusteePrimarySmtpAddress = ''
        }
    }

    if ($recipientSummaryByKey.ContainsKey($normalized)) {
        return $recipientSummaryByKey[$normalized]
    }

        $summary = $null
    try {
        $recipient = Invoke-WithRetry -OperationName "Lookup recipient $IdentityHint" -ScriptBlock {
            Get-ExchangeOnlineRecipient -Identity $IdentityHint -ErrorAction Stop
        }

        $summary = [PSCustomObject]@{
            TrusteeIdentity           = Get-StringPropertyValue -InputObject $recipient -PropertyName 'Identity'
            TrusteePrimarySmtpAddress = Get-StringPropertyValue -InputObject $recipient -PropertyName 'PrimarySmtpAddress'
        }
    }
    catch {
        $summary = [PSCustomObject]@{
            TrusteeIdentity           = Get-TrimmedValue -Value $IdentityHint
            TrusteePrimarySmtpAddress = ''
        }
    }

    $recipientSummaryByKey[$normalized] = $summary
    return $summary
}

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentityInput = Get-TrimmedValue -Value $row.MailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentityInput)) {
            throw 'MailboxIdentity is required. Use * to inventory consolidated mailbox permissions for all user/shared mailboxes.'
        }

        $mailboxes = @()
        if ($mailboxIdentityInput -eq '*') {
            $mailboxes = @(Invoke-WithRetry -OperationName 'Load all user/shared mailboxes for consolidated permission inventory' -ScriptBlock {
                Get-ExchangeOnlineMailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited -Properties $mailboxProperties -ErrorAction Stop
            })
        }
        else {
            $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentityInput" -ScriptBlock {
                Get-ExchangeOnlineMailbox -Identity $mailboxIdentityInput -Properties $mailboxProperties -ErrorAction SilentlyContinue
            }

            if ($mailbox) {
                $mailboxes = @($mailbox)
            }
        }

        if ($mailboxes.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityInput -Action 'GetExchangeMailboxPermissionsConsolidated' -Status 'NotFound' -Message 'No matching mailboxes were found.' -Data ([ordered]@{
                        MailboxIdentity                          = $mailboxIdentityInput
                        MailboxRecipientTypeDetails              = ''
                        FullAccessTrusteeCount                   = ''
                        ReadOnlyTrusteeCount                     = ''
                        SendAsTrusteeCount                       = ''
                        SendOnBehalfTrusteeCount                 = ''
                        AllDelegatedTrusteeCount                 = ''
                        FullAccessTrustees                       = ''
                        ReadOnlyTrustees                         = ''
                        SendAsTrustees                           = ''
                        SendOnBehalfTrustees                     = ''
                        AllDelegatedTrustees                     = ''
                        AllDelegatedTrusteePrimarySmtpAddresses  = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $mailboxIdentityResolved = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'Identity'
            $mailboxLookupIdentity = Get-MailboxLookupIdentity -Mailbox $mailbox
            if ([string]::IsNullOrWhiteSpace($mailboxLookupIdentity)) {
                throw 'Unable to resolve a unique mailbox identity for permission lookup.'
            }
            if ([string]::IsNullOrWhiteSpace($mailboxIdentityResolved)) {
                $mailboxIdentityResolved = $mailboxLookupIdentity
            }

            $fullAccessSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $readOnlySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $sendAsSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $sendOnBehalfSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $allDelegatedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $smtpAddressSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

            $addTrustee = {
                param(
                    [Parameter(Mandatory)]
                    [string]$TrusteeHint,

                    [Parameter(Mandatory)]
                    [string]$PermissionFamily
                )

                $summary = & $resolveRecipientSummary -IdentityHint $TrusteeHint
                $trusteeIdentity = Get-TrimmedValue -Value $summary.TrusteeIdentity
                if ([string]::IsNullOrWhiteSpace($trusteeIdentity)) {
                    $trusteeIdentity = Get-TrimmedValue -Value $TrusteeHint
                }

                if ([string]::IsNullOrWhiteSpace($trusteeIdentity)) {
                    return
                }

                [void]$allDelegatedSet.Add($trusteeIdentity)

                $trusteeSmtp = Get-TrimmedValue -Value $summary.TrusteePrimarySmtpAddress
                if (-not [string]::IsNullOrWhiteSpace($trusteeSmtp)) {
                    [void]$smtpAddressSet.Add($trusteeSmtp.ToLowerInvariant())
                }

                switch ($PermissionFamily) {
                    'FullAccess' { [void]$fullAccessSet.Add($trusteeIdentity) }
                    'ReadOnly' { [void]$readOnlySet.Add($trusteeIdentity) }
                    'SendAs' { [void]$sendAsSet.Add($trusteeIdentity) }
                    'SendOnBehalf' { [void]$sendOnBehalfSet.Add($trusteeIdentity) }
                }
            }

            $mailboxPermissions = @(Invoke-WithRetry -OperationName "Load mailbox permissions $mailboxIdentityResolved" -ScriptBlock {
                Get-ExchangeOnlineMailboxPermission -Identity $mailboxLookupIdentity -ErrorAction Stop
            })

            foreach ($permission in $mailboxPermissions) {
                if ((Get-ObjectPropertyValue -InputObject $permission -PropertyName 'Deny') -eq $true) { continue }
                if ((Get-ObjectPropertyValue -InputObject $permission -PropertyName 'IsInherited') -eq $true) { continue }

                $trustee = Get-StringPropertyValue -InputObject $permission -PropertyName 'User'
                if ([string]::IsNullOrWhiteSpace($trustee)) { continue }
                if ($trustee.Equals('NT AUTHORITY\SELF', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                if ($trustee -match '^S-1-5-') { continue }

                $accessRights = @((Get-ObjectPropertyValue -InputObject $permission -PropertyName 'AccessRights') | ForEach-Object { Get-TrimmedValue -Value $_ })
                if ($accessRights -contains 'FullAccess') {
                    & $addTrustee -TrusteeHint $trustee -PermissionFamily 'FullAccess'
                }
                if ($accessRights -contains 'ReadPermission') {
                    & $addTrustee -TrusteeHint $trustee -PermissionFamily 'ReadOnly'
                }
            }

            $recipientPermissions = @(Invoke-WithRetry -OperationName "Load recipient permissions $mailboxIdentityResolved" -ScriptBlock {
                Get-ExchangeOnlineRecipientPermission -Identity $mailboxLookupIdentity -ErrorAction SilentlyContinue
            })

            foreach ($permission in $recipientPermissions) {
                if ((Get-ObjectPropertyValue -InputObject $permission -PropertyName 'Deny') -eq $true) { continue }

                $accessRights = @((Get-ObjectPropertyValue -InputObject $permission -PropertyName 'AccessRights') | ForEach-Object { Get-TrimmedValue -Value $_ })
                if ($accessRights -notcontains 'SendAs') { continue }

                $trustee = Get-StringPropertyValue -InputObject $permission -PropertyName 'Trustee'
                if ([string]::IsNullOrWhiteSpace($trustee)) { continue }

                & $addTrustee -TrusteeHint $trustee -PermissionFamily 'SendAs'
            }

            foreach ($delegate in @((Get-ObjectPropertyValue -InputObject $mailbox -PropertyName 'GrantSendOnBehalfTo'))) {
                $delegateHint = if ($delegate -is [string]) {
                    Get-TrimmedValue -Value $delegate
                }
                else {
                    Get-StringPropertyValue -InputObject $delegate -PropertyName 'DistinguishedName'
                }
                if ([string]::IsNullOrWhiteSpace($delegateHint)) {
                    if ($delegate -is [string]) {
                        $delegateHint = ''
                    }
                    else {
                        $delegateHint = Get-StringPropertyValue -InputObject $delegate -PropertyName 'Name'
                    }
                }

                if ([string]::IsNullOrWhiteSpace($delegateHint)) { continue }

                & $addTrustee -TrusteeHint $delegateHint -PermissionFamily 'SendOnBehalf'
            }

            $message = if ($allDelegatedSet.Count -eq 0) {
                'No explicit delegated permissions found for mailbox.'
            }
            else {
                'Consolidated mailbox permission row exported.'
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxLookupIdentity -Action 'GetExchangeMailboxPermissionsConsolidated' -Status 'Completed' -Message $message -Data ([ordered]@{
                        MailboxIdentity                          = $mailboxIdentityResolved
                        MailboxRecipientTypeDetails              = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'RecipientTypeDetails'
                        FullAccessTrusteeCount                   = [string]$fullAccessSet.Count
                        ReadOnlyTrusteeCount                     = [string]$readOnlySet.Count
                        SendAsTrusteeCount                       = [string]$sendAsSet.Count
                        SendOnBehalfTrusteeCount                 = [string]$sendOnBehalfSet.Count
                        AllDelegatedTrusteeCount                 = [string]$allDelegatedSet.Count
                        FullAccessTrustees                       = Convert-HashSetToSemicolonString -Set $fullAccessSet
                        ReadOnlyTrustees                         = Convert-HashSetToSemicolonString -Set $readOnlySet
                        SendAsTrustees                           = Convert-HashSetToSemicolonString -Set $sendAsSet
                        SendOnBehalfTrustees                     = Convert-HashSetToSemicolonString -Set $sendOnBehalfSet
                        AllDelegatedTrustees                     = Convert-HashSetToSemicolonString -Set $allDelegatedSet
                        AllDelegatedTrusteePrimarySmtpAddresses  = Convert-HashSetToSemicolonString -Set $smtpAddressSet
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentityInput) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityInput -Action 'GetExchangeMailboxPermissionsConsolidated' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    MailboxIdentity                          = $mailboxIdentityInput
                    MailboxRecipientTypeDetails              = ''
                    FullAccessTrusteeCount                   = ''
                    ReadOnlyTrusteeCount                     = ''
                    SendAsTrusteeCount                       = ''
                    SendOnBehalfTrusteeCount                 = ''
                    AllDelegatedTrusteeCount                 = ''
                    FullAccessTrustees                       = ''
                    ReadOnlyTrustees                         = ''
                    SendAsTrustees                           = ''
                    SendOnBehalfTrustees                     = ''
                    AllDelegatedTrustees                     = ''
                    AllDelegatedTrusteePrimarySmtpAddresses  = ''
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
Write-Status -Message 'Exchange Online consolidated mailbox-permission inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
