<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-173000

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR3129-Get-ExchangeOnlineMailboxPermissionsConsolidated_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-TrimmedValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    return ([string]$Value).Trim()
}

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
        [System.Collections.Generic.HashSet[string]]$Set
    )

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

$requiredHeaders = @(
    'MailboxIdentity'
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
            TrusteeIdentity           = Get-TrimmedValue -Value $recipient.Identity
            TrusteePrimarySmtpAddress = Get-TrimmedValue -Value $recipient.PrimarySmtpAddress
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
                Get-ExchangeOnlineMailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited -ErrorAction Stop
            })
        }
        else {
            $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentityInput" -ScriptBlock {
                Get-ExchangeOnlineMailbox -Identity $mailboxIdentityInput -ErrorAction SilentlyContinue
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
            $mailboxIdentityResolved = Get-TrimmedValue -Value $mailbox.Identity

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
                Get-ExchangeOnlineMailboxPermission -Identity $mailboxIdentityResolved -ErrorAction Stop
            })

            foreach ($permission in $mailboxPermissions) {
                if ($permission.Deny -eq $true) { continue }
                if ($permission.IsInherited -eq $true) { continue }

                $trustee = Get-TrimmedValue -Value $permission.User
                if ([string]::IsNullOrWhiteSpace($trustee)) { continue }
                if ($trustee.Equals('NT AUTHORITY\SELF', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                if ($trustee -match '^S-1-5-') { continue }

                $accessRights = @($permission.AccessRights | ForEach-Object { Get-TrimmedValue -Value $_ })
                if ($accessRights -contains 'FullAccess') {
                    & $addTrustee -TrusteeHint $trustee -PermissionFamily 'FullAccess'
                }
                if ($accessRights -contains 'ReadPermission') {
                    & $addTrustee -TrusteeHint $trustee -PermissionFamily 'ReadOnly'
                }
            }

            $recipientPermissions = @(Invoke-WithRetry -OperationName "Load recipient permissions $mailboxIdentityResolved" -ScriptBlock {
                Get-ExchangeOnlineRecipientPermission -Identity $mailboxIdentityResolved -ErrorAction SilentlyContinue
            })

            foreach ($permission in $recipientPermissions) {
                if ($permission.Deny -eq $true) { continue }

                $accessRights = @($permission.AccessRights | ForEach-Object { Get-TrimmedValue -Value $_ })
                if ($accessRights -notcontains 'SendAs') { continue }

                $trustee = Get-TrimmedValue -Value $permission.Trustee
                if ([string]::IsNullOrWhiteSpace($trustee)) { continue }

                & $addTrustee -TrusteeHint $trustee -PermissionFamily 'SendAs'
            }

            foreach ($delegate in @($mailbox.GrantSendOnBehalfTo)) {
                $delegateHint = Get-TrimmedValue -Value $delegate.DistinguishedName
                if ([string]::IsNullOrWhiteSpace($delegateHint)) {
                    $delegateHint = Get-TrimmedValue -Value $delegate.Name
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

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityResolved -Action 'GetExchangeMailboxPermissionsConsolidated' -Status 'Completed' -Message $message -Data ([ordered]@{
                        MailboxIdentity                          = $mailboxIdentityResolved
                        MailboxRecipientTypeDetails              = Get-TrimmedValue -Value $mailbox.RecipientTypeDetails
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

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online consolidated mailbox-permission inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
