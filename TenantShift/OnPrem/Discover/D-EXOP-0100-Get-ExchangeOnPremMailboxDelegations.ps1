<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-235500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)

.SYNOPSIS
    Gets ExchangeOnPremMailboxDelegations and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnPremMailboxDelegations from Active Directory and writes the results to a CSV file.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all objects in scope (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all objects in scope rather than processing from an input CSV file. Uses the DiscoverAll parameter set.

.PARAMETER SearchBase
    Distinguished name of the Active Directory OU to scope the discovery. If omitted, searches the entire domain.

.PARAMETER Server
    Active Directory domain controller to target. If omitted, uses the default DC for the current domain.

.PARAMETER MaxObjects
    Maximum number of objects to retrieve. 0 (default) means no limit.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-D0220-Get-ExchangeOnPremMailboxDelegations.ps1 -InputCsvPath .\0220.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D0220-Get-ExchangeOnPremMailboxDelegations.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Exchange Management Shell cmdlets (session)
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$SearchBase,

    [Parameter(ParameterSetName = 'FromCsv')]
    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$Server,

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [ValidateRange(0, 10000000)]
    [int]$MaxObjects = 0,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_SM-D0220-Get-ExchangeOnPremMailboxDelegations_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Add-SupportedParameter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$ParameterHashtable,

        [Parameter(Mandatory)]
        [string]$CommandName,

        [Parameter(Mandatory)]
        [string]$ParameterName,

        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return
    }

    $command = Get-Command -Name $CommandName -ErrorAction Stop
    if ($command.Parameters.ContainsKey($ParameterName)) {
        $ParameterHashtable[$ParameterName] = $text
    }
}

function Resolve-MailboxesByScope {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,

        [AllowEmptyString()]
        [string]$SearchBase,

        [AllowEmptyString()]
        [string]$Server
    )

    if ($Identity -eq '*') {
        $params = @{
            RecipientTypeDetails = @('UserMailbox', 'SharedMailbox')
            ResultSize           = 'Unlimited'
            ErrorAction          = 'Stop'
        }

        Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-Mailbox' -ParameterName 'OrganizationalUnit' -Value $SearchBase
        Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-Mailbox' -ParameterName 'DomainController' -Value $Server

        return @(Get-Mailbox @params)
    }

    $params = @{
        Identity    = $Identity
        ErrorAction = 'SilentlyContinue'
    }

    Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-Mailbox' -ParameterName 'DomainController' -Value $Server

    $mailbox = Get-Mailbox @params
    if ($mailbox -and (Get-TrimmedValue -Value $mailbox.RecipientTypeDetails) -in @('UserMailbox', 'SharedMailbox')) {
        return @($mailbox)
    }

    return @()
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

function Get-DelegateHint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$DelegateObject
    )

    $candidateProperties = @('DistinguishedName', 'Name', 'PrimarySmtpAddress', 'Identity')
    foreach ($propertyName in $candidateProperties) {
        $property = $DelegateObject.PSObject.Properties[$propertyName]
        if ($property) {
            $text = Get-TrimmedValue -Value $property.Value
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                return $text
            }
        }
    }

    return Get-TrimmedValue -Value $DelegateObject
}

$requiredHeaders = @(
    'MailboxIdentity'
)

Write-Status -Message 'Starting Exchange on-prem mailbox delegation inventory script.'
Ensure-ExchangeOnPremConnection

$hasRecipientPermissionCmdlet = [bool](Get-Command -Name Get-RecipientPermission -ErrorAction SilentlyContinue)
$hasAdPermissionCmdlet = [bool](Get-Command -Name Get-ADPermission -ErrorAction SilentlyContinue)
if (-not $hasRecipientPermissionCmdlet -and -not $hasAdPermissionCmdlet) {
    Write-Status -Message 'Neither Get-RecipientPermission nor Get-ADPermission is available. SendAs export will be blank.' -Level WARN
}

$scopeMode = 'Csv'
$resolvedServer = Get-TrimmedValue -Value $Server
$resolvedSearchBase = ''
$runWasTruncated = $false

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    $resolvedSearchBase = Get-TrimmedValue -Value $SearchBase
    Write-Status -Message "DiscoverAll enabled for Exchange on-prem mailbox delegations. SearchBase='$resolvedSearchBase'." -Level WARN
    $rows = @([PSCustomObject]@{
            MailboxIdentity = '*'
        })
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
            TrusteeRecipientType      = ''
        }
    }

    if ($recipientSummaryByKey.ContainsKey($normalized)) {
        return $recipientSummaryByKey[$normalized]
    }

    $summary = $null
    try {
        $recipient = Invoke-WithRetry -OperationName "Lookup recipient $IdentityHint" -ScriptBlock {
            Get-Recipient -Identity $IdentityHint -ErrorAction Stop
        }

        $summary = [PSCustomObject]@{
            TrusteeIdentity           = Get-TrimmedValue -Value $recipient.Identity
            TrusteePrimarySmtpAddress = Get-TrimmedValue -Value $recipient.PrimarySmtpAddress
            TrusteeRecipientType      = Get-TrimmedValue -Value $recipient.RecipientType
        }
    }
    catch {
        $summary = [PSCustomObject]@{
            TrusteeIdentity           = $IdentityHint
            TrusteePrimarySmtpAddress = ''
            TrusteeRecipientType      = ''
        }
    }

    $recipientSummaryByKey[$normalized] = $summary
    return $summary
}

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = Get-TrimmedValue -Value $row.MailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
            throw 'MailboxIdentity is required. Use * to inventory delegations for all user/shared mailboxes.'
        }

        $effectiveSearchBase = if ($scopeMode -eq 'DiscoverAll') { $resolvedSearchBase } else { '' }
        $mailboxes = @(Invoke-WithRetry -OperationName "Load mailboxes for $mailboxIdentity" -ScriptBlock {
            Resolve-MailboxesByScope -Identity $mailboxIdentity -SearchBase $effectiveSearchBase -Server $resolvedServer
        })

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $mailboxes.Count -gt $MaxObjects) {
            $mailboxes = @($mailboxes | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($mailboxes.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'GetExchangeMailboxDelegation' -Status 'NotFound' -Message 'No matching mailboxes were found.' -Data ([ordered]@{
                        MailboxIdentity             = $mailboxIdentity
                        MailboxRecipientTypeDetails = ''
                        TrusteeIdentity             = ''
                        TrusteePrimarySmtpAddress   = ''
                        TrusteeRecipientType        = ''
                        FullAccess                  = ''
                        ReadOnly                    = ''
                        SendAs                      = ''
                        SendOnBehalf                = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $mailboxIdentityResolved = Get-TrimmedValue -Value $mailbox.Identity
            $permissionMap = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

            $ensureEntry = {
                param(
                    [Parameter(Mandatory)]
                    [string]$TrusteeHint
                )

                $summary = & $resolveRecipientSummary -IdentityHint $TrusteeHint
                $key = Normalize-TrusteeKey -Value $summary.TrusteeIdentity
                if ([string]::IsNullOrWhiteSpace($key)) {
                    $key = Normalize-TrusteeKey -Value $TrusteeHint
                }
                if ([string]::IsNullOrWhiteSpace($key)) {
                    return $null
                }

                if ($permissionMap.ContainsKey($key)) {
                    return $permissionMap[$key]
                }

                $entry = [PSCustomObject]@{
                    TrusteeIdentity           = $summary.TrusteeIdentity
                    TrusteePrimarySmtpAddress = $summary.TrusteePrimarySmtpAddress
                    TrusteeRecipientType      = $summary.TrusteeRecipientType
                    FullAccess                = $false
                    ReadOnly                  = $false
                    SendAs                    = $false
                    SendOnBehalf              = $false
                }

                $permissionMap[$key] = $entry
                return $entry
            }

            $mailboxPermissions = @(Invoke-WithRetry -OperationName "Load mailbox permissions $mailboxIdentityResolved" -ScriptBlock {
                Get-MailboxPermission -Identity $mailbox.Identity -ErrorAction Stop
            })

            foreach ($permission in $mailboxPermissions) {
                if ($permission.Deny -eq $true) { continue }
                if ($permission.IsInherited -eq $true) { continue }

                $trustee = Get-TrimmedValue -Value $permission.User
                if ([string]::IsNullOrWhiteSpace($trustee)) { continue }
                if ($trustee.Equals('NT AUTHORITY\\SELF', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                if ($trustee -match '^S-1-5-') { continue }

                $entry = & $ensureEntry -TrusteeHint $trustee
                if ($null -eq $entry) { continue }

                $accessRights = @($permission.AccessRights | ForEach-Object { Get-TrimmedValue -Value $_ })
                if ($accessRights -contains 'FullAccess') {
                    $entry.FullAccess = $true
                }
                if ($accessRights -contains 'ReadPermission') {
                    $entry.ReadOnly = $true
                }
            }

            if ($hasRecipientPermissionCmdlet) {
                $recipientPermissions = @(Invoke-WithRetry -OperationName "Load recipient permissions $mailboxIdentityResolved" -ScriptBlock {
                    Get-RecipientPermission -Identity $mailbox.Identity -ErrorAction SilentlyContinue
                })

                foreach ($permission in $recipientPermissions) {
                    if ($permission.Deny -eq $true) { continue }

                    $accessRights = @($permission.AccessRights | ForEach-Object { Get-TrimmedValue -Value $_ })
                    if ($accessRights -notcontains 'SendAs') { continue }

                    $trustee = Get-TrimmedValue -Value $permission.Trustee
                    if ([string]::IsNullOrWhiteSpace($trustee)) { continue }

                    $entry = & $ensureEntry -TrusteeHint $trustee
                    if ($null -eq $entry) { continue }

                    $entry.SendAs = $true
                }
            }
            elseif ($hasAdPermissionCmdlet) {
                $adPermissions = @(Invoke-WithRetry -OperationName "Load AD permissions $mailboxIdentityResolved" -ScriptBlock {
                    Get-ADPermission -Identity $mailbox.Identity -ErrorAction SilentlyContinue
                })

                foreach ($permission in $adPermissions) {
                    if ($permission.Deny -eq $true) { continue }
                    if ($permission.IsInherited -eq $true) { continue }

                    $extendedRights = @($permission.ExtendedRights | ForEach-Object { Get-TrimmedValue -Value $_ })
                    if ($extendedRights -notcontains 'Send As') { continue }

                    $trustee = Get-TrimmedValue -Value $permission.User
                    if ([string]::IsNullOrWhiteSpace($trustee)) { continue }
                    if ($trustee -match '^S-1-5-') { continue }

                    $entry = & $ensureEntry -TrusteeHint $trustee
                    if ($null -eq $entry) { continue }

                    $entry.SendAs = $true
                }
            }

            foreach ($delegate in @($mailbox.GrantSendOnBehalfTo)) {
                $delegateHint = Get-DelegateHint -DelegateObject $delegate
                if ([string]::IsNullOrWhiteSpace($delegateHint)) { continue }

                $entry = & $ensureEntry -TrusteeHint $delegateHint
                if ($null -eq $entry) { continue }

                $entry.SendOnBehalf = $true
            }

            if ($permissionMap.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityResolved -Action 'GetExchangeMailboxDelegation' -Status 'Completed' -Message 'No explicit delegated permissions found for mailbox.' -Data ([ordered]@{
                            MailboxIdentity             = $mailboxIdentityResolved
                            MailboxRecipientTypeDetails = Get-TrimmedValue -Value $mailbox.RecipientTypeDetails
                            TrusteeIdentity             = ''
                            TrusteePrimarySmtpAddress   = ''
                            TrusteeRecipientType        = ''
                            FullAccess                  = ''
                            ReadOnly                    = ''
                            SendAs                      = ''
                            SendOnBehalf                = ''
                        })))
                continue
            }

            foreach ($entry in @($permissionMap.Values | Sort-Object -Property TrusteeIdentity)) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$mailboxIdentityResolved|$($entry.TrusteeIdentity)" -Action 'GetExchangeMailboxDelegation' -Status 'Completed' -Message 'Mailbox delegation row exported.' -Data ([ordered]@{
                            MailboxIdentity             = $mailboxIdentityResolved
                            MailboxRecipientTypeDetails = Get-TrimmedValue -Value $mailbox.RecipientTypeDetails
                            TrusteeIdentity             = $entry.TrusteeIdentity
                            TrusteePrimarySmtpAddress   = $entry.TrusteePrimarySmtpAddress
                            TrusteeRecipientType        = $entry.TrusteeRecipientType
                            FullAccess                  = [string]$entry.FullAccess
                            ReadOnly                    = [string]$entry.ReadOnly
                            SendAs                      = [string]$entry.SendAs
                            SendOnBehalf                = [string]$entry.SendOnBehalf
                        })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'GetExchangeMailboxDelegation' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    MailboxIdentity             = $mailboxIdentity
                    MailboxRecipientTypeDetails = ''
                    TrusteeIdentity             = ''
                    TrusteePrimarySmtpAddress   = ''
                    TrusteeRecipientType        = ''
                    FullAccess                  = ''
                    ReadOnly                    = ''
                    SendAs                      = ''
                    SendOnBehalf                = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeSearchBase' -NotePropertyValue $resolvedSearchBase -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeServer' -NotePropertyValue $resolvedServer -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeMaxObjects' -NotePropertyValue ($(if ($scopeMode -eq 'DiscoverAll') { [string]$MaxObjects } else { '' })) -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeWasTruncated' -NotePropertyValue ([string]$runWasTruncated) -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem mailbox delegation inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
