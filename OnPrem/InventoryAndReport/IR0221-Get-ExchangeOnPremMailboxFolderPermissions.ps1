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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR0221-Get-ExchangeOnPremMailboxFolderPermissions_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Convert-MultiValueToString {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [string]) {
        return ([string]$Value).Trim()
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        $items = [System.Collections.Generic.List[string]]::new()
        foreach ($item in $Value) {
            $text = ([string]$item).Trim()
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $items.Add($text)
            }
        }

        return (@($items | Sort-Object -Unique) -join ';')
    }

    return ([string]$Value).Trim()
}

$requiredHeaders = @(
    'MailboxIdentity'
)

Write-Status -Message 'Starting Exchange on-prem mailbox folder permission inventory script.'
Ensure-ExchangeOnPremConnection

$scopeMode = 'Csv'
$resolvedServer = Get-TrimmedValue -Value $Server
$resolvedSearchBase = ''
$runWasTruncated = $false

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    $resolvedSearchBase = Get-TrimmedValue -Value $SearchBase
    Write-Status -Message "DiscoverAll enabled for Exchange on-prem mailbox folder permissions. SearchBase='$resolvedSearchBase'." -Level WARN
    $rows = @([PSCustomObject]@{
            MailboxIdentity = '*'
        })
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = Get-TrimmedValue -Value $row.MailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
            throw 'MailboxIdentity is required. Use * to inventory folder permissions for all user/shared mailboxes.'
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
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'GetExchangeMailboxFolderPermission' -Status 'NotFound' -Message 'No matching mailboxes were found.' -Data ([ordered]@{
                        MailboxIdentity             = $mailboxIdentity
                        MailboxRecipientTypeDetails = ''
                        FolderIdentity              = ''
                        FolderPath                  = ''
                        TrusteeIdentity             = ''
                        AccessRights                = ''
                        SharingPermissionFlags      = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $mailboxIdentityResolved = Get-TrimmedValue -Value $mailbox.Identity
            $folderStats = @(Invoke-WithRetry -OperationName "Load mailbox folders $mailboxIdentityResolved" -ScriptBlock {
                Get-MailboxFolderStatistics -Identity $mailboxIdentityResolved -ErrorAction Stop
            })

            $rowAdded = $false
            foreach ($folderStat in $folderStats) {
                $folderIdentity = Get-TrimmedValue -Value $folderStat.Identity
                if ([string]::IsNullOrWhiteSpace($folderIdentity)) {
                    continue
                }

                $permissions = @(Invoke-WithRetry -OperationName "Load folder permissions $folderIdentity" -ScriptBlock {
                    Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction SilentlyContinue
                })

                foreach ($permission in $permissions) {
                    $userText = Get-TrimmedValue -Value $permission.User
                    if ([string]::IsNullOrWhiteSpace($userText)) { continue }
                    if ($userText.Equals('Default', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                    if ($userText.Equals('Anonymous', [System.StringComparison]::OrdinalIgnoreCase)) { continue }

                    $rowAdded = $true
                    $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$mailboxIdentityResolved|$folderIdentity|$userText" -Action 'GetExchangeMailboxFolderPermission' -Status 'Completed' -Message 'Mailbox folder permission row exported.' -Data ([ordered]@{
                                MailboxIdentity             = $mailboxIdentityResolved
                                MailboxRecipientTypeDetails = Get-TrimmedValue -Value $mailbox.RecipientTypeDetails
                                FolderIdentity              = $folderIdentity
                                FolderPath                  = Get-TrimmedValue -Value $folderStat.FolderPath
                                TrusteeIdentity             = $userText
                                AccessRights                = Convert-MultiValueToString -Value $permission.AccessRights
                                SharingPermissionFlags      = Convert-MultiValueToString -Value $permission.SharingPermissionFlags
                            })))
                }
            }

            if (-not $rowAdded) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityResolved -Action 'GetExchangeMailboxFolderPermission' -Status 'Completed' -Message 'No explicit folder permissions found.' -Data ([ordered]@{
                            MailboxIdentity             = $mailboxIdentityResolved
                            MailboxRecipientTypeDetails = Get-TrimmedValue -Value $mailbox.RecipientTypeDetails
                            FolderIdentity              = ''
                            FolderPath                  = ''
                            TrusteeIdentity             = ''
                            AccessRights                = ''
                            SharingPermissionFlags      = ''
                        })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'GetExchangeMailboxFolderPermission' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    MailboxIdentity             = $mailboxIdentity
                    MailboxRecipientTypeDetails = ''
                    FolderIdentity              = ''
                    FolderPath                  = ''
                    TrusteeIdentity             = ''
                    AccessRights                = ''
                    SharingPermissionFlags      = ''
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
Write-Status -Message 'Exchange on-prem mailbox folder permission inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
