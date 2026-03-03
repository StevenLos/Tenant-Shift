<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-000100

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR0214-Get-ExchangeOnPremDistributionLists_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Test-IsDistributionListRecipientType {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$RecipientTypeDetails
    )

    $recipientTypeText = Get-TrimmedValue -Value $RecipientTypeDetails
    return ($recipientTypeText -eq 'MailUniversalDistributionGroup' -or $recipientTypeText -eq 'MailNonUniversalGroup')
}

function Resolve-DistributionGroupsByScope {
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
            ResultSize  = 'Unlimited'
            ErrorAction = 'Stop'
        }

        Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-DistributionGroup' -ParameterName 'OrganizationalUnit' -Value $SearchBase
        Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-DistributionGroup' -ParameterName 'DomainController' -Value $Server

        return @(
            Get-DistributionGroup @params |
                Where-Object { Test-IsDistributionListRecipientType -RecipientTypeDetails $_.RecipientTypeDetails }
        )
    }

    $params = @{
        Identity    = $Identity
        ErrorAction = 'SilentlyContinue'
    }

    Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-DistributionGroup' -ParameterName 'DomainController' -Value $Server

    $group = Get-DistributionGroup @params
    if ($group -and (Test-IsDistributionListRecipientType -RecipientTypeDetails $group.RecipientTypeDetails)) {
        return @($group)
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
    'DistributionGroupIdentity'
)

Write-Status -Message 'Starting Exchange on-prem distribution list inventory script.'
Ensure-ExchangeOnPremConnection

$scopeMode = 'Csv'
$resolvedServer = Get-TrimmedValue -Value $Server
$resolvedSearchBase = ''
$runWasTruncated = $false

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    $resolvedSearchBase = Get-TrimmedValue -Value $SearchBase
    Write-Status -Message "DiscoverAll enabled for Exchange on-prem distribution lists. SearchBase='$resolvedSearchBase'." -Level WARN
    $rows = @([PSCustomObject]@{
            DistributionGroupIdentity = '*'
        })
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $distributionGroupIdentity = Get-TrimmedValue -Value $row.DistributionGroupIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($distributionGroupIdentity)) {
            throw 'DistributionGroupIdentity is required. Use * to inventory all distribution lists.'
        }

        $effectiveSearchBase = if ($scopeMode -eq 'DiscoverAll') { $resolvedSearchBase } else { '' }
        $groups = @(Invoke-WithRetry -OperationName "Load distribution lists for $distributionGroupIdentity" -ScriptBlock {
            Resolve-DistributionGroupsByScope -Identity $distributionGroupIdentity -SearchBase $effectiveSearchBase -Server $resolvedServer
        })

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $groups.Count -gt $MaxObjects) {
            $groups = @($groups | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'GetExchangeDistributionList' -Status 'NotFound' -Message 'No matching distribution lists were found.' -Data ([ordered]@{
                        DistributionGroupIdentity            = $distributionGroupIdentity
                        Name                                 = ''
                        Alias                                = ''
                        DisplayName                          = ''
                        PrimarySmtpAddress                   = ''
                        ManagedBy                            = ''
                        Notes                                = ''
                        MemberJoinRestriction                = ''
                        MemberDepartRestriction              = ''
                        ModerationEnabled                    = ''
                        ModeratedBy                          = ''
                        RequireSenderAuthenticationEnabled   = ''
                        HiddenFromAddressListsEnabled        = ''
                        AcceptMessagesOnlyFrom               = ''
                        AcceptMessagesOnlyFromDLMembers      = ''
                        RejectMessagesFrom                   = ''
                        RejectMessagesFromDLMembers          = ''
                        BypassModerationFromSendersOrMembers = ''
                        SendModerationNotifications          = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Identity)) {
            $identity = Get-TrimmedValue -Value $group.Identity
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeDistributionList' -Status 'Completed' -Message 'Distribution list exported.' -Data ([ordered]@{
                        DistributionGroupIdentity            = $identity
                        Name                                 = Get-TrimmedValue -Value $group.Name
                        Alias                                = Get-TrimmedValue -Value $group.Alias
                        DisplayName                          = Get-TrimmedValue -Value $group.DisplayName
                        PrimarySmtpAddress                   = Get-TrimmedValue -Value $group.PrimarySmtpAddress
                        ManagedBy                            = Convert-MultiValueToString -Value $group.ManagedBy
                        Notes                                = Get-TrimmedValue -Value $group.Notes
                        MemberJoinRestriction                = Get-TrimmedValue -Value $group.MemberJoinRestriction
                        MemberDepartRestriction              = Get-TrimmedValue -Value $group.MemberDepartRestriction
                        ModerationEnabled                    = [string]$group.ModerationEnabled
                        ModeratedBy                          = Convert-MultiValueToString -Value $group.ModeratedBy
                        RequireSenderAuthenticationEnabled   = [string]$group.RequireSenderAuthenticationEnabled
                        HiddenFromAddressListsEnabled        = [string]$group.HiddenFromAddressListsEnabled
                        AcceptMessagesOnlyFrom               = Convert-MultiValueToString -Value $group.AcceptMessagesOnlyFrom
                        AcceptMessagesOnlyFromDLMembers      = Convert-MultiValueToString -Value $group.AcceptMessagesOnlyFromDLMembers
                        RejectMessagesFrom                   = Convert-MultiValueToString -Value $group.RejectMessagesFrom
                        RejectMessagesFromDLMembers          = Convert-MultiValueToString -Value $group.RejectMessagesFromDLMembers
                        BypassModerationFromSendersOrMembers = Convert-MultiValueToString -Value $group.BypassModerationFromSendersOrMembers
                        SendModerationNotifications          = Get-TrimmedValue -Value $group.SendModerationNotifications
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($distributionGroupIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'GetExchangeDistributionList' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    DistributionGroupIdentity            = $distributionGroupIdentity
                    Name                                 = ''
                    Alias                                = ''
                    DisplayName                          = ''
                    PrimarySmtpAddress                   = ''
                    ManagedBy                            = ''
                    Notes                                = ''
                    MemberJoinRestriction                = ''
                    MemberDepartRestriction              = ''
                    ModerationEnabled                    = ''
                    ModeratedBy                          = ''
                    RequireSenderAuthenticationEnabled   = ''
                    HiddenFromAddressListsEnabled        = ''
                    AcceptMessagesOnlyFrom               = ''
                    AcceptMessagesOnlyFromDLMembers      = ''
                    RejectMessagesFrom                   = ''
                    RejectMessagesFromDLMembers          = ''
                    BypassModerationFromSendersOrMembers = ''
                    SendModerationNotifications          = ''
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
Write-Status -Message 'Exchange on-prem distribution list inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
