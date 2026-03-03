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
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3114-Get-ExchangeOnlineDistributionLists_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

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

Write-Status -Message 'Starting Exchange Online distribution list inventory script.'
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
    $distributionGroupIdentity = ([string]$row.DistributionGroupIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($distributionGroupIdentity)) {
            throw 'DistributionGroupIdentity is required. Use * to inventory all distribution lists.'
        }

        $groups = @()
        if ($distributionGroupIdentity -eq '*') {
            $groups = @(Invoke-WithRetry -OperationName 'Load all distribution lists' -ScriptBlock {
                Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop
            })
        }
        else {
            $group = Invoke-WithRetry -OperationName "Lookup distribution list $distributionGroupIdentity" -ScriptBlock {
                Get-DistributionGroup -Identity $distributionGroupIdentity -ErrorAction SilentlyContinue
            }

            if ($group) {
                $groups = @($group)
            }
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'GetExchangeDistributionList' -Status 'NotFound' -Message 'No matching distribution lists were found.' -Data ([ordered]@{
                        DistributionGroupIdentity                 = $distributionGroupIdentity
                        Name                                      = ''
                        Alias                                     = ''
                        DisplayName                               = ''
                        PrimarySmtpAddress                        = ''
                        ManagedBy                                 = ''
                        Notes                                     = ''
                        MemberJoinRestriction                     = ''
                        MemberDepartRestriction                   = ''
                        ModerationEnabled                         = ''
                        ModeratedBy                               = ''
                        RequireSenderAuthenticationEnabled        = ''
                        HiddenFromAddressListsEnabled             = ''
                        AcceptMessagesOnlyFrom                    = ''
                        AcceptMessagesOnlyFromDLMembers           = ''
                        RejectMessagesFrom                        = ''
                        RejectMessagesFromDLMembers               = ''
                        BypassModerationFromSendersOrMembers      = ''
                        SendModerationNotifications               = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Identity)) {
            $identity = ([string]$group.Identity).Trim()
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeDistributionList' -Status 'Completed' -Message 'Distribution list exported.' -Data ([ordered]@{
                        DistributionGroupIdentity                 = $identity
                        Name                                      = ([string]$group.Name).Trim()
                        Alias                                     = ([string]$group.Alias).Trim()
                        DisplayName                               = ([string]$group.DisplayName).Trim()
                        PrimarySmtpAddress                        = ([string]$group.PrimarySmtpAddress).Trim()
                        ManagedBy                                 = Convert-MultiValueToString -Value $group.ManagedBy
                        Notes                                     = ([string]$group.Notes).Trim()
                        MemberJoinRestriction                     = ([string]$group.MemberJoinRestriction).Trim()
                        MemberDepartRestriction                   = ([string]$group.MemberDepartRestriction).Trim()
                        ModerationEnabled                         = [string]$group.ModerationEnabled
                        ModeratedBy                               = Convert-MultiValueToString -Value $group.ModeratedBy
                        RequireSenderAuthenticationEnabled        = [string]$group.RequireSenderAuthenticationEnabled
                        HiddenFromAddressListsEnabled             = [string]$group.HiddenFromAddressListsEnabled
                        AcceptMessagesOnlyFrom                    = Convert-MultiValueToString -Value $group.AcceptMessagesOnlyFrom
                        AcceptMessagesOnlyFromDLMembers           = Convert-MultiValueToString -Value $group.AcceptMessagesOnlyFromDLMembers
                        RejectMessagesFrom                        = Convert-MultiValueToString -Value $group.RejectMessagesFrom
                        RejectMessagesFromDLMembers               = Convert-MultiValueToString -Value $group.RejectMessagesFromDLMembers
                        BypassModerationFromSendersOrMembers      = Convert-MultiValueToString -Value $group.BypassModerationFromSendersOrMembers
                        SendModerationNotifications               = ([string]$group.SendModerationNotifications).Trim()
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($distributionGroupIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'GetExchangeDistributionList' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    DistributionGroupIdentity                 = $distributionGroupIdentity
                    Name                                      = ''
                    Alias                                     = ''
                    DisplayName                               = ''
                    PrimarySmtpAddress                        = ''
                    ManagedBy                                 = ''
                    Notes                                     = ''
                    MemberJoinRestriction                     = ''
                    MemberDepartRestriction                   = ''
                    ModerationEnabled                         = ''
                    ModeratedBy                               = ''
                    RequireSenderAuthenticationEnabled        = ''
                    HiddenFromAddressListsEnabled             = ''
                    AcceptMessagesOnlyFrom                    = ''
                    AcceptMessagesOnlyFromDLMembers           = ''
                    RejectMessagesFrom                        = ''
                    RejectMessagesFromDLMembers               = ''
                    BypassModerationFromSendersOrMembers      = ''
                    SendModerationNotifications               = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online distribution list inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}










