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

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3123-Get-ExchangeOnlineDynamicDistributionGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'DynamicDistributionGroupIdentity'
)

Write-Status -Message 'Starting Exchange Online dynamic distribution group inventory script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $dynamicDistributionGroupIdentity = ([string]$row.DynamicDistributionGroupIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($dynamicDistributionGroupIdentity)) {
            throw 'DynamicDistributionGroupIdentity is required. Use * to inventory all dynamic distribution groups.'
        }

        $groups = @()
        if ($dynamicDistributionGroupIdentity -eq '*') {
            $groups = @(Invoke-WithRetry -OperationName 'Load all dynamic distribution groups' -ScriptBlock {
                Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction Stop
            })
        }
        else {
            $group = Invoke-WithRetry -OperationName "Lookup dynamic distribution group $dynamicDistributionGroupIdentity" -ScriptBlock {
                Get-DynamicDistributionGroup -Identity $dynamicDistributionGroupIdentity -ErrorAction SilentlyContinue
            }
            if ($group) {
                $groups = @($group)
            }
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $dynamicDistributionGroupIdentity -Action 'GetExchangeDynamicDistributionGroup' -Status 'NotFound' -Message 'No matching dynamic distribution groups were found.' -Data ([ordered]@{
                        DynamicDistributionGroupIdentity        = $dynamicDistributionGroupIdentity
                        Name                                    = ''
                        Alias                                   = ''
                        DisplayName                             = ''
                        PrimarySmtpAddress                      = ''
                        ManagedBy                               = ''
                        RecipientFilter                         = ''
                        IncludedRecipients                      = ''
                        ConditionalCompany                      = ''
                        ConditionalDepartment                   = ''
                        ConditionalCustomAttribute1             = ''
                        ConditionalCustomAttribute2             = ''
                        ConditionalStateOrProvince              = ''
                        RequireSenderAuthenticationEnabled      = ''
                        HiddenFromAddressListsEnabled           = ''
                        ModerationEnabled                       = ''
                        ModeratedBy                             = ''
                        SendModerationNotifications             = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Identity)) {
            $identity = ([string]$group.Identity).Trim()
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeDynamicDistributionGroup' -Status 'Completed' -Message 'Dynamic distribution group exported.' -Data ([ordered]@{
                        DynamicDistributionGroupIdentity        = $identity
                        Name                                    = ([string]$group.Name).Trim()
                        Alias                                   = ([string]$group.Alias).Trim()
                        DisplayName                             = ([string]$group.DisplayName).Trim()
                        PrimarySmtpAddress                      = ([string]$group.PrimarySmtpAddress).Trim()
                        ManagedBy                               = Convert-MultiValueToString -Value $group.ManagedBy
                        RecipientFilter                         = ([string]$group.RecipientFilter).Trim()
                        IncludedRecipients                      = ([string]$group.IncludedRecipients).Trim()
                        ConditionalCompany                      = Convert-MultiValueToString -Value $group.ConditionalCompany
                        ConditionalDepartment                   = Convert-MultiValueToString -Value $group.ConditionalDepartment
                        ConditionalCustomAttribute1             = Convert-MultiValueToString -Value $group.ConditionalCustomAttribute1
                        ConditionalCustomAttribute2             = Convert-MultiValueToString -Value $group.ConditionalCustomAttribute2
                        ConditionalStateOrProvince              = Convert-MultiValueToString -Value $group.ConditionalStateOrProvince
                        RequireSenderAuthenticationEnabled      = [string]$group.RequireSenderAuthenticationEnabled
                        HiddenFromAddressListsEnabled           = [string]$group.HiddenFromAddressListsEnabled
                        ModerationEnabled                       = [string]$group.ModerationEnabled
                        ModeratedBy                             = Convert-MultiValueToString -Value $group.ModeratedBy
                        SendModerationNotifications             = ([string]$group.SendModerationNotifications).Trim()
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($dynamicDistributionGroupIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $dynamicDistributionGroupIdentity -Action 'GetExchangeDynamicDistributionGroup' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    DynamicDistributionGroupIdentity        = $dynamicDistributionGroupIdentity
                    Name                                    = ''
                    Alias                                   = ''
                    DisplayName                             = ''
                    PrimarySmtpAddress                      = ''
                    ManagedBy                               = ''
                    RecipientFilter                         = ''
                    IncludedRecipients                      = ''
                    ConditionalCompany                      = ''
                    ConditionalDepartment                   = ''
                    ConditionalCustomAttribute1             = ''
                    ConditionalCustomAttribute2             = ''
                    ConditionalStateOrProvince              = ''
                    RequireSenderAuthenticationEnabled      = ''
                    HiddenFromAddressListsEnabled           = ''
                    ModerationEnabled                       = ''
                    ModeratedBy                             = ''
                    SendModerationNotifications             = ''
                })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online dynamic distribution group inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





