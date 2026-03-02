<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-004416

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3115-Get-ExchangeOnlineDistributionListMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$requiredHeaders = @(
    'DistributionGroupIdentity'
)

Write-Status -Message 'Starting Exchange Online distribution list member inventory script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $distributionGroupIdentity = ([string]$row.DistributionGroupIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($distributionGroupIdentity)) {
            throw 'DistributionGroupIdentity is required. Use * to inventory members for all distribution lists.'
        }

        $groups = @()
        if ($distributionGroupIdentity -eq '*') {
            $groups = @(Invoke-WithRetry -OperationName 'Load all distribution lists for membership export' -ScriptBlock {
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
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'GetExchangeDistributionListMember' -Status 'NotFound' -Message 'No matching distribution lists were found.' -Data ([ordered]@{
                        DistributionGroupIdentity         = $distributionGroupIdentity
                        DistributionGroupDisplayName      = ''
                        MemberIdentity                    = ''
                        MemberDisplayName                 = ''
                        MemberPrimarySmtpAddress          = ''
                        MemberRecipientType               = ''
                        MemberExternalEmailAddress        = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Identity)) {
            $groupIdentity = ([string]$group.Identity).Trim()
            $groupDisplayName = ([string]$group.DisplayName).Trim()

            $members = @(Invoke-WithRetry -OperationName "Load members for distribution list $groupIdentity" -ScriptBlock {
                Get-DistributionGroupMember -Identity $groupIdentity -ResultSize Unlimited -ErrorAction Stop
            })

            if ($members.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupIdentity -Action 'GetExchangeDistributionListMember' -Status 'Completed' -Message 'Distribution list has no members.' -Data ([ordered]@{
                            DistributionGroupIdentity         = $groupIdentity
                            DistributionGroupDisplayName      = $groupDisplayName
                            MemberIdentity                    = ''
                            MemberDisplayName                 = ''
                            MemberPrimarySmtpAddress          = ''
                            MemberRecipientType               = ''
                            MemberExternalEmailAddress        = ''
                        })))
                continue
            }

            foreach ($member in @($members | Sort-Object -Property DisplayName, Identity)) {
                $memberIdentity = ([string]$member.Identity).Trim()
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$groupIdentity|$memberIdentity" -Action 'GetExchangeDistributionListMember' -Status 'Completed' -Message 'Distribution list member exported.' -Data ([ordered]@{
                            DistributionGroupIdentity         = $groupIdentity
                            DistributionGroupDisplayName      = $groupDisplayName
                            MemberIdentity                    = $memberIdentity
                            MemberDisplayName                 = ([string]$member.DisplayName).Trim()
                            MemberPrimarySmtpAddress          = ([string]$member.PrimarySmtpAddress).Trim()
                            MemberRecipientType               = ([string]$member.RecipientType).Trim()
                            MemberExternalEmailAddress        = ([string]$member.ExternalEmailAddress).Trim()
                        })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($distributionGroupIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $distributionGroupIdentity -Action 'GetExchangeDistributionListMember' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    DistributionGroupIdentity         = $distributionGroupIdentity
                    DistributionGroupDisplayName      = ''
                    MemberIdentity                    = ''
                    MemberDisplayName                 = ''
                    MemberPrimarySmtpAddress          = ''
                    MemberRecipientType               = ''
                    MemberExternalEmailAddress        = ''
                })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online distribution list member inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}








