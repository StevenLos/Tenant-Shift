<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-000101

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3115-Set-ExchangeOnlineDistributionListMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {


$requiredHeaders = @(
    'DistributionGroupIdentity',
    'MemberUserPrincipalName',
    'BypassSecurityGroupManagerCheck',
    'MemberAction'
)

Write-Status -Message 'Starting Exchange Online distribution list membership script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $distributionGroupIdentity = ([string]$row.DistributionGroupIdentity).Trim()
    $memberUpn = ([string]$row.MemberUserPrincipalName).Trim()
    $actionLabel = 'AddDistributionGroupMember'

    try {
        if ([string]::IsNullOrWhiteSpace($distributionGroupIdentity) -or [string]::IsNullOrWhiteSpace($memberUpn)) {
            throw 'DistributionGroupIdentity and MemberUserPrincipalName are required.'
        }

        $memberActionRaw = 'Add'
        if ($row.PSObject.Properties.Name -contains 'MemberAction') {
            $memberActionCandidate = ([string]$row.MemberAction).Trim()
            if (-not [string]::IsNullOrWhiteSpace($memberActionCandidate)) {
                $memberActionRaw = $memberActionCandidate
            }
        }

        $memberAction = $memberActionRaw.ToLowerInvariant()
        if ($memberAction -notin @('add', 'remove')) {
            throw "MemberAction '$memberActionRaw' is invalid. Use Add or Remove."
        }

        $actionLabel = if ($memberAction -eq 'remove') { 'RemoveDistributionGroupMember' } else { 'AddDistributionGroupMember' }

        $group = Invoke-WithRetry -OperationName "Lookup distribution list $distributionGroupIdentity" -ScriptBlock {
            Get-DistributionGroup -Identity $distributionGroupIdentity -ErrorAction Stop
        }
        $memberRecipient = Invoke-WithRetry -OperationName "Lookup recipient $memberUpn" -ScriptBlock {
            Get-ExchangeOnlineRecipient -Identity $memberUpn -ErrorAction Stop
        }

        $memberDistinguishedName = ([string]$memberRecipient.DistinguishedName).Trim()
        if ([string]::IsNullOrWhiteSpace($memberDistinguishedName)) {
            throw "Recipient '$memberUpn' does not have a DistinguishedName. Cannot safely determine membership."
        }

        $existingMembership = Invoke-WithRetry -OperationName "Check membership for $distributionGroupIdentity -> $memberUpn" -ScriptBlock {
            Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop |
                Where-Object { $_.DistinguishedName -eq $memberDistinguishedName } |
                Select-Object -First 1
        }

        $bypass = ConvertTo-Bool -Value $row.BypassSecurityGroupManagerCheck -Default $false

        if ($memberAction -eq 'add') {
            if ($existingMembership) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberUpn" -Action $actionLabel -Status 'Skipped' -Message 'Member already exists in the distribution list.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$distributionGroupIdentity -> $memberUpn", 'Add member to Exchange Online distribution list')) {
                $params = @{
                    Identity    = $group.Identity
                    Member      = $memberDistinguishedName
                    ErrorAction = 'Stop'
                }

                if ($bypass) {
                    $params.BypassSecurityGroupManagerCheck = $true
                }

                Invoke-WithRetry -OperationName "Add member $memberUpn to $distributionGroupIdentity" -ScriptBlock {
                    Add-DistributionGroupMember @params
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberUpn" -Action $actionLabel -Status 'Added' -Message 'Member added successfully.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberUpn" -Action $actionLabel -Status 'WhatIf' -Message 'Membership update skipped due to WhatIf.'))
            }
        }
        else {
            if (-not $existingMembership) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberUpn" -Action $actionLabel -Status 'Skipped' -Message 'Member is not currently in the distribution list.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$distributionGroupIdentity -> $memberUpn", 'Remove member from Exchange Online distribution list')) {
                $params = @{
                    Identity    = $group.Identity
                    Member      = $memberDistinguishedName
                    Confirm     = $false
                    ErrorAction = 'Stop'
                }

                if ($bypass) {
                    $params.BypassSecurityGroupManagerCheck = $true
                }

                Invoke-WithRetry -OperationName "Remove member $memberUpn from $distributionGroupIdentity" -ScriptBlock {
                    Remove-DistributionGroupMember @params
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberUpn" -Action $actionLabel -Status 'Removed' -Message 'Member removed successfully.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberUpn" -Action $actionLabel -Status 'WhatIf' -Message 'Membership update skipped due to WhatIf.'))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($distributionGroupIdentity|$memberUpn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberUpn" -Action $actionLabel -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online distribution list membership script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





