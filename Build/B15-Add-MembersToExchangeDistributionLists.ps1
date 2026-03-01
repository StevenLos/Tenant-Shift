#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath ("Results_B15-Add-MembersToExchangeDistributionLists_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'M365.Common.psm1') -Force -DisableNameChecking

$requiredHeaders = @(
    'DistributionGroupIdentity',
    'MemberUserPrincipalName',
    'BypassSecurityGroupManagerCheck'
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

    try {
        if ([string]::IsNullOrWhiteSpace($distributionGroupIdentity) -or [string]::IsNullOrWhiteSpace($memberUpn)) {
            throw 'DistributionGroupIdentity and MemberUserPrincipalName are required.'
        }

        $group = Invoke-WithRetry -OperationName "Lookup distribution list $distributionGroupIdentity" -ScriptBlock {
            Get-DistributionGroup -Identity $distributionGroupIdentity -ErrorAction Stop
        }
        $memberRecipient = Invoke-WithRetry -OperationName "Lookup recipient $memberUpn" -ScriptBlock {
            Get-Recipient -Identity $memberUpn -ErrorAction Stop
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

        if ($existingMembership) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberUpn" -Action 'AddDistributionGroupMember' -Status 'Skipped' -Message 'Member already exists in the distribution list.'))
            $rowNumber++
            continue
        }

        $bypass = ConvertTo-Bool -Value $row.BypassSecurityGroupManagerCheck -Default $false

        if ($PSCmdlet.ShouldProcess("$distributionGroupIdentity -> $memberUpn", 'Add member to Exchange Online distribution list')) {
            $params = @{
                Identity   = $group.Identity
                Member     = $memberDistinguishedName
                ErrorAction = 'Stop'
            }

            if ($bypass) {
                $params.BypassSecurityGroupManagerCheck = $true
            }

            Invoke-WithRetry -OperationName "Add member $memberUpn to $distributionGroupIdentity" -ScriptBlock {
                Add-DistributionGroupMember @params
            }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberUpn" -Action 'AddDistributionGroupMember' -Status 'Added' -Message 'Member added successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberUpn" -Action 'AddDistributionGroupMember' -Status 'WhatIf' -Message 'Membership update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($distributionGroupIdentity|$memberUpn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberUpn" -Action 'AddDistributionGroupMember' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online distribution list membership script completed.' -Level SUCCESS

