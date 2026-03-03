<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-220000

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M0215-Set-ExchangeOnPremDistributionListMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

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

function Get-RecipientKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Recipient
    )

    $primary = Get-TrimmedValue -Value $Recipient.PrimarySmtpAddress
    if (-not [string]::IsNullOrWhiteSpace($primary)) {
        return $primary.ToLowerInvariant()
    }

    $identity = Get-TrimmedValue -Value $Recipient.Identity
    if (-not [string]::IsNullOrWhiteSpace($identity)) {
        return $identity.ToLowerInvariant()
    }

    return ''
}

$requiredHeaders = @(
    'DistributionGroupIdentity',
    'MemberIdentity',
    'MemberAction'
)

Write-Status -Message 'Starting Exchange on-prem distribution list membership script.'
Ensure-ExchangeOnPremConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $distributionGroupIdentity = Get-TrimmedValue -Value $row.DistributionGroupIdentity
    $memberIdentity = Get-TrimmedValue -Value $row.MemberIdentity
    $memberActionRaw = Get-TrimmedValue -Value $row.MemberAction

    try {
        if ([string]::IsNullOrWhiteSpace($distributionGroupIdentity) -or [string]::IsNullOrWhiteSpace($memberIdentity)) {
            throw 'DistributionGroupIdentity and MemberIdentity are required.'
        }

        $memberAction = if ([string]::IsNullOrWhiteSpace($memberActionRaw)) { 'Add' } else { $memberActionRaw }
        if ($memberAction -notin @('Add', 'Remove')) {
            throw "MemberAction '$memberAction' is invalid. Use Add or Remove."
        }

        $group = Invoke-WithRetry -OperationName "Lookup distribution list $distributionGroupIdentity" -ScriptBlock {
            Get-DistributionGroup -Identity $distributionGroupIdentity -ErrorAction SilentlyContinue
        }

        if (-not $group) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberIdentity" -Action 'SetDistributionListMember' -Status 'NotFound' -Message 'Distribution list not found.'))
            $rowNumber++
            continue
        }

        if (-not (Test-IsDistributionListRecipientType -RecipientTypeDetails $group.RecipientTypeDetails)) {
            throw "Recipient '$distributionGroupIdentity' is '$($group.RecipientTypeDetails)'. Expected a distribution list recipient type."
        }

        $memberRecipient = Invoke-WithRetry -OperationName "Lookup member recipient $memberIdentity" -ScriptBlock {
            Get-Recipient -Identity $memberIdentity -ErrorAction SilentlyContinue
        }
        if (-not $memberRecipient) {
            throw "Member recipient '$memberIdentity' was not found."
        }

        $memberKey = Get-RecipientKey -Recipient $memberRecipient
        if ([string]::IsNullOrWhiteSpace($memberKey)) {
            throw "Unable to resolve a stable member key for '$memberIdentity'."
        }

        $members = @(Invoke-WithRetry -OperationName "Load distribution list members $distributionGroupIdentity" -ScriptBlock {
            Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop
        })

        $memberSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($existingMember in $members) {
            $key = Get-RecipientKey -Recipient $existingMember
            if (-not [string]::IsNullOrWhiteSpace($key)) {
                $null = $memberSet.Add($key)
            }
        }

        if ($memberAction -eq 'Add') {
            if ($memberSet.Contains($memberKey)) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberIdentity|Add" -Action 'SetDistributionListMember' -Status 'Skipped' -Message 'Member is already in the distribution list.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$distributionGroupIdentity -> $memberIdentity", 'Add distribution list member')) {
                Invoke-WithRetry -OperationName "Add distribution list member $memberIdentity to $distributionGroupIdentity" -ScriptBlock {
                    Add-DistributionGroupMember -Identity $group.Identity -Member $memberRecipient.Identity -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberIdentity|Add" -Action 'SetDistributionListMember' -Status 'Added' -Message 'Member added to distribution list.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberIdentity|Add" -Action 'SetDistributionListMember' -Status 'WhatIf' -Message 'Add operation skipped due to WhatIf.'))
            }
        }
        else {
            if (-not $memberSet.Contains($memberKey)) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberIdentity|Remove" -Action 'SetDistributionListMember' -Status 'Skipped' -Message 'Member is not currently in the distribution list.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$distributionGroupIdentity -> $memberIdentity", 'Remove distribution list member')) {
                Invoke-WithRetry -OperationName "Remove distribution list member $memberIdentity from $distributionGroupIdentity" -ScriptBlock {
                    Remove-DistributionGroupMember -Identity $group.Identity -Member $memberRecipient.Identity -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberIdentity|Remove" -Action 'SetDistributionListMember' -Status 'Removed' -Message 'Member removed from distribution list.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberIdentity|Remove" -Action 'SetDistributionListMember' -Status 'WhatIf' -Message 'Remove operation skipped due to WhatIf.'))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($distributionGroupIdentity|$memberIdentity|$memberActionRaw) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$distributionGroupIdentity|$memberIdentity|$memberActionRaw" -Action 'SetDistributionListMember' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem distribution list membership script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
