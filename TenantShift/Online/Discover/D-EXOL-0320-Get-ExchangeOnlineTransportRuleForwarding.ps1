<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260406-000000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Exports transport/mail flow rules that contain redirect, forward, or BCC actions.

.DESCRIPTION
    Exports all transport rules (mail flow rules) in the tenant that include redirect,
    forward, or blind-carbon-copy (BCC) actions. Transport rules are tenant-level constructs
    and do not require a mailbox scope — this script has no -InputCsvPath parameter and
    always exports all qualifying rules.
    Outputs rule name, priority, enabled state, a summary of rule conditions, and the
    action detail for each forwarding-type action present.
    All results are written to the output CSV.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-EXOL-0320-Get-ExchangeOnlineTransportRuleForwarding.ps1

    Export all transport rules with forwarding or BCC actions.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator
    Limitations:      Transport rules are tenant-scoped — no input CSV is required or accepted.
                      This script can run independently of D-EXOL-0300 and D-EXOL-0310.

    CSV Fields:
    (No input CSV — tenant-level only)
#>
#Requires -Version 7.0

[CmdletBinding()]
param(
    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-EXOL-0320-Get-ExchangeOnlineTransportRuleForwarding_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
        [Parameter(Mandatory)][int]$RowNumber,
        [Parameter(Mandatory)][string]$PrimaryKey,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$Message,
        [Parameter(Mandatory)][hashtable]$Data
    )

    $base    = New-ResultObject -RowNumber $RowNumber -PrimaryKey $PrimaryKey -Action $Action -Status $Status -Message $Message
    $ordered = [ordered]@{}

    foreach ($prop in $base.PSObject.Properties.Name) {
        $ordered[$prop] = $base.$prop
    }

    foreach ($key in $Data.Keys) {
        $ordered[$key] = $Data[$key]
    }

    return [PSCustomObject]$ordered
}

function Resolve-AddressList {
    [CmdletBinding()]
    param([object]$Value)

    if ($null -eq $Value) { return '' }
    $items = @([string]$Value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($items.Count -eq 0) { return '' }
    return $items -join '; '
}

function Summarize-Conditions {
    [CmdletBinding()]
    param([object]$Rule)

    # Build a human-readable conditions summary from selected common transport rule conditions.
    $conditions = [System.Collections.Generic.List[string]]::new()

    $conditionalFields = @(
        @{ Property = 'FromScope';                  Label = 'FromScope' }
        @{ Property = 'SentToScope';                Label = 'SentToScope' }
        @{ Property = 'RecipientDomainIs';          Label = 'RecipientDomain' }
        @{ Property = 'SenderDomainIs';             Label = 'SenderDomain' }
        @{ Property = 'AnyOfRecipientAddressContainsWords'; Label = 'RecipientAddressContains' }
        @{ Property = 'SubjectContainsWords';       Label = 'SubjectContains' }
        @{ Property = 'HeaderContainsMessageHeader'; Label = 'HeaderContains' }
    )

    foreach ($field in $conditionalFields) {
        $val = $Rule.PSObject.Properties[$field.Property]
        if ($val -and -not [string]::IsNullOrWhiteSpace([string]$val.Value)) {
            $conditions.Add("$($field.Label)=$([string]$val.Value)")
        }
    }

    if ($conditions.Count -eq 0) { return '(all messages)' }
    return $conditions -join '; '
}

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'RuleName',
    'Priority',
    'RuleEnabled',
    'ConditionsSummary',
    'ForwardingActionType',
    'ForwardingActionDetail'
)

Write-Status -Message 'Starting Exchange Online transport rule forwarding audit script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

Write-Status -Message 'Fetching all transport rules.'
$allRules = Invoke-WithRetry -OperationName 'Get transport rules' -ScriptBlock {
    Get-TransportRule -ResultSize Unlimited -ErrorAction Stop
}
Write-Status -Message "Fetched $($allRules.Count) transport rules. Filtering for forwarding actions."

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($rule in ($allRules | Sort-Object -Property Priority)) {
    $ruleName   = Get-TrimmedValue -Value $rule.Name
    $primaryKey = $ruleName

    try {
        # Detect forwarding-type actions on this rule.
        $forwardingActions = [System.Collections.Generic.List[object]]::new()

        if ($rule.PSObject.Properties['RedirectMessageTo'] -and $rule.RedirectMessageTo) {
            $detail = Resolve-AddressList -Value $rule.RedirectMessageTo
            if (-not [string]::IsNullOrWhiteSpace($detail)) {
                $forwardingActions.Add([ordered]@{ Type = 'RedirectMessageTo'; Detail = $detail })
            }
        }

        if ($rule.PSObject.Properties['AddToRecipients'] -and $rule.AddToRecipients) {
            $detail = Resolve-AddressList -Value $rule.AddToRecipients
            if (-not [string]::IsNullOrWhiteSpace($detail)) {
                $forwardingActions.Add([ordered]@{ Type = 'AddToRecipients (Forward)'; Detail = $detail })
            }
        }

        if ($rule.PSObject.Properties['CopyTo'] -and $rule.CopyTo) {
            $detail = Resolve-AddressList -Value $rule.CopyTo
            if (-not [string]::IsNullOrWhiteSpace($detail)) {
                $forwardingActions.Add([ordered]@{ Type = 'CopyTo (BCC)'; Detail = $detail })
            }
        }

        if ($rule.PSObject.Properties['BlindCopyTo'] -and $rule.BlindCopyTo) {
            $detail = Resolve-AddressList -Value $rule.BlindCopyTo
            if (-not [string]::IsNullOrWhiteSpace($detail)) {
                $forwardingActions.Add([ordered]@{ Type = 'BlindCopyTo (BCC)'; Detail = $detail })
            }
        }

        # Skip rules with no forwarding actions.
        if ($forwardingActions.Count -eq 0) {
            $rowNumber++
            continue
        }

        $conditionsSummary = Summarize-Conditions -Rule $rule

        foreach ($fa in $forwardingActions) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "${primaryKey}|$([string]$fa.Type)" -Action 'GetTransportRuleForwarding' -Status 'Completed' -Message 'Transport rule forwarding action exported.' -Data ([ordered]@{
                RuleName               = $ruleName
                Priority               = [string]$rule.Priority
                RuleEnabled            = [string]$rule.Enabled
                ConditionsSummary      = $conditionsSummary
                ForwardingActionType   = [string]$fa.Type
                ForwardingActionDetail = [string]$fa.Detail
            })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'GetTransportRuleForwarding' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
            RuleName               = $ruleName
            Priority               = ''
            RuleEnabled            = ''
            ConditionsSummary      = ''
            ForwardingActionType   = ''
            ForwardingActionDetail = ''
        })))
    }

    $rowNumber++
}

if ($results.Count -eq 0) {
    Write-Status -Message 'No transport rules with forwarding actions were found.'
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue 'TenantLevel' -Force
}

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online transport rule forwarding audit script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
