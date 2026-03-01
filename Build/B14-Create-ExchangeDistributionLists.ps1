#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath ("Results_B14-Create-ExchangeDistributionLists_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'M365.Common.psm1') -Force -DisableNameChecking

$requiredHeaders = @(
    'Name',
    'Alias',
    'DisplayName',
    'PrimarySmtpAddress',
    'ManagedBy',
    'Notes',
    'MemberJoinRestriction',
    'MemberDepartRestriction',
    'ModerationEnabled',
    'ModeratedBy',
    'RequireSenderAuthenticationEnabled',
    'HiddenFromAddressListsEnabled'
)

Write-Status -Message 'Starting Exchange Online distribution list creation script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $name = ([string]$row.Name).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Name is required.'
        }

        $alias = ([string]$row.Alias).Trim()
        if ([string]::IsNullOrWhiteSpace($alias)) {
            $alias = ($name -replace '[^a-zA-Z0-9]', '')
            if ([string]::IsNullOrWhiteSpace($alias)) {
                throw 'Alias is empty after sanitization. Provide an Alias value in the CSV.'
            }
        }

        $primarySmtpAddress = ([string]$row.PrimarySmtpAddress).Trim()
        $identityToCheck = if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) { $primarySmtpAddress } else { $name }

        $existingGroup = Invoke-WithRetry -OperationName "Lookup distribution list $identityToCheck" -ScriptBlock {
            Get-DistributionGroup -Identity $identityToCheck -ErrorAction SilentlyContinue
        }
        if ($existingGroup) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateDistributionList' -Status 'Skipped' -Message 'Distribution list already exists.'))
            $rowNumber++
            continue
        }

        $params = @{
            Name  = $name
            Alias = $alias
        }

        $displayName = ([string]$row.DisplayName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $params.DisplayName = $displayName
        }

        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $params.PrimarySmtpAddress = $primarySmtpAddress
        }

        $managedBy = ConvertTo-Array -Value ([string]$row.ManagedBy)
        if ($managedBy.Count -gt 0) {
            $params.ManagedBy = $managedBy
        }

        $notes = ([string]$row.Notes).Trim()
        if (-not [string]::IsNullOrWhiteSpace($notes)) {
            $params.Notes = $notes
        }

        $memberJoinRestriction = ([string]$row.MemberJoinRestriction).Trim()
        if (-not [string]::IsNullOrWhiteSpace($memberJoinRestriction)) {
            $params.MemberJoinRestriction = $memberJoinRestriction
        }

        $memberDepartRestriction = ([string]$row.MemberDepartRestriction).Trim()
        if (-not [string]::IsNullOrWhiteSpace($memberDepartRestriction)) {
            $params.MemberDepartRestriction = $memberDepartRestriction
        }

        $moderationEnabledRaw = ([string]$row.ModerationEnabled).Trim()
        if (-not [string]::IsNullOrWhiteSpace($moderationEnabledRaw)) {
            $params.ModerationEnabled = ConvertTo-Bool -Value $moderationEnabledRaw
        }

        $moderatedBy = ConvertTo-Array -Value ([string]$row.ModeratedBy)
        if ($moderatedBy.Count -gt 0) {
            $params.ModeratedBy = $moderatedBy
        }

        $requireAuthRaw = ([string]$row.RequireSenderAuthenticationEnabled).Trim()
        if (-not [string]::IsNullOrWhiteSpace($requireAuthRaw)) {
            $params.RequireSenderAuthenticationEnabled = ConvertTo-Bool -Value $requireAuthRaw
        }

        $hiddenRaw = ([string]$row.HiddenFromAddressListsEnabled).Trim()
        $setHidden = -not [string]::IsNullOrWhiteSpace($hiddenRaw)

        if ($PSCmdlet.ShouldProcess($identityToCheck, 'Create Exchange Online distribution list')) {
            $createdGroup = Invoke-WithRetry -OperationName "Create distribution list $identityToCheck" -ScriptBlock {
                New-DistributionGroup @params -ErrorAction Stop
            }

            if ($setHidden) {
                $hiddenValue = ConvertTo-Bool -Value $hiddenRaw
                Invoke-WithRetry -OperationName "Set hidden from GAL for $identityToCheck" -ScriptBlock {
                    Set-DistributionGroup -Identity $createdGroup.Identity -HiddenFromAddressListsEnabled $hiddenValue -ErrorAction Stop
                }
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateDistributionList' -Status 'Created' -Message 'Distribution list created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateDistributionList' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($name) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $name -Action 'CreateDistributionList' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online distribution list creation script completed.' -Level SUCCESS

