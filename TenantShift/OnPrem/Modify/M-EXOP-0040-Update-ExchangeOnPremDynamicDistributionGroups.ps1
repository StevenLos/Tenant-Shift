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

.SYNOPSIS
    Modifies ExchangeOnPremDynamicDistributionGroups in Active Directory.

.DESCRIPTION
    Updates ExchangeOnPremDynamicDistributionGroups in Active Directory based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M0223-Update-ExchangeOnPremDynamicDistributionGroups.ps1 -InputCsvPath .\0223.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M0223-Update-ExchangeOnPremDynamicDistributionGroups.ps1 -InputCsvPath .\0223.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Exchange Management Shell cmdlets (session)
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    Column                              Type      Required  Description
    ----------------------------------  ----      --------  -----------
    DynamicDistributionGroupIdentity    String    Yes       <fill in description>
    DisplayName                         String    Yes       <fill in description>
    PrimarySmtpAddress                  String    Yes       <fill in description>
    ManagedBy                           String    Yes       <fill in description>
    RecipientFilter                     String    Yes       <fill in description>
    IncludedRecipients                  String    Yes       <fill in description>
    ConditionalCompany                  String    Yes       <fill in description>
    ConditionalDepartment               String    Yes       <fill in description>
    ConditionalCustomAttribute1         String    Yes       <fill in description>
    ConditionalCustomAttribute2         String    Yes       <fill in description>
    ConditionalStateOrProvince          String    Yes       <fill in description>
    RequireSenderAuthenticationEnabled  String    Yes       <fill in description>
    HiddenFromAddressListsEnabled       String    Yes       <fill in description>
    ModerationEnabled                   String    Yes       <fill in description>
    ModeratedBy                         String    Yes       <fill in description>
    SendModerationNotifications         String    Yes       <fill in description>
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M0223-Update-ExchangeOnPremDynamicDistributionGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'DynamicDistributionGroupIdentity',
    'DisplayName',
    'PrimarySmtpAddress',
    'ManagedBy',
    'RecipientFilter',
    'IncludedRecipients',
    'ConditionalCompany',
    'ConditionalDepartment',
    'ConditionalCustomAttribute1',
    'ConditionalCustomAttribute2',
    'ConditionalStateOrProvince',
    'RequireSenderAuthenticationEnabled',
    'HiddenFromAddressListsEnabled',
    'ModerationEnabled',
    'ModeratedBy',
    'SendModerationNotifications'
)

Write-Status -Message 'Starting Exchange on-prem dynamic distribution group update script.'
Ensure-ExchangeOnPremConnection

$setDynamicCommand = Get-Command -Name Set-DynamicDistributionGroup -ErrorAction Stop
$supportsSendModerationNotifications = $setDynamicCommand.Parameters.ContainsKey('SendModerationNotifications')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $dynamicDistributionGroupIdentity = Get-TrimmedValue -Value $row.DynamicDistributionGroupIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($dynamicDistributionGroupIdentity)) {
            throw 'DynamicDistributionGroupIdentity is required.'
        }

        $group = Invoke-WithRetry -OperationName "Lookup dynamic distribution group $dynamicDistributionGroupIdentity" -ScriptBlock {
            Get-DynamicDistributionGroup -Identity $dynamicDistributionGroupIdentity -ErrorAction SilentlyContinue
        }
        if (-not $group) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $dynamicDistributionGroupIdentity -Action 'UpdateDynamicDistributionGroup' -Status 'NotFound' -Message 'Dynamic distribution group not found.'))
            $rowNumber++
            continue
        }

        $setParams = @{
            Identity = $group.Identity
        }

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $setParams.DisplayName = $displayName
        }

        $primarySmtpAddress = Get-TrimmedValue -Value $row.PrimarySmtpAddress
        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $setParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $managedBy = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ManagedBy)
        if ($managedBy.Count -gt 0) {
            $setParams.ManagedBy = $managedBy
        }

        $recipientFilter = Get-TrimmedValue -Value $row.RecipientFilter
        if (-not [string]::IsNullOrWhiteSpace($recipientFilter)) {
            $setParams.RecipientFilter = $recipientFilter
        }
        else {
            $includedRecipients = Get-TrimmedValue -Value $row.IncludedRecipients
            if (-not [string]::IsNullOrWhiteSpace($includedRecipients)) {
                $setParams.IncludedRecipients = $includedRecipients
            }

            $conditionalCompany = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ConditionalCompany)
            if ($conditionalCompany.Count -gt 0) {
                $setParams.ConditionalCompany = $conditionalCompany
            }

            $conditionalDepartment = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ConditionalDepartment)
            if ($conditionalDepartment.Count -gt 0) {
                $setParams.ConditionalDepartment = $conditionalDepartment
            }

            $conditionalCustomAttribute1 = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ConditionalCustomAttribute1)
            if ($conditionalCustomAttribute1.Count -gt 0) {
                $setParams.ConditionalCustomAttribute1 = $conditionalCustomAttribute1
            }

            $conditionalCustomAttribute2 = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ConditionalCustomAttribute2)
            if ($conditionalCustomAttribute2.Count -gt 0) {
                $setParams.ConditionalCustomAttribute2 = $conditionalCustomAttribute2
            }

            $conditionalStateOrProvince = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ConditionalStateOrProvince)
            if ($conditionalStateOrProvince.Count -gt 0) {
                $setParams.ConditionalStateOrProvince = $conditionalStateOrProvince
            }
        }

        $requireAuthRaw = Get-TrimmedValue -Value $row.RequireSenderAuthenticationEnabled
        if (-not [string]::IsNullOrWhiteSpace($requireAuthRaw)) {
            $setParams.RequireSenderAuthenticationEnabled = ConvertTo-Bool -Value $requireAuthRaw
        }

        $hiddenRaw = Get-TrimmedValue -Value $row.HiddenFromAddressListsEnabled
        if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
            $setParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
        }

        $moderationEnabledRaw = Get-TrimmedValue -Value $row.ModerationEnabled
        if (-not [string]::IsNullOrWhiteSpace($moderationEnabledRaw)) {
            $setParams.ModerationEnabled = ConvertTo-Bool -Value $moderationEnabledRaw
        }

        $moderatedBy = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ModeratedBy)
        if ($moderatedBy.Count -gt 0) {
            $setParams.ModeratedBy = $moderatedBy
        }

        $sendModerationNotifications = Get-TrimmedValue -Value $row.SendModerationNotifications
        if (-not [string]::IsNullOrWhiteSpace($sendModerationNotifications)) {
            if ($supportsSendModerationNotifications) {
                $setParams.SendModerationNotifications = $sendModerationNotifications
            }
            else {
                Write-Status -Message 'Set-DynamicDistributionGroup does not support -SendModerationNotifications in this session. Value ignored.' -Level WARN
            }
        }

        if ($setParams.Count -eq 1) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $dynamicDistributionGroupIdentity -Action 'UpdateDynamicDistributionGroup' -Status 'Skipped' -Message 'No updates specified.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($dynamicDistributionGroupIdentity, 'Update Exchange on-prem dynamic distribution group')) {
            Invoke-WithRetry -OperationName "Update dynamic distribution group $dynamicDistributionGroupIdentity" -ScriptBlock {
                Set-DynamicDistributionGroup @setParams -ErrorAction Stop
            }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $dynamicDistributionGroupIdentity -Action 'UpdateDynamicDistributionGroup' -Status 'Updated' -Message 'Dynamic distribution group updated successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $dynamicDistributionGroupIdentity -Action 'UpdateDynamicDistributionGroup' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($dynamicDistributionGroupIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $dynamicDistributionGroupIdentity -Action 'UpdateDynamicDistributionGroup' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem dynamic distribution group update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
