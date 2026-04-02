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
    Provisions ExchangeOnPremDynamicDistributionGroups in Active Directory.

.DESCRIPTION
    Creates ExchangeOnPremDynamicDistributionGroups in Active Directory based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P0219-Create-ExchangeOnPremDynamicDistributionGroups.ps1 -InputCsvPath .\0219.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P0219-Create-ExchangeOnPremDynamicDistributionGroups.ps1 -InputCsvPath .\0219.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Exchange Management Shell cmdlets (session)
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    Column                              Type      Required  Description
    ----------------------------------  ----      --------  -----------
    Name                                String    Yes       <fill in description>
    Alias                               String    Yes       <fill in description>
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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P0219-Create-ExchangeOnPremDynamicDistributionGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-OptionalColumnValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [psobject]$Row,

        [Parameter(Mandatory)]
        [string]$ColumnName
    )

    $property = $Row.PSObject.Properties[$ColumnName]
    if ($null -eq $property) {
        return ''
    }

    return Get-TrimmedValue -Value $property.Value
}

$requiredHeaders = @(
    'Name',
    'Alias',
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

Write-Status -Message 'Starting Exchange on-prem dynamic distribution group creation script.'
Ensure-ExchangeOnPremConnection

$setDynamicCommand = Get-Command -Name Set-DynamicDistributionGroup -ErrorAction Stop
$supportsSendModerationNotifications = $setDynamicCommand.Parameters.ContainsKey('SendModerationNotifications')
$newDynamicCommand = Get-Command -Name New-DynamicDistributionGroup -ErrorAction Stop
$supportsOrganizationalUnit = $newDynamicCommand.Parameters.ContainsKey('OrganizationalUnit')
$supportsRecipientContainer = $newDynamicCommand.Parameters.ContainsKey('RecipientContainer')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $name = Get-TrimmedValue -Value $row.Name

    try {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Name is required.'
        }

        $alias = Get-TrimmedValue -Value $row.Alias
        if ([string]::IsNullOrWhiteSpace($alias)) {
            $alias = ($name -replace '[^a-zA-Z0-9]', '')
            if ([string]::IsNullOrWhiteSpace($alias)) {
                throw 'Alias is empty after sanitization. Provide an Alias value in the CSV.'
            }
        }

        $primarySmtpAddress = Get-TrimmedValue -Value $row.PrimarySmtpAddress
        $identityToCheck = if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) { $primarySmtpAddress } else { $name }

        $existingGroup = Invoke-WithRetry -OperationName "Lookup dynamic distribution group $identityToCheck" -ScriptBlock {
            Get-DynamicDistributionGroup -Identity $identityToCheck -ErrorAction SilentlyContinue
        }

        if ($existingGroup) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateDynamicDistributionGroup' -Status 'Skipped' -Message 'Dynamic distribution group already exists.'))
            $rowNumber++
            continue
        }

        $createParams = @{
            Name  = $name
            Alias = $alias
        }

        $displayName = Get-TrimmedValue -Value $row.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $createParams.DisplayName = $displayName
        }

        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $createParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $managedBy = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ManagedBy)
        if ($managedBy.Count -gt 0) {
            $createParams.ManagedBy = $managedBy
        }

        $recipientFilter = Get-TrimmedValue -Value $row.RecipientFilter
        $includedRecipients = Get-TrimmedValue -Value $row.IncludedRecipients
        $conditionalCompany = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ConditionalCompany)
        $conditionalDepartment = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ConditionalDepartment)
        $conditionalCustomAttribute1 = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ConditionalCustomAttribute1)
        $conditionalCustomAttribute2 = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ConditionalCustomAttribute2)
        $conditionalStateOrProvince = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.ConditionalStateOrProvince)

        if (-not [string]::IsNullOrWhiteSpace($recipientFilter)) {
            $createParams.RecipientFilter = $recipientFilter
        }
        else {
            if ([string]::IsNullOrWhiteSpace($includedRecipients) -and $conditionalCompany.Count -eq 0 -and $conditionalDepartment.Count -eq 0 -and $conditionalCustomAttribute1.Count -eq 0 -and $conditionalCustomAttribute2.Count -eq 0 -and $conditionalStateOrProvince.Count -eq 0) {
                throw 'Provide RecipientFilter or IncludedRecipients/Conditional* values.'
            }

            if (-not [string]::IsNullOrWhiteSpace($includedRecipients)) {
                $createParams.IncludedRecipients = $includedRecipients
            }
            if ($conditionalCompany.Count -gt 0) {
                $createParams.ConditionalCompany = $conditionalCompany
            }
            if ($conditionalDepartment.Count -gt 0) {
                $createParams.ConditionalDepartment = $conditionalDepartment
            }
            if ($conditionalCustomAttribute1.Count -gt 0) {
                $createParams.ConditionalCustomAttribute1 = $conditionalCustomAttribute1
            }
            if ($conditionalCustomAttribute2.Count -gt 0) {
                $createParams.ConditionalCustomAttribute2 = $conditionalCustomAttribute2
            }
            if ($conditionalStateOrProvince.Count -gt 0) {
                $createParams.ConditionalStateOrProvince = $conditionalStateOrProvince
            }
        }

        $organizationalUnit = Get-OptionalColumnValue -Row $row -ColumnName 'OrganizationalUnit'
        if ($supportsOrganizationalUnit -and -not [string]::IsNullOrWhiteSpace($organizationalUnit)) {
            $createParams.OrganizationalUnit = $organizationalUnit
        }

        $recipientContainer = Get-OptionalColumnValue -Row $row -ColumnName 'RecipientContainer'
        if ($supportsRecipientContainer -and -not [string]::IsNullOrWhiteSpace($recipientContainer)) {
            $createParams.RecipientContainer = $recipientContainer
        }

        if ($PSCmdlet.ShouldProcess($identityToCheck, 'Create Exchange on-prem dynamic distribution group')) {
            $createdGroup = Invoke-WithRetry -OperationName "Create dynamic distribution group $identityToCheck" -ScriptBlock {
                New-DynamicDistributionGroup @createParams -ErrorAction Stop
            }

            $setParams = @{
                Identity = $createdGroup.Identity
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

            if ($setParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Set dynamic distribution group options $identityToCheck" -ScriptBlock {
                    Set-DynamicDistributionGroup @setParams -ErrorAction Stop
                }
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateDynamicDistributionGroup' -Status 'Created' -Message 'Dynamic distribution group created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateDynamicDistributionGroup' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($name) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $name -Action 'CreateDynamicDistributionGroup' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem dynamic distribution group creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
