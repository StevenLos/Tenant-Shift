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

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_P3119-Create-ExchangeOnlineDynamicDistributionGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

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

Write-Status -Message 'Starting Exchange Online dynamic distribution group creation script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$setDynamicCommand = Get-Command -Name Set-DynamicDistributionGroup -ErrorAction Stop
$supportsSendModerationNotifications = $setDynamicCommand.Parameters.ContainsKey('SendModerationNotifications')

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

        $displayName = ([string]$row.DisplayName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $createParams.DisplayName = $displayName
        }

        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $createParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $managedBy = ConvertTo-Array -Value ([string]$row.ManagedBy)
        if ($managedBy.Count -gt 0) {
            $createParams.ManagedBy = $managedBy
        }

        $recipientFilter = ([string]$row.RecipientFilter).Trim()
        $includedRecipients = ([string]$row.IncludedRecipients).Trim()
        $conditionalCompany = ConvertTo-Array -Value ([string]$row.ConditionalCompany)
        $conditionalDepartment = ConvertTo-Array -Value ([string]$row.ConditionalDepartment)
        $conditionalCustomAttribute1 = ConvertTo-Array -Value ([string]$row.ConditionalCustomAttribute1)
        $conditionalCustomAttribute2 = ConvertTo-Array -Value ([string]$row.ConditionalCustomAttribute2)
        $conditionalStateOrProvince = ConvertTo-Array -Value ([string]$row.ConditionalStateOrProvince)

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

        if ($PSCmdlet.ShouldProcess($identityToCheck, 'Create Exchange Online dynamic distribution group')) {
            $createdGroup = Invoke-WithRetry -OperationName "Create dynamic distribution group $identityToCheck" -ScriptBlock {
                New-DynamicDistributionGroup @createParams -ErrorAction Stop
            }

            $setParams = @{
                Identity = $createdGroup.Identity
            }

            $requireAuthRaw = ([string]$row.RequireSenderAuthenticationEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($requireAuthRaw)) {
                $setParams.RequireSenderAuthenticationEnabled = ConvertTo-Bool -Value $requireAuthRaw
            }

            $hiddenRaw = ([string]$row.HiddenFromAddressListsEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
                $setParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
            }

            $moderationEnabledRaw = ([string]$row.ModerationEnabled).Trim()
            if (-not [string]::IsNullOrWhiteSpace($moderationEnabledRaw)) {
                $setParams.ModerationEnabled = ConvertTo-Bool -Value $moderationEnabledRaw
            }

            $moderatedBy = ConvertTo-Array -Value ([string]$row.ModeratedBy)
            if ($moderatedBy.Count -gt 0) {
                $setParams.ModeratedBy = $moderatedBy
            }

            $sendModerationNotifications = ([string]$row.SendModerationNotifications).Trim()
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
Write-Status -Message 'Exchange Online dynamic distribution group creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





