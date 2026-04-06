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
    Provisions ExchangeOnPremMailEnabledSecurityGroups in Active Directory.

.DESCRIPTION
    Creates ExchangeOnPremMailEnabledSecurityGroups in Active Directory based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P0215-Create-ExchangeOnPremMailEnabledSecurityGroups.ps1 -InputCsvPath .\0215.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P0215-Create-ExchangeOnPremMailEnabledSecurityGroups.ps1 -InputCsvPath .\0215.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Exchange Management Shell cmdlets (session)
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    Column                                Type      Required  Description
    ------------------------------------  ----      --------  -----------
    Name                                  String    Yes       <fill in description>
    Alias                                 String    Yes       <fill in description>
    DisplayName                           String    Yes       <fill in description>
    PrimarySmtpAddress                    String    Yes       <fill in description>
    ManagedBy                             String    Yes       <fill in description>
    Notes                                 String    Yes       <fill in description>
    RequireSenderAuthenticationEnabled    String    Yes       <fill in description>
    HiddenFromAddressListsEnabled         String    Yes       <fill in description>
    ModerationEnabled                     String    Yes       <fill in description>
    ModeratedBy                           String    Yes       <fill in description>
    AcceptMessagesOnlyFrom                String    Yes       <fill in description>
    AcceptMessagesOnlyFromDLMembers       String    Yes       <fill in description>
    RejectMessagesFrom                    String    Yes       <fill in description>
    RejectMessagesFromDLMembers           String    Yes       <fill in description>
    BypassModerationFromSendersOrMembers  String    Yes       <fill in description>
    SendModerationNotifications           String    Yes       <fill in description>
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P0215-Create-ExchangeOnPremMailEnabledSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
    'Notes',
    'RequireSenderAuthenticationEnabled',
    'HiddenFromAddressListsEnabled',
    'ModerationEnabled',
    'ModeratedBy',
    'AcceptMessagesOnlyFrom',
    'AcceptMessagesOnlyFromDLMembers',
    'RejectMessagesFrom',
    'RejectMessagesFromDLMembers',
    'BypassModerationFromSendersOrMembers',
    'SendModerationNotifications'
)

Write-Status -Message 'Starting Exchange on-prem mail-enabled security group creation script.'
Ensure-ExchangeOnPremConnection

$setDistributionGroupCommand = Get-Command -Name Set-DistributionGroup -ErrorAction Stop
$supportsSendModerationNotifications = $setDistributionGroupCommand.Parameters.ContainsKey('SendModerationNotifications')
$newDistributionGroupCommand = Get-Command -Name New-DistributionGroup -ErrorAction Stop
$supportsOrganizationalUnit = $newDistributionGroupCommand.Parameters.ContainsKey('OrganizationalUnit')

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

        $existingGroup = Invoke-WithRetry -OperationName "Lookup mail-enabled security group $identityToCheck" -ScriptBlock {
            Get-DistributionGroup -Identity $identityToCheck -ErrorAction SilentlyContinue
        }

        if ($existingGroup) {
            $recipientTypeDetails = Get-TrimmedValue -Value $existingGroup.RecipientTypeDetails
            if ($recipientTypeDetails -eq 'MailUniversalSecurityGroup') {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateMailEnabledSecurityGroup' -Status 'Skipped' -Message 'Mail-enabled security group already exists.'))
                $rowNumber++
                continue
            }

            throw "Recipient '$identityToCheck' already exists with type '$recipientTypeDetails'."
        }

        $createParams = @{
            Name  = $name
            Alias = $alias
            Type  = 'Security'
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

        $organizationalUnit = Get-OptionalColumnValue -Row $row -ColumnName 'OrganizationalUnit'
        if ($supportsOrganizationalUnit -and -not [string]::IsNullOrWhiteSpace($organizationalUnit)) {
            $createParams.OrganizationalUnit = $organizationalUnit
        }

        if ($PSCmdlet.ShouldProcess($identityToCheck, 'Create Exchange on-prem mail-enabled security group')) {
            $createdGroup = Invoke-WithRetry -OperationName "Create mail-enabled security group $identityToCheck" -ScriptBlock {
                New-DistributionGroup @createParams -ErrorAction Stop
            }

            $setParams = @{
                Identity = $createdGroup.Identity
            }

            $notes = Get-TrimmedValue -Value $row.Notes
            if (-not [string]::IsNullOrWhiteSpace($notes)) {
                $setParams.Notes = $notes
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

            $acceptMessagesOnlyFrom = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.AcceptMessagesOnlyFrom)
            if ($acceptMessagesOnlyFrom.Count -gt 0) {
                $setParams.AcceptMessagesOnlyFrom = $acceptMessagesOnlyFrom
            }

            $acceptMessagesOnlyFromDLMembers = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.AcceptMessagesOnlyFromDLMembers)
            if ($acceptMessagesOnlyFromDLMembers.Count -gt 0) {
                $setParams.AcceptMessagesOnlyFromDLMembers = $acceptMessagesOnlyFromDLMembers
            }

            $rejectMessagesFrom = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.RejectMessagesFrom)
            if ($rejectMessagesFrom.Count -gt 0) {
                $setParams.RejectMessagesFrom = $rejectMessagesFrom
            }

            $rejectMessagesFromDLMembers = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.RejectMessagesFromDLMembers)
            if ($rejectMessagesFromDLMembers.Count -gt 0) {
                $setParams.RejectMessagesFromDLMembers = $rejectMessagesFromDLMembers
            }

            $bypassModeration = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.BypassModerationFromSendersOrMembers)
            if ($bypassModeration.Count -gt 0) {
                $setParams.BypassModerationFromSendersOrMembers = $bypassModeration
            }

            $sendModerationNotifications = Get-TrimmedValue -Value $row.SendModerationNotifications
            if (-not [string]::IsNullOrWhiteSpace($sendModerationNotifications)) {
                if ($supportsSendModerationNotifications) {
                    $setParams.SendModerationNotifications = $sendModerationNotifications
                }
                else {
                    Write-Status -Message 'Set-DistributionGroup does not support -SendModerationNotifications in this session. Value ignored.' -Level WARN
                }
            }

            if ($setParams.Count -gt 1) {
                Invoke-WithRetry -OperationName "Set mail-enabled security group options $identityToCheck" -ScriptBlock {
                    Set-DistributionGroup @setParams -ErrorAction Stop
                }
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateMailEnabledSecurityGroup' -Status 'Created' -Message 'Mail-enabled security group created successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $identityToCheck -Action 'CreateMailEnabledSecurityGroup' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($name) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $name -Action 'CreateMailEnabledSecurityGroup' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem mail-enabled security group creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
