<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-174500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineMailEnabledSecurityGroupMembers and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineMailEnabledSecurityGroupMembers from Microsoft 365 and writes the results to a CSV file.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all objects in scope (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all objects in scope rather than processing from an input CSV file. Uses the DiscoverAll parameter set.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-IR3133-Get-ExchangeOnlineMailEnabledSecurityGroupMembers.ps1 -InputCsvPath .\3133.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-IR3133-Get-ExchangeOnlineMailEnabledSecurityGroupMembers.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator
    Limitations:      None known.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_D-EXOL-0100-Get-ExchangeOnlineMailEnabledSecurityGroupMembers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-ObjectPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($property) {
        return $property.Value
    }

    return $null
}

function Get-StringPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    return ([string](Get-ObjectPropertyValue -InputObject $InputObject -PropertyName $PropertyName)).Trim()
}

$requiredHeaders = @(
    'SecurityGroupIdentity'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'SecurityGroupIdentity',
    'SecurityGroupDisplayName',
    'MemberDisplayName',
    'MemberIdentity',
    'MemberPrimarySmtpAddress',
    'MemberRecipientType',
    'MemberExternalEmailAddress'
)

Write-Status -Message 'Starting Exchange Online mail-enabled security group member inventory script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$scopeMode = 'Csv'
if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. CSV input is bypassed.' -Level WARN

    $discoverRow = [ordered]@{}
    foreach ($header in $requiredHeaders) {
        $discoverRow[$header] = '*'
    }

    $rows = @([PSCustomObject]$discoverRow)
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $securityGroupIdentityInput = ([string]$row.SecurityGroupIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($securityGroupIdentityInput)) {
            throw 'SecurityGroupIdentity is required. Use * to inventory members for all mail-enabled security groups.'
        }

        $groups = @()
        if ($securityGroupIdentityInput -eq '*') {
            $groups = @(Invoke-WithRetry -OperationName 'Load all mail-enabled security groups for membership export' -ScriptBlock {
                Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop |
                    Where-Object { (Get-StringPropertyValue -InputObject $_ -PropertyName 'RecipientTypeDetails') -eq 'MailUniversalSecurityGroup' }
            })
        }
        else {
            $group = Invoke-WithRetry -OperationName "Lookup mail-enabled security group $securityGroupIdentityInput" -ScriptBlock {
                Get-DistributionGroup -Identity $securityGroupIdentityInput -ErrorAction SilentlyContinue
            }

            if ($group -and (Get-StringPropertyValue -InputObject $group -PropertyName 'RecipientTypeDetails') -eq 'MailUniversalSecurityGroup') {
                $groups = @($group)
            }
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $securityGroupIdentityInput -Action 'GetExchangeMailEnabledSecurityGroupMember' -Status 'NotFound' -Message 'No matching mail-enabled security groups were found.' -Data ([ordered]@{
                            SecurityGroupIdentity         = $securityGroupIdentityInput
                            SecurityGroupDisplayName      = ''
                            MemberIdentity                = ''
                            MemberDisplayName             = ''
                            MemberPrimarySmtpAddress      = ''
                            MemberRecipientType           = ''
                            MemberExternalEmailAddress    = ''
                        })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Identity)) {
            $groupIdentity = Get-StringPropertyValue -InputObject $group -PropertyName 'Identity'
            $groupDisplayName = Get-StringPropertyValue -InputObject $group -PropertyName 'DisplayName'

            $members = @(Invoke-WithRetry -OperationName "Load members for mail-enabled security group $groupIdentity" -ScriptBlock {
                Get-DistributionGroupMember -Identity $groupIdentity -ResultSize Unlimited -ErrorAction Stop
            })

            if ($members.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $groupIdentity -Action 'GetExchangeMailEnabledSecurityGroupMember' -Status 'Completed' -Message 'Mail-enabled security group has no members.' -Data ([ordered]@{
                                SecurityGroupIdentity         = $groupIdentity
                                SecurityGroupDisplayName      = $groupDisplayName
                                MemberIdentity                = ''
                                MemberDisplayName             = ''
                                MemberPrimarySmtpAddress      = ''
                                MemberRecipientType           = ''
                                MemberExternalEmailAddress    = ''
                            })))
                continue
            }

            foreach ($member in @($members | Sort-Object -Property DisplayName, Identity)) {
                $memberIdentity = Get-StringPropertyValue -InputObject $member -PropertyName 'Identity'
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey "$groupIdentity|$memberIdentity" -Action 'GetExchangeMailEnabledSecurityGroupMember' -Status 'Completed' -Message 'Mail-enabled security group member exported.' -Data ([ordered]@{
                                SecurityGroupIdentity         = $groupIdentity
                                SecurityGroupDisplayName      = $groupDisplayName
                                MemberIdentity                = $memberIdentity
                                MemberDisplayName             = Get-StringPropertyValue -InputObject $member -PropertyName 'DisplayName'
                                MemberPrimarySmtpAddress      = Get-StringPropertyValue -InputObject $member -PropertyName 'PrimarySmtpAddress'
                                MemberRecipientType           = Get-StringPropertyValue -InputObject $member -PropertyName 'RecipientType'
                                MemberExternalEmailAddress    = Get-StringPropertyValue -InputObject $member -PropertyName 'ExternalEmailAddress'
                            })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($securityGroupIdentityInput) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $securityGroupIdentityInput -Action 'GetExchangeMailEnabledSecurityGroupMember' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                        SecurityGroupIdentity         = $securityGroupIdentityInput
                        SecurityGroupDisplayName      = ''
                        MemberIdentity                = ''
                        MemberDisplayName             = ''
                        MemberPrimarySmtpAddress      = ''
                        MemberRecipientType           = ''
                        MemberExternalEmailAddress    = ''
                    })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mail-enabled security group member inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
