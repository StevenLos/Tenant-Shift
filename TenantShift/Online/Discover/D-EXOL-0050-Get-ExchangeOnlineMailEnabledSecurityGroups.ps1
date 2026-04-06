<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-174000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineMailEnabledSecurityGroups and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineMailEnabledSecurityGroups from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-D3122-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1 -InputCsvPath .\3122.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D3122-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1 -DiscoverAll

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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-EXOL-0050-Get-ExchangeOnlineMailEnabledSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Convert-GroupMembersToString {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$Members
    )

    $resolvedMembers = [System.Collections.Generic.List[string]]::new()

    foreach ($member in @($Members)) {
        if ($null -eq $member) {
            continue
        }

        $resolvedValue = ''
        foreach ($candidate in @(
                if ($member -is [string]) { ([string]$member).Trim() } else { Get-StringPropertyValue -InputObject $member -PropertyName 'PrimarySmtpAddress' }
                if ($member -is [string]) { '' } else { Get-StringPropertyValue -InputObject $member -PropertyName 'WindowsEmailAddress' }
                if ($member -is [string]) { '' } else { Get-StringPropertyValue -InputObject $member -PropertyName 'ExternalEmailAddress' }
                if ($member -is [string]) { '' } else { Get-StringPropertyValue -InputObject $member -PropertyName 'Identity' }
                if ($member -is [string]) { '' } else { Get-StringPropertyValue -InputObject $member -PropertyName 'DisplayName' }
                if ($member -is [string]) { '' } else { Get-StringPropertyValue -InputObject $member -PropertyName 'Name' }
            )) {
            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                $resolvedValue = $candidate
                break
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($resolvedValue)) {
            $resolvedMembers.Add($resolvedValue)
        }
    }

    return Convert-MultiValueToString -Value $resolvedMembers
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
    'DisplayName',
    'Name',
    'Alias',
    'PrimarySmtpAddress',
    'ManagedBy',
    'Members',
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

Write-Status -Message 'Starting Exchange Online mail-enabled security group inventory script.'
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
    $securityGroupIdentity = ([string]$row.SecurityGroupIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($securityGroupIdentity)) {
            throw 'SecurityGroupIdentity is required. Use * to inventory all mail-enabled security groups.'
        }

        $groups = @()
        if ($securityGroupIdentity -eq '*') {
            $groups = @(Invoke-WithRetry -OperationName 'Load all mail-enabled security groups' -ScriptBlock {
                Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop | Where-Object { ([string]$_.RecipientTypeDetails).Trim() -eq 'MailUniversalSecurityGroup' }
            })
        }
        else {
            $group = Invoke-WithRetry -OperationName "Lookup mail-enabled security group $securityGroupIdentity" -ScriptBlock {
                Get-DistributionGroup -Identity $securityGroupIdentity -ErrorAction SilentlyContinue
            }
            if ($group -and ([string]$group.RecipientTypeDetails).Trim() -eq 'MailUniversalSecurityGroup') {
                $groups = @($group)
            }
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $securityGroupIdentity -Action 'GetExchangeMailEnabledSecurityGroup' -Status 'NotFound' -Message 'No matching mail-enabled security groups were found.' -Data ([ordered]@{
                        SecurityGroupIdentity                   = $securityGroupIdentity
                        Name                                    = ''
                        Alias                                   = ''
                        DisplayName                             = ''
                        PrimarySmtpAddress                      = ''
                        ManagedBy                               = ''
                        Members                                 = ''
                        Notes                                   = ''
                        RequireSenderAuthenticationEnabled      = ''
                        HiddenFromAddressListsEnabled           = ''
                        ModerationEnabled                       = ''
                        ModeratedBy                             = ''
                        AcceptMessagesOnlyFrom                  = ''
                        AcceptMessagesOnlyFromDLMembers         = ''
                        RejectMessagesFrom                      = ''
                        RejectMessagesFromDLMembers             = ''
                        BypassModerationFromSendersOrMembers    = ''
                        SendModerationNotifications             = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Identity)) {
            $identity = Get-StringPropertyValue -InputObject $group -PropertyName 'Identity'
            $members = @(Invoke-WithRetry -OperationName "Load members for mail-enabled security group $identity" -ScriptBlock {
                Get-DistributionGroupMember -Identity $identity -ResultSize Unlimited -ErrorAction Stop
            })
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeMailEnabledSecurityGroup' -Status 'Completed' -Message 'Mail-enabled security group exported.' -Data ([ordered]@{
                        SecurityGroupIdentity                   = $identity
                        Name                                    = Get-StringPropertyValue -InputObject $group -PropertyName 'Name'
                        Alias                                   = Get-StringPropertyValue -InputObject $group -PropertyName 'Alias'
                        DisplayName                             = Get-StringPropertyValue -InputObject $group -PropertyName 'DisplayName'
                        PrimarySmtpAddress                      = Get-StringPropertyValue -InputObject $group -PropertyName 'PrimarySmtpAddress'
                        ManagedBy                               = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'ManagedBy')
                        Members                                 = Convert-GroupMembersToString -Members $members
                        Notes                                   = Get-StringPropertyValue -InputObject $group -PropertyName 'Notes'
                        RequireSenderAuthenticationEnabled      = [string](Get-ObjectPropertyValue -InputObject $group -PropertyName 'RequireSenderAuthenticationEnabled')
                        HiddenFromAddressListsEnabled           = [string](Get-ObjectPropertyValue -InputObject $group -PropertyName 'HiddenFromAddressListsEnabled')
                        ModerationEnabled                       = [string](Get-ObjectPropertyValue -InputObject $group -PropertyName 'ModerationEnabled')
                        ModeratedBy                             = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'ModeratedBy')
                        AcceptMessagesOnlyFrom                  = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'AcceptMessagesOnlyFrom')
                        AcceptMessagesOnlyFromDLMembers         = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'AcceptMessagesOnlyFromDLMembers')
                        RejectMessagesFrom                      = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'RejectMessagesFrom')
                        RejectMessagesFromDLMembers             = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'RejectMessagesFromDLMembers')
                        BypassModerationFromSendersOrMembers    = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'BypassModerationFromSendersOrMembers')
                        SendModerationNotifications             = Get-StringPropertyValue -InputObject $group -PropertyName 'SendModerationNotifications'
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($securityGroupIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $securityGroupIdentity -Action 'GetExchangeMailEnabledSecurityGroup' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    SecurityGroupIdentity                   = $securityGroupIdentity
                    Name                                    = ''
                    Alias                                   = ''
                    DisplayName                             = ''
                    PrimarySmtpAddress                      = ''
                    ManagedBy                               = ''
                    Members                                 = ''
                    Notes                                   = ''
                    RequireSenderAuthenticationEnabled      = ''
                    HiddenFromAddressListsEnabled           = ''
                    ModerationEnabled                       = ''
                    ModeratedBy                             = ''
                    AcceptMessagesOnlyFrom                  = ''
                    AcceptMessagesOnlyFromDLMembers         = ''
                    RejectMessagesFrom                      = ''
                    RejectMessagesFromDLMembers             = ''
                    BypassModerationFromSendersOrMembers    = ''
                    SendModerationNotifications             = ''
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
Write-Status -Message 'Exchange Online mail-enabled security group inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}








