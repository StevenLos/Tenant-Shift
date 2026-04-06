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
    Gets ExchangeOnPremMailEnabledSecurityGroups and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnPremMailEnabledSecurityGroups from Active Directory and writes the results to a CSV file.
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all objects in scope (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all objects in scope rather than processing from an input CSV file. Uses the DiscoverAll parameter set.

.PARAMETER SearchBase
    Distinguished name of the Active Directory OU to scope the discovery. If omitted, searches the entire domain.

.PARAMETER Server
    Active Directory domain controller to target. If omitted, uses the default DC for the current domain.

.PARAMETER MaxObjects
    Maximum number of objects to retrieve. 0 (default) means no limit.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-D0222-Get-ExchangeOnPremMailEnabledSecurityGroups.ps1 -InputCsvPath .\0222.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D0222-Get-ExchangeOnPremMailEnabledSecurityGroups.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: Exchange Management Shell cmdlets (session)
    Required roles:   Domain Administrator or delegated OU write permission
    Limitations:      None known.

    CSV Fields:
    See the .input.csv template file in the script directory for the full column list.
    (Use Development\Build\Utilities\Generate-CsvHelpTable.ps1 to regenerate this table from
    the template header row when the template changes.)
#>
#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$SearchBase,

    [Parameter(ParameterSetName = 'FromCsv')]
    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string]$Server,

    [Parameter(ParameterSetName = 'DiscoverAll')]
    [ValidateRange(0, 10000000)]
    [int]$MaxObjects = 0,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_SM-D0222-Get-ExchangeOnPremMailEnabledSecurityGroups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Add-SupportedParameter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$ParameterHashtable,

        [Parameter(Mandatory)]
        [string]$CommandName,

        [Parameter(Mandatory)]
        [string]$ParameterName,

        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return
    }

    $command = Get-Command -Name $CommandName -ErrorAction Stop
    if ($command.Parameters.ContainsKey($ParameterName)) {
        $ParameterHashtable[$ParameterName] = $text
    }
}

function Test-IsMailEnabledSecurityGroupRecipientType {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$RecipientTypeDetails
    )

    $recipientTypeText = Get-TrimmedValue -Value $RecipientTypeDetails
    return ($recipientTypeText -match 'SecurityGroup')
}

function Resolve-SecurityGroupsByScope {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,

        [AllowEmptyString()]
        [string]$SearchBase,

        [AllowEmptyString()]
        [string]$Server
    )

    if ($Identity -eq '*') {
        $params = @{
            ResultSize  = 'Unlimited'
            ErrorAction = 'Stop'
        }

        Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-DistributionGroup' -ParameterName 'OrganizationalUnit' -Value $SearchBase
        Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-DistributionGroup' -ParameterName 'DomainController' -Value $Server

        return @(
            Get-DistributionGroup @params |
                Where-Object { Test-IsMailEnabledSecurityGroupRecipientType -RecipientTypeDetails $_.RecipientTypeDetails }
        )
    }

    $params = @{
        Identity    = $Identity
        ErrorAction = 'SilentlyContinue'
    }

    Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-DistributionGroup' -ParameterName 'DomainController' -Value $Server

    $group = Get-DistributionGroup @params
    if ($group -and (Test-IsMailEnabledSecurityGroupRecipientType -RecipientTypeDetails $group.RecipientTypeDetails)) {
        return @($group)
    }

    return @()
}

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

function Convert-MultiValueToString {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [string]) {
        return ([string]$Value).Trim()
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        $items = [System.Collections.Generic.List[string]]::new()
        foreach ($item in $Value) {
            $text = ([string]$item).Trim()
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $items.Add($text)
            }
        }

        return (@($items | Sort-Object -Unique) -join ';')
    }

    return ([string]$Value).Trim()
}

$requiredHeaders = @(
    'SecurityGroupIdentity'
)

Write-Status -Message 'Starting Exchange on-prem mail-enabled security group inventory script.'
Ensure-ExchangeOnPremConnection

$scopeMode = 'Csv'
$resolvedServer = Get-TrimmedValue -Value $Server
$resolvedSearchBase = ''
$runWasTruncated = $false

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    $resolvedSearchBase = Get-TrimmedValue -Value $SearchBase
    Write-Status -Message "DiscoverAll enabled for Exchange on-prem mail-enabled security groups. SearchBase='$resolvedSearchBase'." -Level WARN
    $rows = @([PSCustomObject]@{
            SecurityGroupIdentity = '*'
        })
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $securityGroupIdentity = Get-TrimmedValue -Value $row.SecurityGroupIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($securityGroupIdentity)) {
            throw 'SecurityGroupIdentity is required. Use * to inventory all mail-enabled security groups.'
        }

        $effectiveSearchBase = if ($scopeMode -eq 'DiscoverAll') { $resolvedSearchBase } else { '' }
        $groups = @(Invoke-WithRetry -OperationName "Load mail-enabled security groups for $securityGroupIdentity" -ScriptBlock {
            Resolve-SecurityGroupsByScope -Identity $securityGroupIdentity -SearchBase $effectiveSearchBase -Server $resolvedServer
        })

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $groups.Count -gt $MaxObjects) {
            $groups = @($groups | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($groups.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $securityGroupIdentity -Action 'GetExchangeMailEnabledSecurityGroup' -Status 'NotFound' -Message 'No matching mail-enabled security groups were found.' -Data ([ordered]@{
                        SecurityGroupIdentity              = $securityGroupIdentity
                        Name                               = ''
                        Alias                              = ''
                        DisplayName                        = ''
                        PrimarySmtpAddress                 = ''
                        ManagedBy                          = ''
                        Notes                              = ''
                        RequireSenderAuthenticationEnabled = ''
                        HiddenFromAddressListsEnabled      = ''
                        ModerationEnabled                  = ''
                        ModeratedBy                        = ''
                        AcceptMessagesOnlyFrom             = ''
                        AcceptMessagesOnlyFromDLMembers    = ''
                        RejectMessagesFrom                 = ''
                        RejectMessagesFromDLMembers        = ''
                        BypassModerationFromSendersOrMembers = ''
                        SendModerationNotifications        = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($group in @($groups | Sort-Object -Property DisplayName, Identity)) {
            $identity = Get-TrimmedValue -Value $group.Identity
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeMailEnabledSecurityGroup' -Status 'Completed' -Message 'Mail-enabled security group exported.' -Data ([ordered]@{
                        SecurityGroupIdentity              = $identity
                        Name                               = Get-TrimmedValue -Value $group.Name
                        Alias                              = Get-TrimmedValue -Value $group.Alias
                        DisplayName                        = Get-TrimmedValue -Value $group.DisplayName
                        PrimarySmtpAddress                 = Get-TrimmedValue -Value $group.PrimarySmtpAddress
                        ManagedBy                          = Convert-MultiValueToString -Value $group.ManagedBy
                        Notes                              = Get-TrimmedValue -Value $group.Notes
                        RequireSenderAuthenticationEnabled = [string]$group.RequireSenderAuthenticationEnabled
                        HiddenFromAddressListsEnabled      = [string]$group.HiddenFromAddressListsEnabled
                        ModerationEnabled                  = [string]$group.ModerationEnabled
                        ModeratedBy                        = Convert-MultiValueToString -Value $group.ModeratedBy
                        AcceptMessagesOnlyFrom             = Convert-MultiValueToString -Value $group.AcceptMessagesOnlyFrom
                        AcceptMessagesOnlyFromDLMembers    = Convert-MultiValueToString -Value $group.AcceptMessagesOnlyFromDLMembers
                        RejectMessagesFrom                 = Convert-MultiValueToString -Value $group.RejectMessagesFrom
                        RejectMessagesFromDLMembers        = Convert-MultiValueToString -Value $group.RejectMessagesFromDLMembers
                        BypassModerationFromSendersOrMembers = Convert-MultiValueToString -Value $group.BypassModerationFromSendersOrMembers
                        SendModerationNotifications        = Get-TrimmedValue -Value $group.SendModerationNotifications
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($securityGroupIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $securityGroupIdentity -Action 'GetExchangeMailEnabledSecurityGroup' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    SecurityGroupIdentity              = $securityGroupIdentity
                    Name                               = ''
                    Alias                              = ''
                    DisplayName                        = ''
                    PrimarySmtpAddress                 = ''
                    ManagedBy                          = ''
                    Notes                              = ''
                    RequireSenderAuthenticationEnabled = ''
                    HiddenFromAddressListsEnabled      = ''
                    ModerationEnabled                  = ''
                    ModeratedBy                        = ''
                    AcceptMessagesOnlyFrom             = ''
                    AcceptMessagesOnlyFromDLMembers    = ''
                    RejectMessagesFrom                 = ''
                    RejectMessagesFromDLMembers        = ''
                    BypassModerationFromSendersOrMembers = ''
                    SendModerationNotifications        = ''
                })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeSearchBase' -NotePropertyValue $resolvedSearchBase -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeServer' -NotePropertyValue $resolvedServer -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeMaxObjects' -NotePropertyValue ($(if ($scopeMode -eq 'DiscoverAll') { [string]$MaxObjects } else { '' })) -Force
    Add-Member -InputObject $result -NotePropertyName 'ScopeWasTruncated' -NotePropertyValue ([string]$runWasTruncated) -Force
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem mail-enabled security group inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
