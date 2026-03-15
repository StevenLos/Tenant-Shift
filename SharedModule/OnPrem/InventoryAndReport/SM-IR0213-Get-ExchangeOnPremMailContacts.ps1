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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_SM-IR0213-Get-ExchangeOnPremMailContacts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Resolve-MailContactsByScope {
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

        Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-MailContact' -ParameterName 'OrganizationalUnit' -Value $SearchBase
        Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-MailContact' -ParameterName 'DomainController' -Value $Server

        return @(Get-MailContact @params)
    }

    $params = @{
        Identity    = $Identity
        ErrorAction = 'SilentlyContinue'
    }

    Add-SupportedParameter -ParameterHashtable $params -CommandName 'Get-MailContact' -ParameterName 'DomainController' -Value $Server

    $contact = Get-MailContact @params
    if ($contact) {
        return @($contact)
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

$requiredHeaders = @(
    'MailContactIdentity'
)

Write-Status -Message 'Starting Exchange on-prem mail contact inventory script.'
Ensure-ExchangeOnPremConnection

$scopeMode = 'Csv'
$resolvedServer = Get-TrimmedValue -Value $Server
$resolvedSearchBase = ''
$runWasTruncated = $false

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    $resolvedSearchBase = Get-TrimmedValue -Value $SearchBase
    Write-Status -Message "DiscoverAll enabled for Exchange on-prem mail contacts. SearchBase='$resolvedSearchBase'." -Level WARN
    $rows = @([PSCustomObject]@{
            MailContactIdentity = '*'
        })
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailContactIdentity = Get-TrimmedValue -Value $row.MailContactIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($mailContactIdentity)) {
            throw 'MailContactIdentity is required. Use * to inventory all mail contacts.'
        }

        $effectiveSearchBase = if ($scopeMode -eq 'DiscoverAll') { $resolvedSearchBase } else { '' }
        $contacts = @(Invoke-WithRetry -OperationName "Load mail contacts for $mailContactIdentity" -ScriptBlock {
            Resolve-MailContactsByScope -Identity $mailContactIdentity -SearchBase $effectiveSearchBase -Server $resolvedServer
        })

        if ($scopeMode -eq 'DiscoverAll' -and $MaxObjects -gt 0 -and $contacts.Count -gt $MaxObjects) {
            $contacts = @($contacts | Select-Object -First $MaxObjects)
            $runWasTruncated = $true
        }

        if ($contacts.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'GetExchangeMailContact' -Status 'NotFound' -Message 'No matching mail contacts were found.' -Data ([ordered]@{
                        MailContactIdentity           = $mailContactIdentity
                        Name                          = ''
                        Alias                         = ''
                        DisplayName                   = ''
                        ExternalEmailAddress          = ''
                        PrimarySmtpAddress            = ''
                        FirstName                     = ''
                        LastName                      = ''
                        HiddenFromAddressListsEnabled = ''
                        WhenCreatedUTC                = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($contact in @($contacts | Sort-Object -Property DisplayName, Identity)) {
            $identity = Get-TrimmedValue -Value $contact.Identity
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeMailContact' -Status 'Completed' -Message 'Mail contact exported.' -Data ([ordered]@{
                        MailContactIdentity           = $identity
                        Name                          = Get-TrimmedValue -Value $contact.Name
                        Alias                         = Get-TrimmedValue -Value $contact.Alias
                        DisplayName                   = Get-TrimmedValue -Value $contact.DisplayName
                        ExternalEmailAddress          = Get-TrimmedValue -Value $contact.ExternalEmailAddress
                        PrimarySmtpAddress            = Get-TrimmedValue -Value $contact.PrimarySmtpAddress
                        FirstName                     = Get-TrimmedValue -Value $contact.FirstName
                        LastName                      = Get-TrimmedValue -Value $contact.LastName
                        HiddenFromAddressListsEnabled = [string]$contact.HiddenFromAddressListsEnabled
                        WhenCreatedUTC                = Get-TrimmedValue -Value $contact.WhenCreatedUTC
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailContactIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'GetExchangeMailContact' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    MailContactIdentity           = $mailContactIdentity
                    Name                          = ''
                    Alias                         = ''
                    DisplayName                   = ''
                    ExternalEmailAddress          = ''
                    PrimarySmtpAddress            = ''
                    FirstName                     = ''
                    LastName                      = ''
                    HiddenFromAddressListsEnabled = ''
                    WhenCreatedUTC                = ''
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
Write-Status -Message 'Exchange on-prem mail contact inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
