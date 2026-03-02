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

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'InventoryAndReport_OutputCsvPath') -ChildPath ("Results_IR3113-Get-ExchangeOnlineMailContacts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

$requiredHeaders = @(
    'MailContactIdentity'
)

Write-Status -Message 'Starting Exchange Online mail contact inventory script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailContactIdentity = ([string]$row.MailContactIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($mailContactIdentity)) {
            throw 'MailContactIdentity is required. Use * to inventory all mail contacts.'
        }

        $contacts = @()
        if ($mailContactIdentity -eq '*') {
            $contacts = @(Invoke-WithRetry -OperationName 'Load all mail contacts' -ScriptBlock {
                Get-MailContact -ResultSize Unlimited -ErrorAction Stop
            })
        }
        else {
            $contact = Invoke-WithRetry -OperationName "Lookup mail contact $mailContactIdentity" -ScriptBlock {
                Get-MailContact -Identity $mailContactIdentity -ErrorAction SilentlyContinue
            }
            if ($contact) {
                $contacts = @($contact)
            }
        }

        if ($contacts.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'GetExchangeMailContact' -Status 'NotFound' -Message 'No matching mail contacts were found.' -Data ([ordered]@{
                        MailContactIdentity            = $mailContactIdentity
                        Name                           = ''
                        Alias                          = ''
                        DisplayName                    = ''
                        ExternalEmailAddress           = ''
                        PrimarySmtpAddress             = ''
                        FirstName                      = ''
                        LastName                       = ''
                        HiddenFromAddressListsEnabled  = ''
                        WhenCreatedUTC                 = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($contact in @($contacts | Sort-Object -Property DisplayName, Identity)) {
            $identity = ([string]$contact.Identity).Trim()
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $identity -Action 'GetExchangeMailContact' -Status 'Completed' -Message 'Mail contact exported.' -Data ([ordered]@{
                        MailContactIdentity            = $identity
                        Name                           = ([string]$contact.Name).Trim()
                        Alias                          = ([string]$contact.Alias).Trim()
                        DisplayName                    = ([string]$contact.DisplayName).Trim()
                        ExternalEmailAddress           = ([string]$contact.ExternalEmailAddress).Trim()
                        PrimarySmtpAddress             = ([string]$contact.PrimarySmtpAddress).Trim()
                        FirstName                      = ([string]$contact.FirstName).Trim()
                        LastName                       = ([string]$contact.LastName).Trim()
                        HiddenFromAddressListsEnabled  = [string]$contact.HiddenFromAddressListsEnabled
                        WhenCreatedUTC                 = [string]$contact.WhenCreatedUTC
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailContactIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'GetExchangeMailContact' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    MailContactIdentity            = $mailContactIdentity
                    Name                           = ''
                    Alias                          = ''
                    DisplayName                    = ''
                    ExternalEmailAddress           = ''
                    PrimarySmtpAddress             = ''
                    FirstName                      = ''
                    LastName                       = ''
                    HiddenFromAddressListsEnabled  = ''
                    WhenCreatedUTC                 = ''
                })))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mail contact inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





