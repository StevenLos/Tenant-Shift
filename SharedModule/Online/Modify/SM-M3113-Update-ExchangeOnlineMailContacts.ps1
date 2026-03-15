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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3113-Update-ExchangeOnlineMailContacts_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'MailContactIdentity',
    'Name',
    'DisplayName',
    'FirstName',
    'LastName',
    'ExternalEmailAddress',
    'PrimarySmtpAddress',
    'HiddenFromAddressListsEnabled'
)

Write-Status -Message 'Starting Exchange Online mail contact update script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailContactIdentity = ([string]$row.MailContactIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($mailContactIdentity)) {
            throw 'MailContactIdentity is required.'
        }

        $contact = Invoke-WithRetry -OperationName "Lookup mail contact $mailContactIdentity" -ScriptBlock {
            Get-MailContact -Identity $mailContactIdentity -ErrorAction SilentlyContinue
        }
        if (-not $contact) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'UpdateMailContact' -Status 'NotFound' -Message 'Mail contact not found.'))
            $rowNumber++
            continue
        }

        $setParams = @{
            Identity = $contact.Identity
        }

        $name = ([string]$row.Name).Trim()
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $setParams.Name = $name
        }

        $displayName = ([string]$row.DisplayName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $setParams.DisplayName = $displayName
        }

        $firstName = ([string]$row.FirstName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($firstName)) {
            $setParams.FirstName = $firstName
        }

        $lastName = ([string]$row.LastName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($lastName)) {
            $setParams.LastName = $lastName
        }

        $externalEmailAddress = ([string]$row.ExternalEmailAddress).Trim()
        if (-not [string]::IsNullOrWhiteSpace($externalEmailAddress)) {
            $setParams.ExternalEmailAddress = $externalEmailAddress
        }

        $primarySmtpAddress = ([string]$row.PrimarySmtpAddress).Trim()
        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $setParams.PrimarySmtpAddress = $primarySmtpAddress
        }

        $hiddenRaw = ([string]$row.HiddenFromAddressListsEnabled).Trim()
        if (-not [string]::IsNullOrWhiteSpace($hiddenRaw)) {
            $setParams.HiddenFromAddressListsEnabled = ConvertTo-Bool -Value $hiddenRaw
        }

        if ($setParams.Count -eq 1) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'UpdateMailContact' -Status 'Skipped' -Message 'No updates specified.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($mailContactIdentity, 'Update Exchange Online mail contact')) {
            Invoke-WithRetry -OperationName "Update mail contact $mailContactIdentity" -ScriptBlock {
                Set-MailContact @setParams -ErrorAction Stop
            }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'UpdateMailContact' -Status 'Updated' -Message 'Mail contact updated successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'UpdateMailContact' -Status 'WhatIf' -Message 'Update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailContactIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailContactIdentity -Action 'UpdateMailContact' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mail contact update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}





