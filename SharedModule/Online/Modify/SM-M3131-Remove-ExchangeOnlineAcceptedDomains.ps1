<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-161500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement
Microsoft.Graph.Authentication
Microsoft.Graph.Identity.DirectoryManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3131-Remove-ExchangeOnlineAcceptedDomains_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-TrimmedValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    return ([string]$Value).Trim()
}

function Get-NormalizedDomainName {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value,

        [Parameter(Mandatory)]
        [string]$FieldName
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    $normalized = $text.Trim('.').ToLowerInvariant()
    if ($normalized -notmatch '^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?)+$') {
        throw "$FieldName value '$text' is not a valid domain name."
    }

    return $normalized
}

function Get-BoolWithDefault {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value,

        [Parameter(Mandatory)]
        [bool]$Default
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $Default
    }

    return (ConvertTo-Bool -Value $text)
}

$requiredHeaders = @(
    'DomainName',
    'RemoveAcceptedDomain',
    'RemoveTenantDomain',
    'ForceRemoval',
    'AllowOnMicrosoftDomainRemoval',
    'Notes'
)

Write-Status -Message 'Starting Exchange Online accepted-domain removal script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Identity.DirectoryManagement')
Ensure-ExchangeConnection
Ensure-GraphConnection -RequiredScopes @('Domain.ReadWrite.All', 'Directory.Read.All')

$removeAcceptedDomainCommand = Get-Command -Name Remove-AcceptedDomain -ErrorAction Stop
$supportsExchangeForce = $removeAcceptedDomainCommand.Parameters.ContainsKey('Force')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $domainName = ''

    try {
        $domainName = Get-NormalizedDomainName -Value $row.DomainName -FieldName 'DomainName'
        if ([string]::IsNullOrWhiteSpace($domainName)) {
            throw 'DomainName is required.'
        }

        $removeAcceptedDomain = Get-BoolWithDefault -Value $row.RemoveAcceptedDomain -Default $true
        $removeTenantDomain = Get-BoolWithDefault -Value $row.RemoveTenantDomain -Default $false
        $forceRemoval = Get-BoolWithDefault -Value $row.ForceRemoval -Default $false
        $allowOnMicrosoftDomainRemoval = Get-BoolWithDefault -Value $row.AllowOnMicrosoftDomainRemoval -Default $false

        if (-not $removeAcceptedDomain -and -not $removeTenantDomain) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'RemoveAcceptedDomain' -Status 'Skipped' -Message 'No removal action requested.'))
            $rowNumber++
            continue
        }

        if ($domainName.EndsWith('.onmicrosoft.com', [System.StringComparison]::OrdinalIgnoreCase) -and -not $allowOnMicrosoftDomainRemoval) {
            throw 'Refusing to remove *.onmicrosoft.com domain. Set AllowOnMicrosoftDomainRemoval to TRUE to override.'
        }

        $messages = [System.Collections.Generic.List[string]]::new()
        $removed = $false
        $whatIfTriggered = $false

        if ($removeAcceptedDomain) {
            $acceptedDomain = Invoke-WithRetry -OperationName "Lookup accepted domain $domainName" -ScriptBlock {
                Get-AcceptedDomain -Identity $domainName -ErrorAction SilentlyContinue
            }

            if (-not $acceptedDomain) {
                $messages.Add('Accepted domain not found in Exchange Online.')
            }
            else {
                if ($PSCmdlet.ShouldProcess($domainName, 'Remove Exchange Online accepted domain')) {
                    $removeParams = @{
                        Identity = $acceptedDomain.Identity
                        Confirm  = $false
                    }

                    if ($forceRemoval -and $supportsExchangeForce) {
                        $removeParams['Force'] = $true
                    }

                    Invoke-WithRetry -OperationName "Remove accepted domain $domainName" -ScriptBlock {
                        Remove-AcceptedDomain @removeParams -ErrorAction Stop
                    }

                    $removed = $true
                    $messages.Add('Accepted domain removed from Exchange Online.')
                }
                else {
                    $whatIfTriggered = $true
                    $messages.Add('Accepted-domain removal skipped due to WhatIf.')
                }
            }
        }

        if ($removeTenantDomain) {
            $tenantDomain = Invoke-WithRetry -OperationName "Lookup tenant domain $domainName" -ScriptBlock {
                Get-MgDomain -DomainId $domainName -ErrorAction SilentlyContinue
            }

            if (-not $tenantDomain) {
                $messages.Add('Tenant domain not found in Entra.')
            }
            else {
                if ($PSCmdlet.ShouldProcess($domainName, 'Remove Entra tenant domain')) {
                    Invoke-WithRetry -OperationName "Remove tenant domain $domainName" -ScriptBlock {
                        Remove-MgDomain -DomainId $domainName -Confirm:$false -ErrorAction Stop
                    }

                    $removed = $true
                    $messages.Add('Tenant domain removed from Entra.')
                }
                else {
                    $whatIfTriggered = $true
                    $messages.Add('Tenant-domain removal skipped due to WhatIf.')
                }
            }
        }

        if ($whatIfTriggered -and -not $removed) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'RemoveAcceptedDomain' -Status 'WhatIf' -Message ($messages -join ' ')))
            $rowNumber++
            continue
        }

        if ($removed) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'RemoveAcceptedDomain' -Status 'Updated' -Message ($messages -join ' ')))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'RemoveAcceptedDomain' -Status 'Skipped' -Message ($messages -join ' ')))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($domainName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'RemoveAcceptedDomain' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online accepted-domain removal script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
