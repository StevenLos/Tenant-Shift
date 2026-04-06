<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-160500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement
Microsoft.Graph.Authentication
Microsoft.Graph.Identity.DirectoryManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Modifies ExchangeOnlineAcceptedDomains in Microsoft 365.

.DESCRIPTION
    Updates ExchangeOnlineAcceptedDomains in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M3130-Verify-ExchangeOnlineAcceptedDomains.ps1 -InputCsvPath .\3130.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M3130-Verify-ExchangeOnlineAcceptedDomains.ps1 -InputCsvPath .\3130.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement
    Required roles:   Exchange Administrator
    Limitations:      None known.

    CSV Fields:
    Column                 Type      Required  Description
    ---------------------  ----      --------  -----------
    DomainName             String    Yes       <fill in description>
    AttemptVerification    String    Yes       <fill in description>
    RequireAcceptedDomain  String    Yes       <fill in description>
    RequireTenantDomain    String    Yes       <fill in description>
    Notes                  String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3130-Verify-ExchangeOnlineAcceptedDomains_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

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
    'AttemptVerification',
    'RequireAcceptedDomain',
    'RequireTenantDomain',
    'Notes'
)

Write-Status -Message 'Starting Exchange Online accepted-domain verification script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Identity.DirectoryManagement')
Ensure-ExchangeConnection
Ensure-GraphConnection -RequiredScopes @('Domain.ReadWrite.All', 'Directory.Read.All')

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

        $attemptVerification = Get-BoolWithDefault -Value $row.AttemptVerification -Default $true
        $requireAcceptedDomain = Get-BoolWithDefault -Value $row.RequireAcceptedDomain -Default $true
        $requireTenantDomain = Get-BoolWithDefault -Value $row.RequireTenantDomain -Default $true

        $acceptedDomain = Invoke-WithRetry -OperationName "Lookup accepted domain $domainName" -ScriptBlock {
            Get-AcceptedDomain -Identity $domainName -ErrorAction SilentlyContinue
        }

        $tenantDomain = Invoke-WithRetry -OperationName "Lookup tenant domain $domainName" -ScriptBlock {
            Get-MgDomain -DomainId $domainName -ErrorAction SilentlyContinue
        }

        $messages = [System.Collections.Generic.List[string]]::new()

        if (-not $acceptedDomain) {
            $messages.Add('Accepted domain not found in Exchange Online.')
            if ($requireAcceptedDomain) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'VerifyAcceptedDomain' -Status 'NotFound' -Message ($messages -join ' ')))
                $rowNumber++
                continue
            }
        }
        else {
            $messages.Add('Accepted domain located in Exchange Online.')
        }

        if (-not $tenantDomain) {
            $messages.Add('Tenant domain not found in Entra.')
            if ($requireTenantDomain) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'VerifyAcceptedDomain' -Status 'NotFound' -Message ($messages -join ' ')))
                $rowNumber++
                continue
            }
        }

        if ($tenantDomain -and -not [bool]$tenantDomain.IsVerified -and $attemptVerification) {
            if ($PSCmdlet.ShouldProcess($domainName, 'Confirm Entra domain verification')) {
                Invoke-WithRetry -OperationName "Confirm domain $domainName" -ScriptBlock {
                    Confirm-MgDomain -DomainId $domainName -ErrorAction Stop | Out-Null
                }

                $messages.Add('Domain verification was requested.')
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'VerifyAcceptedDomain' -Status 'WhatIf' -Message 'Domain verification skipped due to WhatIf.'))
                $rowNumber++
                continue
            }

            $tenantDomain = Invoke-WithRetry -OperationName "Refresh tenant domain $domainName" -ScriptBlock {
                Get-MgDomain -DomainId $domainName -ErrorAction SilentlyContinue
            }
        }

        if ($tenantDomain -and [bool]$tenantDomain.IsVerified) {
            $messages.Add('Tenant domain is verified.')
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'VerifyAcceptedDomain' -Status 'Verified' -Message ($messages -join ' ')))
        }
        elseif ($tenantDomain) {
            $messages.Add('Tenant domain remains unverified. Publish required DNS records and retry.')
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'VerifyAcceptedDomain' -Status 'PendingDns' -Message ($messages -join ' ')))
        }
        else {
            $messages.Add('Verification could not be evaluated because tenant domain details were unavailable.')
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'VerifyAcceptedDomain' -Status 'Skipped' -Message ($messages -join ' ')))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($domainName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'VerifyAcceptedDomain' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online accepted-domain verification script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
