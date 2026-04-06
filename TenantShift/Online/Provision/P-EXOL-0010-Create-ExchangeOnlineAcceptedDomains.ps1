<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-154500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement
Microsoft.Graph.Authentication
Microsoft.Graph.Identity.DirectoryManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Provisions ExchangeOnlineAcceptedDomains in Microsoft 365.

.DESCRIPTION
    Creates ExchangeOnlineAcceptedDomains in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P3124-Create-ExchangeOnlineAcceptedDomains.ps1 -InputCsvPath .\3124.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P3124-Create-ExchangeOnlineAcceptedDomains.ps1 -InputCsvPath .\3124.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement
    Required roles:   Exchange Administrator
    Limitations:      None known.

    CSV Fields:
    Column                  Type      Required  Description
    ----------------------  ----      --------  -----------
    DomainName              String    Yes       <fill in description>
    AcceptedDomainName      String    Yes       <fill in description>
    DomainType              String    Yes       <fill in description>
    SetAsDefault            String    Yes       <fill in description>
    MatchSubDomains         String    Yes       <fill in description>
    AutoCreateTenantDomain  String    Yes       <fill in description>
    Notes                   String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P3124-Create-ExchangeOnlineAcceptedDomains_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

function Get-NullableBool {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    $text = Get-TrimmedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return (ConvertTo-Bool -Value $text)
}

function Set-DefaultAcceptedDomainFlag {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Target,

        [Parameter(Mandatory)]
        [System.Management.Automation.CommandInfo]$Command,

        [Parameter(Mandatory)]
        [bool]$Value
    )

    if ($Command.Parameters.ContainsKey('Default')) {
        $Target['Default'] = $Value
        return $true
    }

    if ($Value -and $Command.Parameters.ContainsKey('MakeDefault')) {
        $Target['MakeDefault'] = $true
        return $true
    }

    return $false
}

$requiredHeaders = @(
    'DomainName',
    'AcceptedDomainName',
    'DomainType',
    'SetAsDefault',
    'MatchSubDomains',
    'AutoCreateTenantDomain',
    'Notes'
)

$validDomainTypes = @('Authoritative', 'InternalRelay', 'ExternalRelay')

Write-Status -Message 'Starting Exchange Online accepted-domain creation script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Identity.DirectoryManagement')
Ensure-ExchangeConnection
Ensure-GraphConnection -RequiredScopes @('Domain.ReadWrite.All', 'Directory.Read.All')

$newAcceptedDomainCommand = Get-Command -Name New-AcceptedDomain -ErrorAction Stop
$setAcceptedDomainCommand = Get-Command -Name Set-AcceptedDomain -ErrorAction Stop
$supportsMatchSubDomainsOnCreate = $newAcceptedDomainCommand.Parameters.ContainsKey('MatchSubDomains')
$supportsMatchSubDomainsOnSet = $setAcceptedDomainCommand.Parameters.ContainsKey('MatchSubDomains')

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

        $acceptedDomainName = Get-TrimmedValue -Value $row.AcceptedDomainName
        if ([string]::IsNullOrWhiteSpace($acceptedDomainName)) {
            $acceptedDomainName = $domainName
        }

        $domainTypeRaw = Get-TrimmedValue -Value $row.DomainType
        $domainType = if ([string]::IsNullOrWhiteSpace($domainTypeRaw)) { 'Authoritative' } else { $domainTypeRaw }
        if ($domainType -notin $validDomainTypes) {
            throw "DomainType '$domainType' is invalid. Use Authoritative, InternalRelay, or ExternalRelay."
        }

        $setAsDefault = ConvertTo-Bool -Value (Get-TrimmedValue -Value $row.SetAsDefault)
        $autoCreateTenantDomain = ConvertTo-Bool -Value (Get-TrimmedValue -Value $row.AutoCreateTenantDomain)
        $matchSubDomains = Get-NullableBool -Value $row.MatchSubDomains

        $messages = [System.Collections.Generic.List[string]]::new()

        $tenantDomain = Invoke-WithRetry -OperationName "Lookup tenant domain $domainName" -ScriptBlock {
            Get-MgDomain -DomainId $domainName -ErrorAction SilentlyContinue
        }

        if (-not $tenantDomain) {
            if (-not $autoCreateTenantDomain) {
                throw "Tenant domain '$domainName' was not found. Set AutoCreateTenantDomain to TRUE or create the Entra domain first."
            }

            if ($PSCmdlet.ShouldProcess($domainName, 'Create Entra tenant domain')) {
                Invoke-WithRetry -OperationName "Create tenant domain $domainName" -ScriptBlock {
                    New-MgDomain -Id $domainName -ErrorAction Stop | Out-Null
                }

                $messages.Add('Tenant domain created in Entra.')
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'CreateAcceptedDomain' -Status 'WhatIf' -Message 'Tenant-domain creation skipped due to WhatIf.'))
                $rowNumber++
                continue
            }
        }

        $existingAcceptedDomain = Invoke-WithRetry -OperationName "Lookup accepted domain $domainName" -ScriptBlock {
            Get-AcceptedDomain -Identity $domainName -ErrorAction SilentlyContinue
        }

        if (-not $existingAcceptedDomain) {
            $createParams = @{
                Name       = $acceptedDomainName
                DomainName = $domainName
                DomainType = $domainType
            }

            if ($setAsDefault) {
                $defaultApplied = Set-DefaultAcceptedDomainFlag -Target $createParams -Command $newAcceptedDomainCommand -Value $true
                if (-not $defaultApplied) {
                    $messages.Add('Default accepted-domain assignment was requested but is unsupported in this session.')
                }
            }

            if ($null -ne $matchSubDomains) {
                if ($supportsMatchSubDomainsOnCreate) {
                    $createParams['MatchSubDomains'] = [bool]$matchSubDomains
                }
                else {
                    $messages.Add('MatchSubDomains was provided but New-AcceptedDomain does not support it in this session.')
                }
            }

            if ($PSCmdlet.ShouldProcess($domainName, 'Create Exchange Online accepted domain')) {
                Invoke-WithRetry -OperationName "Create accepted domain $domainName" -ScriptBlock {
                    New-AcceptedDomain @createParams -ErrorAction Stop | Out-Null
                }

                $message = 'Accepted domain created successfully.'
                if ($messages.Count -gt 0) {
                    $message = "$message $($messages -join ' ')"
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'CreateAcceptedDomain' -Status 'Created' -Message $message))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'CreateAcceptedDomain' -Status 'WhatIf' -Message 'Accepted-domain creation skipped due to WhatIf.'))
            }

            $rowNumber++
            continue
        }

        $setParams = @{
            Identity = $existingAcceptedDomain.Identity
        }

        $currentDomainType = Get-TrimmedValue -Value $existingAcceptedDomain.DomainType
        if (-not [string]::IsNullOrWhiteSpace($currentDomainType) -and $currentDomainType -ne $domainType) {
            $setParams['DomainType'] = $domainType
        }

        if ($setAsDefault) {
            $currentIsDefault = [bool]$existingAcceptedDomain.Default
            if (-not $currentIsDefault) {
                $defaultApplied = Set-DefaultAcceptedDomainFlag -Target $setParams -Command $setAcceptedDomainCommand -Value $true
                if (-not $defaultApplied) {
                    $messages.Add('Default accepted-domain assignment was requested but is unsupported in this session.')
                }
            }
        }

        if ($null -ne $matchSubDomains) {
            if ($supportsMatchSubDomainsOnSet) {
                $currentMatchSubDomains = [bool]$existingAcceptedDomain.MatchSubDomains
                if ($currentMatchSubDomains -ne [bool]$matchSubDomains) {
                    $setParams['MatchSubDomains'] = [bool]$matchSubDomains
                }
            }
            else {
                $messages.Add('MatchSubDomains was provided but Set-AcceptedDomain does not support it in this session.')
            }
        }

        if ($setParams.Count -eq 1) {
            $message = 'Accepted domain already exists with requested settings.'
            if ($messages.Count -gt 0) {
                $message = "$message $($messages -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'CreateAcceptedDomain' -Status 'Skipped' -Message $message))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($domainName, 'Update Exchange Online accepted domain settings')) {
            Invoke-WithRetry -OperationName "Update accepted domain $domainName" -ScriptBlock {
                Set-AcceptedDomain @setParams -ErrorAction Stop
            }

            $message = 'Accepted domain updated successfully.'
            if ($messages.Count -gt 0) {
                $message = "$message $($messages -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'CreateAcceptedDomain' -Status 'Updated' -Message $message))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'CreateAcceptedDomain' -Status 'WhatIf' -Message 'Accepted-domain update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($domainName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $domainName -Action 'CreateAcceptedDomain' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online accepted-domain creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
