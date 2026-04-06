<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260406-000000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Audits SPF, DKIM, DMARC, and MX DNS records for Exchange Online accepted domains.

.DESCRIPTION
    Validates SPF, DKIM, DMARC, and MX DNS records for accepted domains in the tenant.
    In DiscoverAll mode, the accepted domain list is fetched from Exchange Online (single
    connection), the connection is then closed, and all DNS validation runs independently.
    In FromCsv mode, no Exchange Online connection is required — domain names are read
    from the input CSV and DNS validation is performed directly.
    One output row is written per domain per record type (MX, SPF, DMARC, DKIM).
    All results — including domains that could not be validated — are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include DomainName.
    Optionally include DkimSelectors (semicolon-separated) to test non-default DKIM selectors.
    See the companion .input.csv template for the full column list.

.PARAMETER DiscoverAll
    Fetch all accepted domains from Exchange Online and validate DNS records for each.
    The Exchange Online connection is closed before DNS queries begin.

.PARAMETER DkimSelectors
    DKIM selector names to check. Defaults to selector1 and selector2 (standard for Exchange Online).
    Values from this parameter are merged with any DkimSelectors column in the input CSV.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-EXOL-0290-Get-ExchangeOnlineEmailDnsAudit.ps1 -InputCsvPath .\D-EXOL-0290-Get-ExchangeOnlineEmailDnsAudit.input.csv

    Audit DNS records for domains listed in the input CSV.

.EXAMPLE
    .\D-EXOL-0290-Get-ExchangeOnlineEmailDnsAudit.ps1 -DiscoverAll

    Fetch all accepted domains from Exchange Online and audit their DNS records.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement (for DiscoverAll mode); Resolve-DnsName (DnsClient, Windows built-in)
    Required roles:   Exchange Administrator (DiscoverAll mode); no Exchange role required for FromCsv mode
    Limitations:      Resolve-DnsName is Windows-only. This script targets Windows hosts only.
                      DKIM check uses selector1 and selector2 by default (standard Exchange Online selectors).
                      DNS results reflect public DNS at query time; cached or propagating records may vary.

    CSV Fields:
    Column          Type      Required  Description
    --------------  --------  --------  -----------
    DomainName      String    Yes       Accepted domain name to audit (e.g., contoso.com)
    DkimSelectors   String    No        Semicolon-separated DKIM selector names (default: selector1;selector2)
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [Parameter(ParameterSetName = 'FromCsv')]
    [Parameter(ParameterSetName = 'DiscoverAll')]
    [string[]]$DkimSelectors = @('selector1', 'selector2'),

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-EXOL-0290-Get-ExchangeOnlineEmailDnsAudit_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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
        [Parameter(Mandatory)][int]$RowNumber,
        [Parameter(Mandatory)][string]$PrimaryKey,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$Message,
        [Parameter(Mandatory)][hashtable]$Data
    )

    $base    = New-ResultObject -RowNumber $RowNumber -PrimaryKey $PrimaryKey -Action $Action -Status $Status -Message $Message
    $ordered = [ordered]@{}

    foreach ($prop in $base.PSObject.Properties.Name) {
        $ordered[$prop] = $base.$prop
    }

    foreach ($key in $Data.Keys) {
        $ordered[$key] = $Data[$key]
    }

    return [PSCustomObject]$ordered
}

function Invoke-DnsQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Type
    )

    try {
        $records = Resolve-DnsName -Name $Name -Type $Type -ErrorAction Stop
        return @($records)
    }
    catch {
        return @()
    }
}

function Test-MxRecord {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Domain)

    $records = Invoke-DnsQuery -Name $Domain -Type 'MX'
    $mxRecords = @($records | Where-Object { $_.Type -eq 'MX' })

    if ($mxRecords.Count -eq 0) {
        return [ordered]@{
            Status           = 'Fail'
            CurrentValue     = ''
            RecommendedValue = 'At least one MX record is required.'
            Notes            = 'No MX records found.'
        }
    }

    $valueList = ($mxRecords | Sort-Object Preference | ForEach-Object { "$($_.Preference) $($_.NameExchange)" }) -join '; '
    $hasO365Mx = $mxRecords | Where-Object { [string]$_.NameExchange -imatch '\.mail\.protection\.outlook\.com\.?$' }

    return [ordered]@{
        Status           = if ($hasO365Mx) { 'Pass' } else { 'Warning' }
        CurrentValue     = $valueList
        RecommendedValue = "10 <tenant>.mail.protection.outlook.com"
        Notes            = if ($hasO365Mx) { 'Exchange Online MX record detected.' } else { 'MX record found but does not point to Exchange Online.' }
    }
}

function Test-SpfRecord {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Domain)

    $records = Invoke-DnsQuery -Name $Domain -Type 'TXT'
    $spfRecords = @($records | Where-Object { $_.Type -eq 'TXT' -and ([string]$_.Strings -imatch 'v=spf1' -or [string]($_.Strings -join '') -imatch 'v=spf1') })

    if ($spfRecords.Count -eq 0) {
        return [ordered]@{
            Status           = 'Fail'
            CurrentValue     = ''
            RecommendedValue = 'v=spf1 include:spf.protection.outlook.com -all'
            Notes            = 'No SPF TXT record found.'
        }
    }

    if ($spfRecords.Count -gt 1) {
        return [ordered]@{
            Status           = 'Fail'
            CurrentValue     = ($spfRecords | ForEach-Object { $_.Strings -join '' }) -join ' | '
            RecommendedValue = 'v=spf1 include:spf.protection.outlook.com -all'
            Notes            = 'Multiple SPF records detected — only one is permitted.'
        }
    }

    $spfValue  = $spfRecords[0].Strings -join ''
    $hasO365   = $spfValue -imatch 'include:spf\.protection\.outlook\.com'
    $hasHardFail = $spfValue -imatch '\-all$'

    $notes = @()
    if (-not $hasO365)    { $notes += 'Missing include:spf.protection.outlook.com.' }
    if (-not $hasHardFail) { $notes += 'Recommend using -all (hard fail) rather than ~all or ?all.' }

    return [ordered]@{
        Status           = if ($hasO365 -and $hasHardFail) { 'Pass' } elseif ($hasO365) { 'Warning' } else { 'Fail' }
        CurrentValue     = $spfValue
        RecommendedValue = 'v=spf1 include:spf.protection.outlook.com -all'
        Notes            = if ($notes.Count -gt 0) { $notes -join ' ' } else { 'SPF record looks correct for Exchange Online.' }
    }
}

function Test-DmarcRecord {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Domain)

    $dmarcName  = "_dmarc.$Domain"
    $records    = Invoke-DnsQuery -Name $dmarcName -Type 'TXT'
    $dmarcRecs  = @($records | Where-Object { $_.Type -eq 'TXT' -and ([string]($_.Strings -join '') -imatch 'v=DMARC1') })

    if ($dmarcRecs.Count -eq 0) {
        return [ordered]@{
            Status           = 'Fail'
            CurrentValue     = ''
            RecommendedValue = 'v=DMARC1; p=quarantine; rua=mailto:dmarc@<domain>'
            Notes            = "No DMARC TXT record found at $dmarcName."
        }
    }

    $dmarcValue = $dmarcRecs[0].Strings -join ''
    $policyMatch = [regex]::Match($dmarcValue, 'p=(\w+)')
    $policy     = if ($policyMatch.Success) { $policyMatch.Groups[1].Value.ToLowerInvariant() } else { '' }

    $status = switch ($policy) {
        'reject'     { 'Pass' }
        'quarantine' { 'Pass' }
        'none'       { 'Warning' }
        default      { 'Warning' }
    }

    return [ordered]@{
        Status           = $status
        CurrentValue     = $dmarcValue
        RecommendedValue = 'v=DMARC1; p=quarantine; rua=mailto:dmarc@<domain>'
        Notes            = "Policy: $policy. Recommend p=quarantine or p=reject for production."
    }
}

function Test-DkimRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Domain,
        [Parameter(Mandatory)][string]$Selector
    )

    $dkimName = "$Selector._domainkey.$Domain"
    $records  = Invoke-DnsQuery -Name $dkimName -Type 'CNAME'
    $cnameRec = @($records | Where-Object { $_.Type -eq 'CNAME' })

    if ($cnameRec.Count -gt 0) {
        $cnameTarget = [string]$cnameRec[0].NameHost
        $isO365Cname = $cnameTarget -imatch '_domainkey\.(outlook\.com|protection\.outlook\.com)' -or $cnameTarget -imatch '\._domainkey\.'

        return [ordered]@{
            Status           = if ($isO365Cname) { 'Pass' } else { 'Warning' }
            CurrentValue     = "CNAME -> $cnameTarget"
            RecommendedValue = "$Selector._domainkey.<tenant>.onmicrosoft.com"
            Notes            = if ($isO365Cname) { 'DKIM CNAME points to Exchange Online.' } else { 'DKIM CNAME found but does not appear to target Exchange Online.' }
        }
    }

    # Fall back to TXT record check (some environments publish TXT directly).
    $txtRecords = Invoke-DnsQuery -Name $dkimName -Type 'TXT'
    $dkimTxt    = @($txtRecords | Where-Object { $_.Type -eq 'TXT' -and ([string]($_.Strings -join '') -imatch 'v=DKIM1') })

    if ($dkimTxt.Count -gt 0) {
        return [ordered]@{
            Status           = 'Pass'
            CurrentValue     = $dkimTxt[0].Strings -join ''
            RecommendedValue = ''
            Notes            = 'DKIM TXT record found directly (not CNAME). Verify DKIM signing is active in Exchange Online.'
        }
    }

    return [ordered]@{
        Status           = 'Fail'
        CurrentValue     = ''
        RecommendedValue = "CNAME $Selector._domainkey.<domain> -> <selector>-<domain-sanitized>._domainkey.<tenant>.onmicrosoft.com"
        Notes            = "No DKIM CNAME or TXT record found at $dkimName."
    }
}

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'DomainName',
    'RecordType',
    'DkimSelector',
    'AuditStatus',
    'CurrentValue',
    'RecommendedValue',
    'Notes'
)

$requiredHeaders = @('DomainName')

Write-Status -Message 'Starting Exchange Online email DNS audit script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')

$scopeMode = 'Csv'

if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. Fetching accepted domains from Exchange Online.' -Level WARN
    Ensure-ExchangeConnection

    $acceptedDomains = Invoke-WithRetry -OperationName 'Get accepted domains' -ScriptBlock {
        Get-AcceptedDomain -ErrorAction Stop | Select-Object -ExpandProperty DomainName
    }

    Write-Status -Message "Fetched $($acceptedDomains.Count) accepted domains. Disconnecting from Exchange Online."
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

    $rows = @($acceptedDomains | ForEach-Object {
        [PSCustomObject]@{ DomainName = [string]$_; DkimSelectors = '' }
    })
} else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $domainName = Get-TrimmedValue -Value $row.DomainName
    $primaryKey = $domainName

    if ([string]::IsNullOrWhiteSpace($domainName)) {
        Write-Status -Message "Row $rowNumber skipped: DomainName is empty." -Level WARN
        $rowNumber++
        continue
    }

    # Resolve DKIM selectors: merge parameter defaults with optional CSV column.
    $effectiveSelectors = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($sel in $DkimSelectors) {
        $s = Get-TrimmedValue -Value $sel
        if (-not [string]::IsNullOrWhiteSpace($s)) { [void]$effectiveSelectors.Add($s) }
    }

    if ($row.PSObject.Properties['DkimSelectors']) {
        $csvSelectors = ConvertTo-Array -Value $row.DkimSelectors
        foreach ($sel in $csvSelectors) {
            $s = Get-TrimmedValue -Value $sel
            if (-not [string]::IsNullOrWhiteSpace($s)) { [void]$effectiveSelectors.Add($s) }
        }
    }

    # Record type checks: MX, SPF, DMARC, DKIM (per selector).
    $checksToRun = [System.Collections.Generic.List[object]]::new()
    $checksToRun.Add([ordered]@{ RecordType = 'MX';    DkimSelector = '' })
    $checksToRun.Add([ordered]@{ RecordType = 'SPF';   DkimSelector = '' })
    $checksToRun.Add([ordered]@{ RecordType = 'DMARC'; DkimSelector = '' })
    foreach ($sel in ($effectiveSelectors | Sort-Object)) {
        $checksToRun.Add([ordered]@{ RecordType = 'DKIM'; DkimSelector = $sel })
    }

    foreach ($check in $checksToRun) {
        $recordType   = [string]$check.RecordType
        $dkimSelector = [string]$check.DkimSelector
        $checkKey     = if ($dkimSelector) { "${domainName}:${recordType}:${dkimSelector}" } else { "${domainName}:${recordType}" }

        try {
            $auditResult = switch ($recordType) {
                'MX'    { Test-MxRecord    -Domain $domainName }
                'SPF'   { Test-SpfRecord   -Domain $domainName }
                'DMARC' { Test-DmarcRecord -Domain $domainName }
                'DKIM'  { Test-DkimRecord  -Domain $domainName -Selector $dkimSelector }
            }

            $overallStatus = if ($auditResult.Status -in @('Pass', 'Warning')) { 'Completed' } else { 'Completed' }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $checkKey -Action 'AuditEmailDnsRecord' -Status $overallStatus -Message $auditResult.Notes -Data ([ordered]@{
                DomainName       = $domainName
                RecordType       = $recordType
                DkimSelector     = $dkimSelector
                AuditStatus      = $auditResult.Status
                CurrentValue     = $auditResult.CurrentValue
                RecommendedValue = $auditResult.RecommendedValue
                Notes            = $auditResult.Notes
            })))
        }
        catch {
            Write-Status -Message "Row $rowNumber ($checkKey) DNS check failed: $($_.Exception.Message)" -Level ERROR
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $checkKey -Action 'AuditEmailDnsRecord' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                DomainName       = $domainName
                RecordType       = $recordType
                DkimSelector     = $dkimSelector
                AuditStatus      = 'Error'
                CurrentValue     = ''
                RecommendedValue = ''
                Notes            = $_.Exception.Message
            })))
        }
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
Write-Status -Message 'Exchange Online email DNS audit script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
