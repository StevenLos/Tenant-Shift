<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-175500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement
Microsoft.Graph.Authentication
Microsoft.Graph.Identity.DirectoryManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineDomainVerificationRecords and exports results to CSV.

.DESCRIPTION
    Gets ExchangeOnlineDomainVerificationRecords from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-D3124-Get-ExchangeOnlineDomainVerificationRecords.ps1 -InputCsvPath .\3124.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D3124-Get-ExchangeOnlineDomainVerificationRecords.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement
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

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-EXOL-0010-Get-ExchangeOnlineDomainVerificationRecords_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
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

    if ($InputObject.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $additional = $InputObject.AdditionalProperties
        if ($additional) {
            try {
                if ($additional.ContainsKey($PropertyName)) {
                    return $additional[$PropertyName]
                }
            }
            catch {
            }
        }
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

    return Get-TrimmedValue -Value (Get-ObjectPropertyValue -InputObject $InputObject -PropertyName $PropertyName)
}

function New-DomainRecordData {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$InputDomainName = '',

        [AllowNull()]
        [string]$AcceptedDomainIdentity = '',

        [AllowNull()]
        [string]$AcceptedDomainName = '',

        [AllowNull()]
        [string]$DomainName = '',

        [AllowNull()]
        [string]$DomainType = '',

        [AllowNull()]
        [string]$AcceptedDomainIsDefault = '',

        [AllowNull()]
        [string]$AcceptedDomainMatchSubDomains = '',

        [AllowNull()]
        [string]$TenantDomainIsVerified = '',

        [AllowNull()]
        [string]$TenantDomainIsAdminManaged = '',

        [AllowNull()]
        [string]$TenantDomainAuthenticationType = '',

        [AllowNull()]
        [string]$TenantDomainSupportedServices = '',

        [AllowNull()]
        [string]$VerificationRecordType = '',

        [AllowNull()]
        [string]$VerificationLabel = '',

        [AllowNull()]
        [string]$VerificationText = '',

        [AllowNull()]
        [string]$VerificationMailExchange = '',

        [AllowNull()]
        [string]$VerificationCanonicalName = '',

        [AllowNull()]
        [string]$VerificationTtl = '',

        [AllowNull()]
        [string]$VerificationIsOptional = '',

        [AllowNull()]
        [string]$VerificationSupportedService = ''
    )

    return [ordered]@{
        InputDomainName                = $InputDomainName
        AcceptedDomainIdentity         = $AcceptedDomainIdentity
        AcceptedDomainName             = $AcceptedDomainName
        DomainName                     = $DomainName
        DomainType                     = $DomainType
        AcceptedDomainIsDefault        = $AcceptedDomainIsDefault
        AcceptedDomainMatchSubDomains  = $AcceptedDomainMatchSubDomains
        TenantDomainIsVerified         = $TenantDomainIsVerified
        TenantDomainIsAdminManaged     = $TenantDomainIsAdminManaged
        TenantDomainAuthenticationType = $TenantDomainAuthenticationType
        TenantDomainSupportedServices  = $TenantDomainSupportedServices
        VerificationRecordType         = $VerificationRecordType
        VerificationLabel              = $VerificationLabel
        VerificationText               = $VerificationText
        VerificationMailExchange       = $VerificationMailExchange
        VerificationCanonicalName      = $VerificationCanonicalName
        VerificationTtl                = $VerificationTtl
        VerificationIsOptional         = $VerificationIsOptional
        VerificationSupportedService   = $VerificationSupportedService
    }
}

$requiredHeaders = @(
    'DomainName'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'InputDomainName',
    'AcceptedDomainIdentity',
    'AcceptedDomainName',
    'DomainName',
    'DomainType',
    'AcceptedDomainIsDefault',
    'AcceptedDomainMatchSubDomains',
    'TenantDomainIsVerified',
    'TenantDomainIsAdminManaged',
    'TenantDomainAuthenticationType',
    'TenantDomainSupportedServices',
    'VerificationRecordType',
    'VerificationLabel',
    'VerificationText',
    'VerificationMailExchange',
    'VerificationCanonicalName',
    'VerificationTtl',
    'VerificationIsOptional',
    'VerificationSupportedService'
)

Write-Status -Message 'Starting Exchange Online domain verification-record inventory script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Identity.DirectoryManagement')
Ensure-ExchangeConnection
Ensure-GraphConnection -RequiredScopes @('Domain.Read.All', 'Directory.Read.All')

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
    $inputDomainName = Get-TrimmedValue -Value $row.DomainName

    try {
        if ([string]::IsNullOrWhiteSpace($inputDomainName)) {
            throw 'DomainName is required. Use * to inventory all accepted domains.'
        }

        $acceptedDomains = @()
        if ($inputDomainName -eq '*') {
            $acceptedDomains = @(Invoke-WithRetry -OperationName 'Load all accepted domains' -ScriptBlock {
                Get-AcceptedDomain -ErrorAction Stop
            })
        }
        else {
            $acceptedDomain = Invoke-WithRetry -OperationName "Lookup accepted domain $inputDomainName" -ScriptBlock {
                Get-AcceptedDomain -Identity $inputDomainName -ErrorAction SilentlyContinue
            }

            if ($acceptedDomain) {
                $acceptedDomains = @($acceptedDomain)
            }
        }

        if ($acceptedDomains.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $inputDomainName -Action 'GetAcceptedDomainVerificationRecords' -Status 'NotFound' -Message 'No matching accepted domains were found.' -Data (New-DomainRecordData -InputDomainName $inputDomainName)))
            $rowNumber++
            continue
        }

        foreach ($acceptedDomain in @($acceptedDomains | Sort-Object -Property DomainName, Name)) {
            $acceptedIdentity = Get-TrimmedValue -Value $acceptedDomain.Identity
            $domainName = Get-TrimmedValue -Value $acceptedDomain.DomainName
            $acceptedName = Get-TrimmedValue -Value $acceptedDomain.Name
            $domainType = Get-TrimmedValue -Value $acceptedDomain.DomainType
            $isDefault = [string][bool]$acceptedDomain.Default
            $matchSubDomains = [string][bool]$acceptedDomain.MatchSubDomains

            $graphDomain = Invoke-WithRetry -OperationName "Lookup tenant domain $domainName" -ScriptBlock {
                Get-MgDomain -DomainId $domainName -ErrorAction SilentlyContinue
            }

            if (-not $graphDomain) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $acceptedIdentity -Action 'GetAcceptedDomainVerificationRecords' -Status 'Completed' -Message 'Accepted domain exported. Tenant domain was not found in Entra.' -Data (New-DomainRecordData -InputDomainName $inputDomainName -AcceptedDomainIdentity $acceptedIdentity -AcceptedDomainName $acceptedName -DomainName $domainName -DomainType $domainType -AcceptedDomainIsDefault $isDefault -AcceptedDomainMatchSubDomains $matchSubDomains)))
                continue
            }

            $supportedServices = Convert-MultiValueToString -Value (Get-ObjectPropertyValue -InputObject $graphDomain -PropertyName 'SupportedServices')
            $tenantIsVerified = if ($null -eq (Get-ObjectPropertyValue -InputObject $graphDomain -PropertyName 'IsVerified')) { '' } else { [string][bool](Get-ObjectPropertyValue -InputObject $graphDomain -PropertyName 'IsVerified') }
            $tenantIsAdminManaged = if ($null -eq (Get-ObjectPropertyValue -InputObject $graphDomain -PropertyName 'IsAdminManaged')) { '' } else { [string][bool](Get-ObjectPropertyValue -InputObject $graphDomain -PropertyName 'IsAdminManaged') }
            $authenticationType = Get-StringPropertyValue -InputObject $graphDomain -PropertyName 'AuthenticationType'

            $dnsRecords = @(Invoke-WithRetry -OperationName "Load verification DNS records for $domainName" -ScriptBlock {
                Get-MgDomainVerificationDnsRecord -DomainId $domainName -All -ErrorAction Stop
            })

            if ($dnsRecords.Count -eq 0) {
                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $acceptedIdentity -Action 'GetAcceptedDomainVerificationRecords' -Status 'Completed' -Message 'Accepted domain exported. No verification DNS records were returned.' -Data (New-DomainRecordData -InputDomainName $inputDomainName -AcceptedDomainIdentity $acceptedIdentity -AcceptedDomainName $acceptedName -DomainName $domainName -DomainType $domainType -AcceptedDomainIsDefault $isDefault -AcceptedDomainMatchSubDomains $matchSubDomains -TenantDomainIsVerified $tenantIsVerified -TenantDomainIsAdminManaged $tenantIsAdminManaged -TenantDomainAuthenticationType $authenticationType -TenantDomainSupportedServices $supportedServices)))
                continue
            }

            foreach ($dnsRecord in $dnsRecords) {
                $recordType = Get-StringPropertyValue -InputObject $dnsRecord -PropertyName 'RecordType'
                $recordLabel = Get-StringPropertyValue -InputObject $dnsRecord -PropertyName 'Label'
                $recordText = Get-StringPropertyValue -InputObject $dnsRecord -PropertyName 'Text'
                $recordMailExchange = Get-StringPropertyValue -InputObject $dnsRecord -PropertyName 'MailExchange'
                $recordCanonicalName = Get-StringPropertyValue -InputObject $dnsRecord -PropertyName 'CanonicalName'
                $recordTtl = Get-StringPropertyValue -InputObject $dnsRecord -PropertyName 'Ttl'
                $recordIsOptional = if ($null -eq (Get-ObjectPropertyValue -InputObject $dnsRecord -PropertyName 'IsOptional')) { '' } else { [string][bool](Get-ObjectPropertyValue -InputObject $dnsRecord -PropertyName 'IsOptional') }
                $recordSupportedService = Get-StringPropertyValue -InputObject $dnsRecord -PropertyName 'SupportedService'

                $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $acceptedIdentity -Action 'GetAcceptedDomainVerificationRecords' -Status 'Completed' -Message 'Accepted domain verification record exported.' -Data (New-DomainRecordData -InputDomainName $inputDomainName -AcceptedDomainIdentity $acceptedIdentity -AcceptedDomainName $acceptedName -DomainName $domainName -DomainType $domainType -AcceptedDomainIsDefault $isDefault -AcceptedDomainMatchSubDomains $matchSubDomains -TenantDomainIsVerified $tenantIsVerified -TenantDomainIsAdminManaged $tenantIsAdminManaged -TenantDomainAuthenticationType $authenticationType -TenantDomainSupportedServices $supportedServices -VerificationRecordType $recordType -VerificationLabel $recordLabel -VerificationText $recordText -VerificationMailExchange $recordMailExchange -VerificationCanonicalName $recordCanonicalName -VerificationTtl $recordTtl -VerificationIsOptional $recordIsOptional -VerificationSupportedService $recordSupportedService)))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($inputDomainName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $inputDomainName -Action 'GetAcceptedDomainVerificationRecords' -Status 'Failed' -Message $_.Exception.Message -Data (New-DomainRecordData -InputDomainName $inputDomainName)))
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
Write-Status -Message 'Exchange Online domain verification-record inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
