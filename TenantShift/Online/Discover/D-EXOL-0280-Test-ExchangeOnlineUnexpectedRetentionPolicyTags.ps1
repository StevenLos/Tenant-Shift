<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-183500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Tests ExchangeOnlineUnexpectedRetentionPolicyTags and exports results to CSV.

.DESCRIPTION
    Tests ExchangeOnlineUnexpectedRetentionPolicyTags from Microsoft 365 and writes the results to a CSV file.
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
    .\SM-D3130-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.ps1 -InputCsvPath .\3130.input.csv

    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\SM-D3130-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.ps1 -DiscoverAll

    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement
    Required roles:   Exchange Administrator
    Limitations:      None known.

    CSV Fields:
    Column                          Type      Required  Description
    ------------------------------  ----      --------  -----------
    MailboxIdentity                 String    Yes       <fill in description>
    ExpectedTagNames                String    Yes       <fill in description>
    UseAssignedRetentionPolicyTags  String    Yes       <fill in description>
    Notes                           String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-EXOL-0280-Test-ExchangeOnlineUnexpectedRetentionPolicyTags_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function ConvertTo-TagSet {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    if ($null -eq $Value) {
        return $set
    }

    $items = @()
    if ($Value -is [string]) {
        $items = ConvertTo-Array -Value ([string]$Value)
    }
    elseif ($Value -is [System.Collections.IEnumerable]) {
        $items = @($Value)
    }
    else {
        $items = @([string]$Value)
    }

    foreach ($item in $items) {
        $text = Get-TrimmedValue -Value $item
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }

        [void]$set.Add($text)
    }

    return ,$set
}

function Convert-HashSetToSemicolonString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Set
    )

    if ($null -eq $Set -or $Set.Count -eq 0) {
        return ''
    }

    return (@($Set | Sort-Object) -join ';')
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

    return Get-TrimmedValue -Value (Get-ObjectPropertyValue -InputObject $InputObject -PropertyName $PropertyName)
}

function Get-MailboxLookupIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox
    )

    $primarySmtpAddress = Get-StringPropertyValue -InputObject $Mailbox -PropertyName 'PrimarySmtpAddress'
    if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
        return $primarySmtpAddress
    }

    return Get-StringPropertyValue -InputObject $Mailbox -PropertyName 'Identity'
}

$requiredHeaders = @(
    'MailboxIdentity',
    'ExpectedTagNames',
    'UseAssignedRetentionPolicyTags',
    'Notes'
)

$mailboxProperties = @(
    'RetentionPolicy'
)

$reportPropertyOrder = @(
    'TimestampUtc',
    'RowNumber',
    'PrimaryKey',
    'Action',
    'Status',
    'Message',
    'ScopeMode',
    'MailboxIdentity',
    'DisplayName',
    'PrimarySmtpAddress',
    'RetentionPolicy',
    'ExpectedTagSource',
    'ExpectedTagNames',
    'AppliedDeletePolicyTags',
    'AppliedArchivePolicyTags',
    'AppliedTagNames',
    'UnexpectedTagNames',
    'MissingExpectedTagNames',
    'HasUnexpectedTags',
    'HasMissingExpectedTags'
)

Write-Status -Message 'Starting Exchange Online unexpected-retention-policy-tag test script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$scopeMode = 'Csv'
if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. CSV input is bypassed.' -Level WARN

    $discoverRow = [ordered]@{}
    foreach ($header in $requiredHeaders) {
        if ($header -eq 'MailboxIdentity') {
            $discoverRow[$header] = '*'
        }
        else {
            $discoverRow[$header] = ''
        }
    }

    $rows = @([PSCustomObject]$discoverRow)
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentityInput = Get-TrimmedValue -Value $row.MailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentityInput)) {
            throw 'MailboxIdentity is required. Use * to test all user/shared mailboxes.'
        }

        $mailboxes = @()
        if ($mailboxIdentityInput -eq '*') {
            $mailboxes = @(Invoke-WithRetry -OperationName 'Load all user/shared mailboxes for retention tag test' -ScriptBlock {
                Get-ExchangeOnlineMailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited -Properties $mailboxProperties -ErrorAction Stop
            })
        }
        else {
            $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentityInput" -ScriptBlock {
                Get-ExchangeOnlineMailbox -Identity $mailboxIdentityInput -Properties $mailboxProperties -ErrorAction SilentlyContinue
            }

            if ($mailbox) {
                $mailboxes = @($mailbox)
            }
        }

        if ($mailboxes.Count -eq 0) {
            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityInput -Action 'TestUnexpectedRetentionPolicyTags' -Status 'NotFound' -Message 'No matching mailboxes were found.' -Data ([ordered]@{
                        MailboxIdentity                = $mailboxIdentityInput
                        DisplayName                    = ''
                        PrimarySmtpAddress             = ''
                        RetentionPolicy                = ''
                        ExpectedTagSource              = ''
                        ExpectedTagNames               = ''
                        AppliedDeletePolicyTags        = ''
                        AppliedArchivePolicyTags       = ''
                        AppliedTagNames                = ''
                        UnexpectedTagNames             = ''
                        MissingExpectedTagNames        = ''
                        HasUnexpectedTags              = ''
                        HasMissingExpectedTags         = ''
                    })))
            $rowNumber++
            continue
        }

        foreach ($mailbox in @($mailboxes | Sort-Object -Property DisplayName, PrimarySmtpAddress)) {
            $mailboxIdentityResolved = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'Identity'
            $mailboxLookupIdentity = Get-MailboxLookupIdentity -Mailbox $mailbox
            if ([string]::IsNullOrWhiteSpace($mailboxLookupIdentity)) {
                throw 'Unable to resolve a unique mailbox identity for retention tag inspection.'
            }
            if ([string]::IsNullOrWhiteSpace($mailboxIdentityResolved)) {
                $mailboxIdentityResolved = $mailboxLookupIdentity
            }

            $retentionPolicy = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'RetentionPolicy'

            $expectedTagSet = ConvertTo-TagSet -Value $row.ExpectedTagNames
            $expectedSource = [System.Collections.Generic.List[string]]::new()
            if ($expectedTagSet.Count -gt 0) {
                $expectedSource.Add('Csv')
            }

            $useAssignedPolicyTagsRaw = Get-TrimmedValue -Value $row.UseAssignedRetentionPolicyTags
            $useAssignedPolicyTags = $true
            if (-not [string]::IsNullOrWhiteSpace($useAssignedPolicyTagsRaw)) {
                $useAssignedPolicyTags = ConvertTo-Bool -Value $useAssignedPolicyTagsRaw
            }

            if ($useAssignedPolicyTags -and -not [string]::IsNullOrWhiteSpace($retentionPolicy)) {
                $policy = Invoke-WithRetry -OperationName "Load retention policy $retentionPolicy" -ScriptBlock {
                    Get-RetentionPolicy -Identity $retentionPolicy -ErrorAction SilentlyContinue
                }

                if ($policy) {
                    $retentionPolicyTagLinks = @((Get-ObjectPropertyValue -InputObject $policy -PropertyName 'RetentionPolicyTagLinks'))

                    foreach ($tagLink in $retentionPolicyTagLinks) {
                        $tagIdentity = Get-TrimmedValue -Value $tagLink
                        if ([string]::IsNullOrWhiteSpace($tagIdentity)) {
                            continue
                        }

                        $tag = Invoke-WithRetry -OperationName "Load retention policy tag $tagIdentity" -ScriptBlock {
                            Get-RetentionPolicyTag -Identity $tagIdentity -ErrorAction SilentlyContinue
                        }

                        if ($tag) {
                            $tagName = Get-TrimmedValue -Value $tag.Name
                            if (-not [string]::IsNullOrWhiteSpace($tagName)) {
                                [void]$expectedTagSet.Add($tagName)
                            }
                        }
                    }

                    if ($retentionPolicyTagLinks.Count -gt 0) {
                        $expectedSource.Add('RetentionPolicy')
                    }
                }
            }

            $folderStats = @(Invoke-WithRetry -OperationName "Load mailbox folder statistics for retention test $mailboxLookupIdentity" -ScriptBlock {
                Get-ExchangeOnlineMailboxFolderStatistics -Identity $mailboxLookupIdentity -IncludeOldestAndNewestItems -ErrorAction Stop
            })

            $appliedDeleteTagSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $appliedArchiveTagSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $appliedTagSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

            foreach ($folder in $folderStats) {
                $deleteTag = Get-StringPropertyValue -InputObject $folder -PropertyName 'DeletePolicy'
                if (-not [string]::IsNullOrWhiteSpace($deleteTag)) {
                    [void]$appliedDeleteTagSet.Add($deleteTag)
                    [void]$appliedTagSet.Add($deleteTag)
                }

                $archiveTag = Get-StringPropertyValue -InputObject $folder -PropertyName 'ArchivePolicy'
                if (-not [string]::IsNullOrWhiteSpace($archiveTag)) {
                    [void]$appliedArchiveTagSet.Add($archiveTag)
                    [void]$appliedTagSet.Add($archiveTag)
                }
            }

            $unexpectedTagSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($tag in $appliedTagSet) {
                if (-not $expectedTagSet.Contains($tag)) {
                    [void]$unexpectedTagSet.Add($tag)
                }
            }

            $missingExpectedTagSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($tag in $expectedTagSet) {
                if (-not $appliedTagSet.Contains($tag)) {
                    [void]$missingExpectedTagSet.Add($tag)
                }
            }

            $status = 'Passed'
            $message = 'No unexpected retention policy tags were found.'

            if ($expectedTagSet.Count -eq 0) {
                $status = 'Skipped'
                $message = 'No expected tag set was resolved from CSV or mailbox retention policy. Comparison skipped.'
            }
            elseif ($unexpectedTagSet.Count -gt 0) {
                $status = 'UnexpectedTags'
                $message = "Unexpected retention policy tags found: $((Convert-HashSetToSemicolonString -Set $unexpectedTagSet))."
            }
            elseif ($missingExpectedTagSet.Count -gt 0) {
                $status = 'MissingExpectedTags'
                $message = "Expected retention policy tags were not found on folders: $((Convert-HashSetToSemicolonString -Set $missingExpectedTagSet))."
            }

            $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxLookupIdentity -Action 'TestUnexpectedRetentionPolicyTags' -Status $status -Message $message -Data ([ordered]@{
                        MailboxIdentity                = $mailboxIdentityResolved
                        DisplayName                    = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'DisplayName'
                        PrimarySmtpAddress             = Get-StringPropertyValue -InputObject $mailbox -PropertyName 'PrimarySmtpAddress'
                        RetentionPolicy                = $retentionPolicy
                        ExpectedTagSource              = ($expectedSource -join ';')
                        ExpectedTagNames               = Convert-HashSetToSemicolonString -Set $expectedTagSet
                        AppliedDeletePolicyTags        = Convert-HashSetToSemicolonString -Set $appliedDeleteTagSet
                        AppliedArchivePolicyTags       = Convert-HashSetToSemicolonString -Set $appliedArchiveTagSet
                        AppliedTagNames                = Convert-HashSetToSemicolonString -Set $appliedTagSet
                        UnexpectedTagNames             = Convert-HashSetToSemicolonString -Set $unexpectedTagSet
                        MissingExpectedTagNames        = Convert-HashSetToSemicolonString -Set $missingExpectedTagSet
                        HasUnexpectedTags              = [string]($unexpectedTagSet.Count -gt 0)
                        HasMissingExpectedTags         = [string]($missingExpectedTagSet.Count -gt 0)
                    })))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentityInput) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult -RowNumber $rowNumber -PrimaryKey $mailboxIdentityInput -Action 'TestUnexpectedRetentionPolicyTags' -Status 'Failed' -Message $_.Exception.Message -Data ([ordered]@{
                    MailboxIdentity                = $mailboxIdentityInput
                    DisplayName                    = ''
                    PrimarySmtpAddress             = ''
                    RetentionPolicy                = ''
                    ExpectedTagSource              = ''
                    ExpectedTagNames               = ''
                    AppliedDeletePolicyTags        = ''
                    AppliedArchivePolicyTags       = ''
                    AppliedTagNames                = ''
                    UnexpectedTagNames             = ''
                    MissingExpectedTagNames        = ''
                    HasUnexpectedTags              = ''
                    HasMissingExpectedTags         = ''
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
Write-Status -Message 'Exchange Online unexpected-retention-policy-tag test script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
