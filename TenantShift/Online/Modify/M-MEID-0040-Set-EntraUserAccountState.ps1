<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-120000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Modifies EntraUserAccountState in Microsoft 365.

.DESCRIPTION
    Updates EntraUserAccountState in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one modify operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating
    what changed or why a row was skipped.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-M3002-Set-EntraUserAccountState.ps1 -InputCsvPath .\3002.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-M3002-Set-EntraUserAccountState.ps1 -InputCsvPath .\3002.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Users
    Required roles:   Global Administrator or appropriate workload-specific role
    Limitations:      None known.

    CSV Fields:
    Column                Type      Required  Description
    --------------------  ----      --------  -----------
    Action                String    Yes       <fill in description>
    Notes                 String    Yes       <fill in description>
    UserPrincipalName     String    Yes       <fill in description>
    AccountEnabled        String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3002-Set-EntraUserAccountState_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

$requiredHeaders = @(
    'Action',
    'Notes',
    'UserPrincipalName',
    'AccountEnabled'
)

Write-Status -Message 'Starting Entra ID user account-state script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('User.ReadWrite.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $upn = Get-TrimmedValue -Value $row.UserPrincipalName

    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw 'UserPrincipalName is required.'
        }

        $accountEnabledRaw = Get-TrimmedValue -Value $row.AccountEnabled
        if ([string]::IsNullOrWhiteSpace($accountEnabledRaw)) {
            throw 'AccountEnabled is required and must be TRUE/FALSE.'
        }

        $targetState = ConvertTo-Bool -Value $accountEnabledRaw

        $escapedUpn = Escape-ODataString -Value $upn
        $users = @(Invoke-WithRetry -OperationName "Lookup user $upn" -ScriptBlock {
            Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -Property 'id,userPrincipalName,accountEnabled' -ErrorAction Stop
        })

        if ($users.Count -eq 0) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'SetUserAccountState' -Status 'NotFound' -Message 'User not found.'))
            $rowNumber++
            continue
        }

        if ($users.Count -gt 1) {
            throw "Multiple users were returned for UPN '$upn'. Resolve duplicate directory objects before retrying."
        }

        $user = $users[0]
        $currentState = [bool]$user.AccountEnabled
        if ($currentState -eq $targetState) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'SetUserAccountState' -Status 'Skipped' -Message "AccountEnabled already set to '$targetState'."))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($upn, "Set AccountEnabled to $targetState")) {
            Invoke-WithRetry -OperationName "Set account state $upn" -ScriptBlock {
                Update-MgUser -UserId $user.Id -BodyParameter @{ accountEnabled = $targetState } -ErrorAction Stop | Out-Null
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'SetUserAccountState' -Status 'Updated' -Message "AccountEnabled set to '$targetState'."))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'SetUserAccountState' -Status 'WhatIf' -Message 'Account-state update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $upn -Action 'SetUserAccountState' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID user account-state script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
