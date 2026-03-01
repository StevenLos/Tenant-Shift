#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath ("Results_B16-Create-ExchangeSharedMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'M365.Common.psm1') -Force -DisableNameChecking

$requiredHeaders = @(
    'Name',
    'Alias',
    'DisplayName',
    'UserPrincipalName',
    'PrimarySmtpAddress',
    'HiddenFromAddressListsEnabled',
    'GrantSendOnBehalfTo'
)

Write-Status -Message 'Starting shared mailbox creation script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection
$newMailboxSupportsUserPrincipalName = (Get-Command -Name New-Mailbox -ErrorAction Stop).Parameters.ContainsKey('UserPrincipalName')

if (-not $newMailboxSupportsUserPrincipalName) {
    Write-Status -Message "New-Mailbox in this session does not support -UserPrincipalName. The 'UserPrincipalName' CSV value will be ignored." -Level WARN
}

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $name = ([string]$row.Name).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Name is required.'
        }

        $alias = ([string]$row.Alias).Trim()
        $displayName = ([string]$row.DisplayName).Trim()
        $userPrincipalName = ([string]$row.UserPrincipalName).Trim()
        $primarySmtpAddress = ([string]$row.PrimarySmtpAddress).Trim()

        $lookupIdentity = if (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) {
            $userPrincipalName
        }
        elseif (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $primarySmtpAddress
        }
        elseif (-not [string]::IsNullOrWhiteSpace($alias)) {
            $alias
        }
        else {
            $name
        }

        $existingMailbox = Invoke-WithRetry -OperationName "Lookup shared mailbox $lookupIdentity" -ScriptBlock {
            Get-Mailbox -Identity $lookupIdentity -ErrorAction SilentlyContinue
        }
        if ($existingMailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateSharedMailbox' -Status 'Skipped' -Message 'Shared mailbox already exists.'))
            $rowNumber++
            continue
        }

        $params = @{
            Shared = $true
            Name   = $name
        }

        if (-not [string]::IsNullOrWhiteSpace($alias)) {
            $params.Alias = $alias
        }

        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $params.DisplayName = $displayName
        }

        $upnIgnored = $false
        if (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) {
            if ($newMailboxSupportsUserPrincipalName) {
                $params.UserPrincipalName = $userPrincipalName
            }
            else {
                $upnIgnored = $true
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            $params.PrimarySmtpAddress = $primarySmtpAddress
        }

        $hiddenRaw = ([string]$row.HiddenFromAddressListsEnabled).Trim()
        $setHidden = -not [string]::IsNullOrWhiteSpace($hiddenRaw)
        $sendOnBehalfList = ConvertTo-Array -Value ([string]$row.GrantSendOnBehalfTo)

        if ($PSCmdlet.ShouldProcess($lookupIdentity, 'Create Exchange Online shared mailbox')) {
            $createdMailbox = Invoke-WithRetry -OperationName "Create shared mailbox $lookupIdentity" -ScriptBlock {
                New-Mailbox @params -ErrorAction Stop
            }

            if ($setHidden) {
                $hiddenValue = ConvertTo-Bool -Value $hiddenRaw
                Invoke-WithRetry -OperationName "Set hidden from GAL for $lookupIdentity" -ScriptBlock {
                    Set-Mailbox -Identity $createdMailbox.Identity -HiddenFromAddressListsEnabled $hiddenValue -ErrorAction Stop
                }
            }

            if ($sendOnBehalfList.Count -gt 0) {
                Invoke-WithRetry -OperationName "Grant send on behalf for $lookupIdentity" -ScriptBlock {
                    Set-Mailbox -Identity $createdMailbox.Identity -GrantSendOnBehalfTo @{ Add = $sendOnBehalfList } -ErrorAction Stop
                }
            }

            $successMessage = 'Shared mailbox created successfully.'
            if ($upnIgnored) {
                $successMessage = "$successMessage UserPrincipalName was provided but ignored because this New-Mailbox session does not support -UserPrincipalName."
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateSharedMailbox' -Status 'Created' -Message $successMessage))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateSharedMailbox' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($name) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $name -Action 'CreateSharedMailbox' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Shared mailbox creation script completed.' -Level SUCCESS

