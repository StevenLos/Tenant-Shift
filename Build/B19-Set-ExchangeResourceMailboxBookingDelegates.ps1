#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath ("Results_B19-Set-ExchangeResourceMailboxBookingDelegates_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'M365.Common.psm1') -Force -DisableNameChecking

$requiredHeaders = @(
    'ResourceMailboxIdentity',
    'DelegateUserPrincipalNames',
    'AutomateProcessing',
    'ForwardRequestsToDelegates',
    'AllBookInPolicy',
    'AllRequestInPolicy',
    'AllRequestOutOfPolicy'
)

Write-Status -Message 'Starting resource mailbox booking delegate script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $resourceMailboxIdentity = ([string]$row.ResourceMailboxIdentity).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($resourceMailboxIdentity)) {
            throw 'ResourceMailboxIdentity is required.'
        }

        $delegateUpns = ConvertTo-Array -Value ([string]$row.DelegateUserPrincipalNames)
        if ($delegateUpns.Count -eq 0) {
            throw 'DelegateUserPrincipalNames is required and must contain at least one value.'
        }

        $automateProcessing = ([string]$row.AutomateProcessing).Trim()
        if ([string]::IsNullOrWhiteSpace($automateProcessing)) {
            $automateProcessing = 'AutoAccept'
        }

        if ($automateProcessing -notin @('AutoAccept', 'AutoUpdate', 'None')) {
            throw "AutomateProcessing '$automateProcessing' is invalid. Use AutoAccept, AutoUpdate, or None."
        }

        $forwardRequestsToDelegates = ConvertTo-Bool -Value $row.ForwardRequestsToDelegates -Default $true
        $allBookInPolicy = ConvertTo-Bool -Value $row.AllBookInPolicy -Default $false
        $allRequestInPolicy = ConvertTo-Bool -Value $row.AllRequestInPolicy -Default $false
        $allRequestOutOfPolicy = ConvertTo-Bool -Value $row.AllRequestOutOfPolicy -Default $false

        $mailbox = Invoke-WithRetry -OperationName "Lookup resource mailbox $resourceMailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $resourceMailboxIdentity -ErrorAction Stop
        }

        $recipientTypeDetails = ([string]$mailbox.RecipientTypeDetails).Trim()
        if ($recipientTypeDetails -notin @('RoomMailbox', 'EquipmentMailbox')) {
            throw "Mailbox '$resourceMailboxIdentity' is '$recipientTypeDetails'. Only RoomMailbox and EquipmentMailbox are supported."
        }

        $calendarProcessing = Invoke-WithRetry -OperationName "Load calendar processing $resourceMailboxIdentity" -ScriptBlock {
            Get-CalendarProcessing -Identity $mailbox.Identity -ErrorAction Stop
        }

        $currentDelegateMap = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($existingDelegate in @($calendarProcessing.ResourceDelegates)) {
            if ($null -eq $existingDelegate) {
                continue
            }

            $delegateIdentity = ([string]$existingDelegate).Trim()
            if ([string]::IsNullOrWhiteSpace($delegateIdentity)) {
                continue
            }

            try {
                $delegateRecipient = Invoke-WithRetry -OperationName "Resolve existing delegate $delegateIdentity" -ScriptBlock {
                    Get-Recipient -Identity $delegateIdentity -ErrorAction Stop
                }
                $delegateKey = ([string]$delegateRecipient.DistinguishedName).Trim().ToLowerInvariant()
            }
            catch {
                $delegateKey = $delegateIdentity.ToLowerInvariant()
            }

            if (-not $currentDelegateMap.ContainsKey($delegateKey)) {
                $currentDelegateMap[$delegateKey] = $delegateIdentity
            }
        }

        $desiredDelegateMap = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($delegateUpn in $delegateUpns) {
            $delegateRecipient = Invoke-WithRetry -OperationName "Lookup booking delegate $delegateUpn" -ScriptBlock {
                Get-Recipient -Identity $delegateUpn -ErrorAction Stop
            }

            $delegateKey = ([string]$delegateRecipient.DistinguishedName).Trim().ToLowerInvariant()
            if ([string]::IsNullOrWhiteSpace($delegateKey)) {
                throw "Delegate '$delegateUpn' does not have a DistinguishedName."
            }

            if (-not $desiredDelegateMap.ContainsKey($delegateKey)) {
                $desiredDelegateMap[$delegateKey] = $delegateRecipient
            }
        }

        $delegatesToAdd = [System.Collections.Generic.List[string]]::new()
        foreach ($desiredPair in $desiredDelegateMap.GetEnumerator()) {
            if (-not $currentDelegateMap.ContainsKey($desiredPair.Key)) {
                $delegatesToAdd.Add([string]$desiredPair.Value.Identity)
            }
        }

        $finalDelegateIdentities = [System.Collections.Generic.List[string]]::new()
        foreach ($existingIdentity in $currentDelegateMap.Values) {
            if (-not [string]::IsNullOrWhiteSpace($existingIdentity)) {
                $finalDelegateIdentities.Add($existingIdentity)
            }
        }
        foreach ($delegateIdentityToAdd in $delegatesToAdd) {
            if (-not [string]::IsNullOrWhiteSpace($delegateIdentityToAdd)) {
                $finalDelegateIdentities.Add($delegateIdentityToAdd)
            }
        }

        $dedupedFinalDelegateIdentities = [System.Collections.Generic.List[string]]::new()
        $seenDelegateIdentityValues = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($candidateIdentity in $finalDelegateIdentities) {
            if ([string]::IsNullOrWhiteSpace($candidateIdentity)) {
                continue
            }

            if ($seenDelegateIdentityValues.Add($candidateIdentity)) {
                $dedupedFinalDelegateIdentities.Add($candidateIdentity)
            }
        }

        $automateMatches = ([string]$calendarProcessing.AutomateProcessing).Trim().Equals($automateProcessing, [System.StringComparison]::OrdinalIgnoreCase)
        $forwardMatches = ([bool]$calendarProcessing.ForwardRequestsToDelegates) -eq $forwardRequestsToDelegates
        $allBookMatches = ([bool]$calendarProcessing.AllBookInPolicy) -eq $allBookInPolicy
        $allRequestInMatches = ([bool]$calendarProcessing.AllRequestInPolicy) -eq $allRequestInPolicy
        $allRequestOutMatches = ([bool]$calendarProcessing.AllRequestOutOfPolicy) -eq $allRequestOutOfPolicy

        $needsDelegateUpdate = $delegatesToAdd.Count -gt 0
        $needsPolicyUpdate = -not ($automateMatches -and $forwardMatches -and $allBookMatches -and $allRequestInMatches -and $allRequestOutMatches)

        if (-not ($needsDelegateUpdate -or $needsPolicyUpdate)) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'SetResourceMailboxBookingDelegates' -Status 'Skipped' -Message 'Booking delegates and policy settings already match requested values.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($resourceMailboxIdentity, 'Set resource mailbox booking delegates and policy')) {
            $setParams = @{
                Identity                   = $mailbox.Identity
                ResourceDelegates          = $dedupedFinalDelegateIdentities.ToArray()
                AutomateProcessing         = $automateProcessing
                ForwardRequestsToDelegates = $forwardRequestsToDelegates
                AllBookInPolicy            = $allBookInPolicy
                AllRequestInPolicy         = $allRequestInPolicy
                AllRequestOutOfPolicy      = $allRequestOutOfPolicy
                ErrorAction                = 'Stop'
            }

            Invoke-WithRetry -OperationName "Set booking delegates $resourceMailboxIdentity" -ScriptBlock {
                Set-CalendarProcessing @setParams
            }

            $messageParts = [System.Collections.Generic.List[string]]::new()
            if ($needsDelegateUpdate) {
                $messageParts.Add("Delegates added: $($delegatesToAdd.Count).")
            }
            else {
                $messageParts.Add('Delegates unchanged.')
            }

            if ($needsPolicyUpdate) {
                $messageParts.Add('Booking policy settings updated.')
            }
            else {
                $messageParts.Add('Booking policy settings unchanged.')
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'SetResourceMailboxBookingDelegates' -Status 'Completed' -Message ($messageParts -join ' ')))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'SetResourceMailboxBookingDelegates' -Status 'WhatIf' -Message 'Booking delegate update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($resourceMailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $resourceMailboxIdentity -Action 'SetResourceMailboxBookingDelegates' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Resource mailbox booking delegate script completed.' -Level SUCCESS

