#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath ("Results_B03-Assign-EntraUserLicenses_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'M365.Common.psm1') -Force -DisableNameChecking

$requiredHeaders = @(
    'UserPrincipalName',
    'SkuPartNumber',
    'DisabledPlans'
)

Write-Status -Message 'Starting Entra ID user license assignment script.'
Assert-ModuleCurrent -ModuleNames @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Users.Actions',
    'Microsoft.Graph.Identity.DirectoryManagement'
)
Ensure-GraphConnection -RequiredScopes @('User.ReadWrite.All', 'Directory.Read.All', 'Organization.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

Write-Status -Message 'Loading subscribed SKUs from Microsoft Graph.'
$subscribedSkus = @(Invoke-WithRetry -OperationName 'Load subscribed SKUs' -ScriptBlock {
    Get-MgSubscribedSku -All -ErrorAction Stop
})

if ($subscribedSkus.Count -eq 0) {
    throw 'No subscribed SKUs were returned from Microsoft Graph. Confirm that your tenant has at least one available license SKU.'
}

$skuByPartNumber = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$availableUnitsByPartNumber = [System.Collections.Generic.Dictionary[string, int]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($sku in $subscribedSkus) {
    $partNumber = ([string]$sku.SkuPartNumber).Trim()
    if ([string]::IsNullOrWhiteSpace($partNumber)) {
        continue
    }

    if (-not $skuByPartNumber.ContainsKey($partNumber)) {
        $skuByPartNumber[$partNumber] = $sku
    }

    try {
        $enabledUnits = $null
        $consumedUnits = $null

        if ($sku.PSObject.Properties.Name -contains 'PrepaidUnits' -and $sku.PrepaidUnits -and $sku.PrepaidUnits.PSObject.Properties.Name -contains 'Enabled') {
            $enabledUnits = [int]$sku.PrepaidUnits.Enabled
        }

        if ($sku.PSObject.Properties.Name -contains 'ConsumedUnits' -and $null -ne $sku.ConsumedUnits) {
            $consumedUnits = [int]$sku.ConsumedUnits
        }

        if ($null -ne $enabledUnits -and $null -ne $consumedUnits) {
            $availableUnitsByPartNumber[$partNumber] = [Math]::Max($enabledUnits - $consumedUnits, 0)
        }
    }
    catch {
        Write-Status -Message "Unable to calculate available unit count for SKU '$partNumber': $($_.Exception.Message)" -Level WARN
    }
}

if ($skuByPartNumber.Count -eq 0) {
    throw 'No subscribed SKUs with usable SkuPartNumber values were returned from Microsoft Graph.'
}

$userByUpn = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$assignedLicensesByUpn = [System.Collections.Generic.Dictionary[string, object[]]]::new([System.StringComparer]::OrdinalIgnoreCase)

$rowNumber = 1
foreach ($row in $rows) {
    $upn = ([string]$row.UserPrincipalName).Trim()
    $skuPartNumber = ([string]$row.SkuPartNumber).Trim()
    $primaryKey = "$upn|$skuPartNumber"

    try {
        if ([string]::IsNullOrWhiteSpace($upn) -or [string]::IsNullOrWhiteSpace($skuPartNumber)) {
            throw 'UserPrincipalName and SkuPartNumber are required.'
        }

        if (-not $skuByPartNumber.ContainsKey($skuPartNumber)) {
            $availableSkuPartNumbers = @($skuByPartNumber.Keys | Sort-Object)
            throw "SkuPartNumber '$skuPartNumber' was not found in tenant subscriptions. Available values: $($availableSkuPartNumbers -join ', ')"
        }

        $sku = $skuByPartNumber[$skuPartNumber]
        $skuId = [Guid]$sku.SkuId
        $servicePlans = @($sku.ServicePlans)

        $disabledPlanTokens = ConvertTo-Array -Value ([string]$row.DisabledPlans)
        $resolvedDisabledPlans = [System.Collections.Generic.List[Guid]]::new()

        foreach ($token in $disabledPlanTokens) {
            $tokenValue = ([string]$token).Trim()
            if ([string]::IsNullOrWhiteSpace($tokenValue)) {
                continue
            }

            $planId = [Guid]::Empty
            if ([Guid]::TryParse($tokenValue, [ref]$planId)) {
                $planExists = $false
                foreach ($servicePlan in $servicePlans) {
                    $existingPlanId = [Guid]::Empty
                    if ([Guid]::TryParse([string]$servicePlan.ServicePlanId, [ref]$existingPlanId) -and $existingPlanId -eq $planId) {
                        $planExists = $true
                        break
                    }
                }

                if (-not $planExists) {
                    throw "Disabled plan GUID '$tokenValue' is not valid for SKU '$skuPartNumber'."
                }
            }
            else {
                $matchingPlanIds = [System.Collections.Generic.List[Guid]]::new()

                foreach ($servicePlan in $servicePlans) {
                    $servicePlanName = ([string]$servicePlan.ServicePlanName).Trim()
                    if ([string]::IsNullOrWhiteSpace($servicePlanName)) {
                        continue
                    }

                    if ($servicePlanName.Equals($tokenValue, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $servicePlanId = [Guid]::Empty
                        if (-not [Guid]::TryParse([string]$servicePlan.ServicePlanId, [ref]$servicePlanId)) {
                            throw "Service plan '$servicePlanName' on SKU '$skuPartNumber' has an invalid ID."
                        }

                        $matchingPlanIds.Add($servicePlanId)
                    }
                }

                if ($matchingPlanIds.Count -eq 0) {
                    $availablePlans = @(
                        $servicePlans |
                            ForEach-Object { ([string]$_.ServicePlanName).Trim() } |
                            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                            Sort-Object -Unique
                    )
                    throw "Disabled plan '$tokenValue' is not valid for SKU '$skuPartNumber'. Available plans: $($availablePlans -join ', ')"
                }

                if ($matchingPlanIds.Count -gt 1) {
                    throw "Disabled plan '$tokenValue' matched multiple service plans for SKU '$skuPartNumber'. Use a service plan GUID to disambiguate."
                }

                $planId = $matchingPlanIds[0]
            }

            if (-not ($resolvedDisabledPlans -contains $planId)) {
                $resolvedDisabledPlans.Add($planId)
            }
        }

        if ($userByUpn.ContainsKey($upn)) {
            $user = $userByUpn[$upn]
        }
        else {
            $escapedUpn = Escape-ODataString -Value $upn
            $users = @(Invoke-WithRetry -OperationName "Lookup user $upn" -ScriptBlock {
                Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -ErrorAction Stop
            })

            if ($users.Count -eq 0) {
                throw "User '$upn' was not found."
            }

            if ($users.Count -gt 1) {
                throw "Multiple users were returned for UPN '$upn'. Resolve duplicate directory objects before retrying."
            }

            $user = $users[0]
            $userByUpn[$upn] = $user
        }

        if ($assignedLicensesByUpn.ContainsKey($upn)) {
            $assignedLicenses = @($assignedLicensesByUpn[$upn])
        }
        else {
            $userWithLicenses = Invoke-WithRetry -OperationName "Load assigned licenses for $upn" -ScriptBlock {
                Get-MgUser -UserId $user.Id -Property 'id,userPrincipalName,assignedLicenses' -ErrorAction Stop
            }
            $assignedLicenses = @($userWithLicenses.AssignedLicenses)
            $assignedLicensesByUpn[$upn] = $assignedLicenses
        }

        $currentAssignments = @($assignedLicenses | Where-Object { [Guid]$_.SkuId -eq $skuId })
        if ($currentAssignments.Count -gt 1) {
            throw "User '$upn' has duplicate assignments for SKU '$skuPartNumber'. Resolve license state before retrying."
        }

        $requestedDisabledCanonical = @(
            $resolvedDisabledPlans |
                ForEach-Object { $_.Guid.ToLowerInvariant() } |
                Sort-Object -Unique
        )

        $currentDisabledCanonical = @()
        if ($currentAssignments.Count -eq 1) {
            $currentDisabledCanonical = @(
                @($currentAssignments[0].DisabledPlans) |
                    Where-Object { $null -ne $_ } |
                    ForEach-Object {
                        $currentDisabledGuid = [Guid]::Empty
                        if (-not [Guid]::TryParse([string]$_, [ref]$currentDisabledGuid)) {
                            throw "User '$upn' has an invalid disabled plan ID '$([string]$_)' for SKU '$skuPartNumber'."
                        }

                        $currentDisabledGuid.Guid.ToLowerInvariant()
                    } |
                    Sort-Object -Unique
            )
        }

        $requestedMinusCurrent = @($requestedDisabledCanonical | Where-Object { $_ -notin $currentDisabledCanonical })
        $currentMinusRequested = @($currentDisabledCanonical | Where-Object { $_ -notin $requestedDisabledCanonical })
        $disabledPlansMatch = ($requestedMinusCurrent.Count -eq 0 -and $currentMinusRequested.Count -eq 0)

        if ($currentAssignments.Count -eq 1 -and $disabledPlansMatch) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetUserLicense' -Status 'Skipped' -Message "License '$skuPartNumber' is already assigned with the requested disabled plans."))
            $rowNumber++
            continue
        }

        $isNewAssignment = ($currentAssignments.Count -eq 0)

        $addLicense = @{
            SkuId         = $skuId
            DisabledPlans = @($resolvedDisabledPlans)
        }

        $operationLabel = if ($isNewAssignment) { 'Assign license' } else { 'Update license options' }
        if ($PSCmdlet.ShouldProcess($primaryKey, $operationLabel)) {
            if ($isNewAssignment -and $availableUnitsByPartNumber.ContainsKey($skuPartNumber) -and $availableUnitsByPartNumber[$skuPartNumber] -le 0) {
                throw "No available units remain for SKU '$skuPartNumber'."
            }

            Invoke-WithRetry -OperationName "$operationLabel $primaryKey" -ScriptBlock {
                Set-MgUserLicense -UserId $user.Id -AddLicenses @($addLicense) -RemoveLicenses @() -ErrorAction Stop | Out-Null
            }

            $refreshedUser = Invoke-WithRetry -OperationName "Refresh assigned licenses for $upn" -ScriptBlock {
                Get-MgUser -UserId $user.Id -Property 'id,userPrincipalName,assignedLicenses' -ErrorAction Stop
            }
            $assignedLicensesByUpn[$upn] = @($refreshedUser.AssignedLicenses)

            if ($isNewAssignment -and $availableUnitsByPartNumber.ContainsKey($skuPartNumber)) {
                $availableUnitsByPartNumber[$skuPartNumber] = [Math]::Max(($availableUnitsByPartNumber[$skuPartNumber] - 1), 0)
            }

            $resultStatus = if ($isNewAssignment) { 'Assigned' } else { 'Updated' }
            $disabledPlanSummary = if ($resolvedDisabledPlans.Count -gt 0) {
                " Disabled plans applied: $($resolvedDisabledPlans.Count)."
            }
            else {
                ' No disabled plans were specified.'
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetUserLicense' -Status $resultStatus -Message "License '$skuPartNumber' processed successfully.$disabledPlanSummary"))
        }
        else {
            $whatIfMessage = if ($isNewAssignment) {
                'License assignment skipped due to WhatIf.'
            }
            else {
                'License option update skipped due to WhatIf.'
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetUserLicense' -Status 'WhatIf' -Message $whatIfMessage))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($primaryKey) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $primaryKey -Action 'SetUserLicense' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID user license assignment script completed.' -Level SUCCESS

