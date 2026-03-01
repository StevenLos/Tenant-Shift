#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath ("Results_B18-Create-ExchangeResourceMailboxes_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'M365.Common.psm1') -Force -DisableNameChecking

$requiredHeaders = @(
    'ResourceType',
    'Name',
    'Alias',
    'DisplayName',
    'UserPrincipalName',
    'PrimarySmtpAddress',
    'HiddenFromAddressListsEnabled',
    'ResourceCapacity',
    'Office',
    'Phone'
)

Write-Status -Message 'Starting Exchange Online resource mailbox creation script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$newMailboxCommand = Get-Command -Name New-Mailbox -ErrorAction Stop
$newMailboxSupportsUserPrincipalName = $newMailboxCommand.Parameters.ContainsKey('UserPrincipalName')
$setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop
$setMailboxSupportsResourceCapacity = $setMailboxCommand.Parameters.ContainsKey('ResourceCapacity')
$setMailboxSupportsOffice = $setMailboxCommand.Parameters.ContainsKey('Office')
$setMailboxSupportsPhone = $setMailboxCommand.Parameters.ContainsKey('Phone')

if (-not $newMailboxSupportsUserPrincipalName) {
    Write-Status -Message "New-Mailbox in this session does not support -UserPrincipalName. The 'UserPrincipalName' CSV value will be ignored." -Level WARN
}

if (-not $setMailboxSupportsResourceCapacity) {
    Write-Status -Message "Set-Mailbox in this session does not support -ResourceCapacity. The 'ResourceCapacity' CSV value will be ignored." -Level WARN
}

if (-not $setMailboxSupportsOffice) {
    Write-Status -Message "Set-Mailbox in this session does not support -Office. The 'Office' CSV value will be ignored." -Level WARN
}

if (-not $setMailboxSupportsPhone) {
    Write-Status -Message "Set-Mailbox in this session does not support -Phone. The 'Phone' CSV value will be ignored." -Level WARN
}

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $name = ([string]$row.Name).Trim()
    $resourceTypeRaw = ([string]$row.ResourceType).Trim()
    $resourceType = $resourceTypeRaw.ToLowerInvariant()

    try {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Name is required.'
        }

        if ($resourceType -notin @('room', 'equipment')) {
            throw "ResourceType '$resourceTypeRaw' is invalid. Use Room or Equipment."
        }

        $alias = ([string]$row.Alias).Trim()
        $displayName = ([string]$row.DisplayName).Trim()
        $userPrincipalName = ([string]$row.UserPrincipalName).Trim()
        $primarySmtpAddress = ([string]$row.PrimarySmtpAddress).Trim()
        $office = ([string]$row.Office).Trim()
        $phone = ([string]$row.Phone).Trim()

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

        $existingMailbox = Invoke-WithRetry -OperationName "Lookup resource mailbox $lookupIdentity" -ScriptBlock {
            Get-Mailbox -Identity $lookupIdentity -ErrorAction SilentlyContinue
        }

        if ($existingMailbox) {
            $recipientTypeDetails = ([string]$existingMailbox.RecipientTypeDetails).Trim()
            if ($resourceType -eq 'room' -and $recipientTypeDetails -eq 'RoomMailbox') {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateResourceMailbox' -Status 'Skipped' -Message 'Room mailbox already exists.'))
                $rowNumber++
                continue
            }

            if ($resourceType -eq 'equipment' -and $recipientTypeDetails -eq 'EquipmentMailbox') {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateResourceMailbox' -Status 'Skipped' -Message 'Equipment mailbox already exists.'))
                $rowNumber++
                continue
            }

            throw "Mailbox '$lookupIdentity' already exists with recipient type '$recipientTypeDetails', which does not match requested resource type '$resourceTypeRaw'."
        }

        $params = @{
            Name = $name
        }

        if ($resourceType -eq 'room') {
            $params.Room = $true
        }
        else {
            $params.Equipment = $true
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

        $resourceCapacityRaw = ([string]$row.ResourceCapacity).Trim()
        $resourceCapacity = 0
        $setResourceCapacity = $false
        if (-not [string]::IsNullOrWhiteSpace($resourceCapacityRaw)) {
            if (-not [int]::TryParse($resourceCapacityRaw, [ref]$resourceCapacity)) {
                throw "ResourceCapacity '$resourceCapacityRaw' is not a valid integer."
            }

            if ($resourceCapacity -lt 0) {
                throw 'ResourceCapacity must be zero or greater.'
            }

            $setResourceCapacity = $true
        }

        $hiddenRaw = ([string]$row.HiddenFromAddressListsEnabled).Trim()
        $setHidden = -not [string]::IsNullOrWhiteSpace($hiddenRaw)

        $resourceTypeTitle = if ($resourceType -eq 'room') { 'conference room' } else { 'equipment' }
        if ($PSCmdlet.ShouldProcess($lookupIdentity, "Create Exchange Online $resourceTypeTitle mailbox")) {
            $createdMailbox = Invoke-WithRetry -OperationName "Create resource mailbox $lookupIdentity" -ScriptBlock {
                New-Mailbox @params -ErrorAction Stop
            }

            if ($setHidden) {
                $hiddenValue = ConvertTo-Bool -Value $hiddenRaw
                Invoke-WithRetry -OperationName "Set hidden from GAL for $lookupIdentity" -ScriptBlock {
                    Set-Mailbox -Identity $createdMailbox.Identity -HiddenFromAddressListsEnabled $hiddenValue -ErrorAction Stop
                }
            }

            $ignoredUpdates = [System.Collections.Generic.List[string]]::new()
            if ($setResourceCapacity) {
                if ($setMailboxSupportsResourceCapacity) {
                    Invoke-WithRetry -OperationName "Set resource capacity for $lookupIdentity" -ScriptBlock {
                        Set-Mailbox -Identity $createdMailbox.Identity -ResourceCapacity $resourceCapacity -ErrorAction Stop
                    }
                }
                else {
                    $ignoredUpdates.Add('ResourceCapacity')
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($office)) {
                if ($setMailboxSupportsOffice) {
                    Invoke-WithRetry -OperationName "Set office for $lookupIdentity" -ScriptBlock {
                        Set-Mailbox -Identity $createdMailbox.Identity -Office $office -ErrorAction Stop
                    }
                }
                else {
                    $ignoredUpdates.Add('Office')
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($phone)) {
                if ($setMailboxSupportsPhone) {
                    Invoke-WithRetry -OperationName "Set phone for $lookupIdentity" -ScriptBlock {
                        Set-Mailbox -Identity $createdMailbox.Identity -Phone $phone -ErrorAction Stop
                    }
                }
                else {
                    $ignoredUpdates.Add('Phone')
                }
            }

            $successMessage = "$resourceTypeTitle mailbox created successfully."
            if ($upnIgnored) {
                $successMessage = "$successMessage UserPrincipalName was provided but ignored because this New-Mailbox session does not support -UserPrincipalName."
            }

            if ($ignoredUpdates.Count -gt 0) {
                $successMessage = "$successMessage Ignored unsupported update field(s): $($ignoredUpdates -join ', ')."
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateResourceMailbox' -Status 'Created' -Message $successMessage))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $lookupIdentity -Action 'CreateResourceMailbox' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($name) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $name -Action 'CreateResourceMailbox' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online resource mailbox creation script completed.' -Level SUCCESS

