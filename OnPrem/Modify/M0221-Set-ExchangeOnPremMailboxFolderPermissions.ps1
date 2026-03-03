<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260303-235500

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
Exchange Management Shell cmdlets (session)

.MODULEVERSIONPOLICY
Exchange on-prem cmdlets are validated by session command checks (offline-safe)
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M0221-Set-ExchangeOnPremMailboxFolderPermissions_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\OnPrem\OnPrem.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function ConvertTo-CanonicalStringSet {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$Values
    )

    if ($null -eq $Values) {
        return ,@()
    }

    return ,@(
        @($Values) |
            ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
}

function Test-SameStringSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Left,

        [Parameter(Mandatory)]
        [string[]]$Right
    )

    $leftOnly = @($Left | Where-Object { $_ -notin $Right })
    $rightOnly = @($Right | Where-Object { $_ -notin $Left })
    return ($leftOnly.Count -eq 0 -and $rightOnly.Count -eq 0)
}

$requiredHeaders = @(
    'MailboxIdentity',
    'TrusteeIdentity',
    'FolderPath',
    'PermissionAction',
    'AccessRights',
    'CalendarDelegate',
    'CalendarCanViewPrivateItems',
    'SendNotificationToUser'
)

Write-Status -Message 'Starting Exchange on-prem mailbox folder permission script.'
Ensure-ExchangeOnPremConnection

$addMailboxFolderPermissionCommand = Get-Command -Name Add-MailboxFolderPermission -ErrorAction Stop
$setMailboxFolderPermissionCommand = Get-Command -Name Set-MailboxFolderPermission -ErrorAction Stop
$removeMailboxFolderPermissionCommand = Get-Command -Name Remove-MailboxFolderPermission -ErrorAction Stop
$addSupportsSharingFlags = $addMailboxFolderPermissionCommand.Parameters.ContainsKey('SharingPermissionFlags')
$setSupportsSharingFlags = $setMailboxFolderPermissionCommand.Parameters.ContainsKey('SharingPermissionFlags')
$addSupportsSendNotification = $addMailboxFolderPermissionCommand.Parameters.ContainsKey('SendNotificationToUser')
$setSupportsSendNotification = $setMailboxFolderPermissionCommand.Parameters.ContainsKey('SendNotificationToUser')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = Get-TrimmedValue -Value $row.MailboxIdentity
    $trusteeIdentity = Get-TrimmedValue -Value $row.TrusteeIdentity
    $folderPathRaw = Get-TrimmedValue -Value $row.FolderPath

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity) -or [string]::IsNullOrWhiteSpace($trusteeIdentity) -or [string]::IsNullOrWhiteSpace($folderPathRaw)) {
            throw 'MailboxIdentity, TrusteeIdentity, and FolderPath are required.'
        }

        $permissionActionRaw = Get-TrimmedValue -Value $row.PermissionAction
        $permissionAction = if ([string]::IsNullOrWhiteSpace($permissionActionRaw)) { 'Set' } else { $permissionActionRaw }
        if ($permissionAction -notin @('Add', 'Set', 'Remove')) {
            throw "PermissionAction '$permissionAction' is invalid. Use Add, Set, or Remove."
        }

        $accessRights = ConvertTo-Array -Value (Get-TrimmedValue -Value $row.AccessRights)
        if ($permissionAction -ne 'Remove' -and $accessRights.Count -eq 0) {
            throw 'AccessRights is required for Add/Set actions and must contain at least one value.'
        }

        $calendarDelegate = ConvertTo-Bool -Value $row.CalendarDelegate -Default $false
        $calendarCanViewPrivateItems = ConvertTo-Bool -Value $row.CalendarCanViewPrivateItems -Default $false
        $sendNotificationToUser = ConvertTo-Bool -Value $row.SendNotificationToUser -Default $false

        if ($calendarCanViewPrivateItems -and -not $calendarDelegate) {
            throw 'CalendarCanViewPrivateItems requires CalendarDelegate = TRUE.'
        }

        $folderIdentity = if ($folderPathRaw -match '^[^:]+:\\') {
            $folderPathRaw
        }
        else {
            "${mailboxIdentity}:\\$folderPathRaw"
        }

        $folderPathForCheck = if ($folderPathRaw -match '^[^:]+:\\') {
            $folderPathRaw.Split(':\\', 2)[1]
        }
        else {
            $folderPathRaw
        }

        $folderSegments = @($folderPathForCheck -split '[\\/]' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $folderLeaf = ''
        if ($folderSegments.Count -gt 0) {
            $folderLeaf = Get-TrimmedValue -Value $folderSegments[-1]
        }
        $isCalendarFolder = $folderLeaf.Equals('Calendar', [System.StringComparison]::OrdinalIgnoreCase)

        if (($calendarDelegate -or $calendarCanViewPrivateItems -or $sendNotificationToUser) -and -not $isCalendarFolder) {
            throw 'CalendarDelegate, CalendarCanViewPrivateItems, and SendNotificationToUser are only supported when FolderPath targets Calendar.'
        }

        if ($isCalendarFolder -and ($calendarDelegate -or $calendarCanViewPrivateItems) -and (-not $addSupportsSharingFlags -or -not $setSupportsSharingFlags)) {
            throw 'This Exchange on-prem cmdlet session does not support -SharingPermissionFlags required for calendar delegate configuration.'
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $mailboxIdentity -ErrorAction SilentlyContinue
        }
        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|$folderPathRaw" -Action 'SetMailboxFolderPermission' -Status 'NotFound' -Message 'Mailbox not found.'))
            $rowNumber++
            continue
        }

        $trustee = Invoke-WithRetry -OperationName "Lookup trustee $trusteeIdentity" -ScriptBlock {
            Get-Recipient -Identity $trusteeIdentity -ErrorAction SilentlyContinue
        }
        if (-not $trustee) {
            throw "Trustee '$trusteeIdentity' was not found."
        }

        $existingPermissions = @(Invoke-WithRetry -OperationName "Check folder permission $folderIdentity -> $trusteeIdentity" -ScriptBlock {
            Get-MailboxFolderPermission -Identity $folderIdentity -User $trustee.Identity -ErrorAction SilentlyContinue
        })
        $existingPermission = if ($existingPermissions.Count -gt 0) { $existingPermissions[0] } else { $null }

        if ($permissionAction -eq 'Remove') {
            if (-not $existingPermission) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Remove" -Action 'SetMailboxFolderPermission' -Status 'Skipped' -Message 'Folder permission does not exist.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$folderIdentity -> $trusteeIdentity", 'Remove mailbox folder permission')) {
                Invoke-WithRetry -OperationName "Remove folder permission $folderIdentity -> $trusteeIdentity" -ScriptBlock {
                    Remove-MailboxFolderPermission -Identity $folderIdentity -User $trustee.Identity -Confirm:$false -ErrorAction Stop | Out-Null
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Remove" -Action 'SetMailboxFolderPermission' -Status 'Removed' -Message 'Folder permission removed successfully.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Remove" -Action 'SetMailboxFolderPermission' -Status 'WhatIf' -Message 'Folder permission removal skipped due to WhatIf.'))
            }

            $rowNumber++
            continue
        }

        $desiredAccessRightsCanonical = ConvertTo-CanonicalStringSet -Values $accessRights
        $desiredSharingFlagsCanonical = @()
        if ($isCalendarFolder -and $setSupportsSharingFlags) {
            if ($calendarDelegate) {
                $flags = [System.Collections.Generic.List[string]]::new()
                $flags.Add('Delegate')
                if ($calendarCanViewPrivateItems) {
                    $flags.Add('CanViewPrivateItems')
                }

                $desiredSharingFlagsCanonical = ConvertTo-CanonicalStringSet -Values $flags.ToArray()
            }
            else {
                $desiredSharingFlagsCanonical = @('none')
            }
        }

        if ($existingPermission) {
            if ($permissionAction -eq 'Add') {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Add" -Action 'SetMailboxFolderPermission' -Status 'Skipped' -Message 'Folder permission already exists. Use PermissionAction=Set to update rights.'))
                $rowNumber++
                continue
            }

            $currentAccessRightsCanonical = ConvertTo-CanonicalStringSet -Values @($existingPermission.AccessRights)
            $rightsMatch = Test-SameStringSet -Left $currentAccessRightsCanonical -Right $desiredAccessRightsCanonical

            $flagsMatch = $true
            if ($isCalendarFolder -and $setSupportsSharingFlags) {
                $currentFlagsCanonical = ConvertTo-CanonicalStringSet -Values @($existingPermission.SharingPermissionFlags)
                if ($currentFlagsCanonical.Count -eq 0) {
                    $currentFlagsCanonical = @('none')
                }

                $flagsMatch = Test-SameStringSet -Left $currentFlagsCanonical -Right $desiredSharingFlagsCanonical
            }

            if ($rightsMatch -and $flagsMatch) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Set" -Action 'SetMailboxFolderPermission' -Status 'Skipped' -Message 'Folder permissions already match requested values.'))
                $rowNumber++
                continue
            }

            if ($PSCmdlet.ShouldProcess("$folderIdentity -> $trusteeIdentity", 'Set mailbox folder permission')) {
                $params = @{
                    Identity     = $folderIdentity
                    User         = $trustee.Identity
                    AccessRights = $accessRights
                    ErrorAction  = 'Stop'
                }

                if ($isCalendarFolder -and $setSupportsSharingFlags) {
                    if ($desiredSharingFlagsCanonical.Count -eq 1 -and $desiredSharingFlagsCanonical[0] -eq 'none') {
                        $params.SharingPermissionFlags = @('None')
                    }
                    else {
                        $params.SharingPermissionFlags = $desiredSharingFlagsCanonical
                    }
                }

                if ($sendNotificationToUser) {
                    if ($setSupportsSendNotification) {
                        $params.SendNotificationToUser = $true
                    }
                    else {
                        Write-Status -Message "Set-MailboxFolderPermission in this session does not support -SendNotificationToUser. Notification request for '$folderIdentity' was ignored." -Level WARN
                    }
                }

                Invoke-WithRetry -OperationName "Set folder permission $folderIdentity -> $trusteeIdentity" -ScriptBlock {
                    Set-MailboxFolderPermission @params | Out-Null
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Set" -Action 'SetMailboxFolderPermission' -Status 'Updated' -Message 'Folder permission updated successfully.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Set" -Action 'SetMailboxFolderPermission' -Status 'WhatIf' -Message 'Folder permission update skipped due to WhatIf.'))
            }

            $rowNumber++
            continue
        }

        if ($permissionAction -eq 'Set') {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Set" -Action 'SetMailboxFolderPermission' -Status 'Skipped' -Message 'Folder permission does not exist. Use PermissionAction=Add to create it.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess("$folderIdentity -> $trusteeIdentity", 'Add mailbox folder permission')) {
            $params = @{
                Identity     = $folderIdentity
                User         = $trustee.Identity
                AccessRights = $accessRights
                ErrorAction  = 'Stop'
            }

            if ($isCalendarFolder -and $addSupportsSharingFlags) {
                if ($desiredSharingFlagsCanonical.Count -eq 1 -and $desiredSharingFlagsCanonical[0] -eq 'none') {
                    $params.SharingPermissionFlags = @('None')
                }
                else {
                    $params.SharingPermissionFlags = $desiredSharingFlagsCanonical
                }
            }

            if ($sendNotificationToUser) {
                if ($addSupportsSendNotification) {
                    $params.SendNotificationToUser = $true
                }
                else {
                    Write-Status -Message "Add-MailboxFolderPermission in this session does not support -SendNotificationToUser. Notification request for '$folderIdentity' was ignored." -Level WARN
                }
            }

            Invoke-WithRetry -OperationName "Add folder permission $folderIdentity -> $trusteeIdentity" -ScriptBlock {
                Add-MailboxFolderPermission @params | Out-Null
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Add" -Action 'SetMailboxFolderPermission' -Status 'Added' -Message 'Folder permission added successfully.'))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeIdentity|Add" -Action 'SetMailboxFolderPermission' -Status 'WhatIf' -Message 'Folder permission add skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity|$trusteeIdentity|$folderPathRaw) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeIdentity|$folderPathRaw" -Action 'SetMailboxFolderPermission' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange on-prem mailbox folder permission script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
