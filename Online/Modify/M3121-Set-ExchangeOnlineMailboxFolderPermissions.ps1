<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-004416

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_M3121-Set-ExchangeOnlineMailboxFolderPermissions_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
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
    'TrusteeUserPrincipalName',
    'FolderPath',
    'AccessRights',
    'CalendarDelegate',
    'CalendarCanViewPrivateItems',
    'SendNotificationToUser'
)

Write-Status -Message 'Starting Exchange Online mailbox folder permission assignment script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$addMailboxFolderPermissionCommand = Get-Command -Name Add-MailboxFolderPermission -ErrorAction Stop
$setMailboxFolderPermissionCommand = Get-Command -Name Set-MailboxFolderPermission -ErrorAction Stop
$addSupportsSharingFlags = $addMailboxFolderPermissionCommand.Parameters.ContainsKey('SharingPermissionFlags')
$setSupportsSharingFlags = $setMailboxFolderPermissionCommand.Parameters.ContainsKey('SharingPermissionFlags')
$addSupportsSendNotification = $addMailboxFolderPermissionCommand.Parameters.ContainsKey('SendNotificationToUser')
$setSupportsSendNotification = $setMailboxFolderPermissionCommand.Parameters.ContainsKey('SendNotificationToUser')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = ([string]$row.MailboxIdentity).Trim()
    $trusteeUpn = ([string]$row.TrusteeUserPrincipalName).Trim()
    $folderPathRaw = ([string]$row.FolderPath).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity) -or [string]::IsNullOrWhiteSpace($trusteeUpn) -or [string]::IsNullOrWhiteSpace($folderPathRaw)) {
            throw 'MailboxIdentity, TrusteeUserPrincipalName, and FolderPath are required.'
        }

        $accessRights = ConvertTo-Array -Value ([string]$row.AccessRights)
        if ($accessRights.Count -eq 0) {
            throw 'AccessRights is required and must contain at least one value.'
        }

        $calendarDelegate = ConvertTo-Bool -Value $row.CalendarDelegate -Default $false
        $calendarCanViewPrivateItems = ConvertTo-Bool -Value $row.CalendarCanViewPrivateItems -Default $false
        $sendNotificationToUser = ConvertTo-Bool -Value $row.SendNotificationToUser -Default $false

        if ($calendarCanViewPrivateItems -and -not $calendarDelegate) {
            throw 'CalendarCanViewPrivateItems requires CalendarDelegate to be TRUE.'
        }

        $folderIdentity = if ($folderPathRaw -match '^[^:]+:\\') {
            $folderPathRaw
        }
        else {
            "${mailboxIdentity}:\$folderPathRaw"
        }

        $folderPathForCheck = if ($folderPathRaw -match '^[^:]+:\\') {
            $folderPathRaw.Split(':\', 2)[1]
        }
        else {
            $folderPathRaw
        }
        $folderSegments = @($folderPathForCheck -split '[\\/]' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $folderLeaf = ''
        if ($folderSegments.Count -gt 0) {
            $folderLeaf = ([string]$folderSegments[-1]).Trim()
        }
        $isCalendarFolder = $folderLeaf.Equals('Calendar', [System.StringComparison]::OrdinalIgnoreCase)

        if (($calendarDelegate -or $calendarCanViewPrivateItems -or $sendNotificationToUser) -and -not $isCalendarFolder) {
            throw 'CalendarDelegate, CalendarCanViewPrivateItems, and SendNotificationToUser are only supported when FolderPath targets the Calendar folder.'
        }

        if ($isCalendarFolder -and ($calendarDelegate -or $calendarCanViewPrivateItems) -and (-not $addSupportsSharingFlags -or -not $setSupportsSharingFlags)) {
            throw 'This Exchange Online cmdlet session does not support -SharingPermissionFlags required for calendar delegate configuration.'
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentity" -ScriptBlock {
            Get-Mailbox -Identity $mailboxIdentity -ErrorAction Stop
        }
        $trustee = Invoke-WithRetry -OperationName "Lookup trustee $trusteeUpn" -ScriptBlock {
            Get-Recipient -Identity $trusteeUpn -ErrorAction Stop
        }

        $existingPermissions = @(Invoke-WithRetry -OperationName "Check folder permission $folderIdentity -> $trusteeUpn" -ScriptBlock {
            Get-MailboxFolderPermission -Identity $folderIdentity -User $trustee.Identity -ErrorAction SilentlyContinue
        })
        $existingPermission = if ($existingPermissions.Count -gt 0) { $existingPermissions[0] } else { $null }

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
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeUpn" -Action 'SetMailboxFolderPermission' -Status 'Skipped' -Message 'Folder permissions already match the requested configuration.'))
                $rowNumber++
                continue
            }
        }

        if ($PSCmdlet.ShouldProcess("$folderIdentity -> $trusteeUpn", 'Set mailbox folder permissions')) {
            if ($existingPermission) {
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

                Invoke-WithRetry -OperationName "Update folder permission $folderIdentity -> $trusteeUpn" -ScriptBlock {
                    Set-MailboxFolderPermission @params
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeUpn" -Action 'SetMailboxFolderPermission' -Status 'Updated' -Message 'Folder permissions updated successfully.'))
            }
            else {
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

                Invoke-WithRetry -OperationName "Add folder permission $folderIdentity -> $trusteeUpn" -ScriptBlock {
                    Add-MailboxFolderPermission @params
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeUpn" -Action 'SetMailboxFolderPermission' -Status 'Added' -Message 'Folder permissions added successfully.'))
            }
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$folderIdentity|$trusteeUpn" -Action 'SetMailboxFolderPermission' -Status 'WhatIf' -Message 'Folder permission update skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity|$trusteeUpn|$folderPathRaw) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey "$mailboxIdentity|$trusteeUpn|$folderPathRaw" -Action 'SetMailboxFolderPermission' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mailbox folder permission assignment script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}









