#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath ("Results_B02-Invite-EntraGuestUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'M365.Common.psm1') -Force -DisableNameChecking

$requiredHeaders = @(
    'InvitedUserEmailAddress',
    'InvitedUserDisplayName',
    'InviteRedirectUrl',
    'SendInvitationMessage',
    'InvitedUserType',
    'CustomMessageBody',
    'CcRecipients'
)

Write-Status -Message 'Starting Entra ID guest invitation script.'
Assert-ModuleCurrent -ModuleNames @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Identity.SignIns'
)
Ensure-GraphConnection -RequiredScopes @('User.Read.All', 'User.Invite.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$defaultRedirectUrl = 'https://myapplications.microsoft.com'

$rowNumber = 1
foreach ($row in $rows) {
    $invitedUserEmailAddress = ([string]$row.InvitedUserEmailAddress).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($invitedUserEmailAddress)) {
            throw 'InvitedUserEmailAddress is required.'
        }

        $invitedUserDisplayName = ([string]$row.InvitedUserDisplayName).Trim()
        $inviteRedirectUrl = ([string]$row.InviteRedirectUrl).Trim()
        if ([string]::IsNullOrWhiteSpace($inviteRedirectUrl)) {
            $inviteRedirectUrl = $defaultRedirectUrl
        }

        $sendInvitationMessage = ConvertTo-Bool -Value $row.SendInvitationMessage -Default $true

        $invitedUserType = ([string]$row.InvitedUserType).Trim()
        if ([string]::IsNullOrWhiteSpace($invitedUserType)) {
            $invitedUserType = 'Guest'
        }
        elseif ($invitedUserType -notin @('Guest', 'Member')) {
            throw "InvitedUserType '$invitedUserType' is invalid. Use Guest or Member."
        }

        $escapedInviteEmail = Escape-ODataString -Value $invitedUserEmailAddress
        $existingByMail = @(Invoke-WithRetry -OperationName "Lookup user by mail $invitedUserEmailAddress" -ScriptBlock {
            Get-MgUser -Filter "mail eq '$escapedInviteEmail'" -ConsistencyLevel eventual -ErrorAction Stop
        })

        $existingByOtherMails = @()
        try {
            $existingByOtherMails = @(Invoke-WithRetry -OperationName "Lookup user by otherMails $invitedUserEmailAddress" -ScriptBlock {
                Get-MgUser -Filter "otherMails/any(c:c eq '$escapedInviteEmail')" -ConsistencyLevel eventual -ErrorAction Stop
            })
        }
        catch {
            Write-Status -Message "Could not query otherMails for '$invitedUserEmailAddress'. Continuing with mail-only match. Error: $($_.Exception.Message)" -Level WARN
        }

        $existingUsersById = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($candidate in @($existingByMail + $existingByOtherMails)) {
            $candidateId = ([string]$candidate.Id).Trim()
            if ([string]::IsNullOrWhiteSpace($candidateId)) {
                continue
            }

            if (-not $existingUsersById.ContainsKey($candidateId)) {
                $existingUsersById[$candidateId] = $candidate
            }
        }

        $existingUsers = @($existingUsersById.Values)
        if ($existingUsers.Count -gt 1) {
            $matchingIds = @(
                $existingUsers |
                    ForEach-Object { ([string]$_.Id).Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Sort-Object -Unique
            )
            throw "Multiple users match '$invitedUserEmailAddress'. Matching object IDs: $($matchingIds -join ', ')"
        }

        if ($existingUsers.Count -eq 1) {
            $existingUser = $existingUsers[0]
            $existingUserType = ([string]$existingUser.UserType).Trim()
            $existingUpn = ([string]$existingUser.UserPrincipalName).Trim()

            if ($existingUserType.Equals('Guest', [System.StringComparison]::OrdinalIgnoreCase)) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $invitedUserEmailAddress -Action 'InviteGuestUser' -Status 'Skipped' -Message "Guest already exists (UPN '$existingUpn')."))
                $rowNumber++
                continue
            }

            $userTypeForMessage = if ([string]::IsNullOrWhiteSpace($existingUserType)) { '<empty>' } else { $existingUserType }
            throw "A non-guest user already exists for '$invitedUserEmailAddress' (UPN '$existingUpn', userType '$userTypeForMessage')."
        }

        $invitationBody = @{
            invitedUserEmailAddress = $invitedUserEmailAddress
            inviteRedirectUrl       = $inviteRedirectUrl
            sendInvitationMessage   = $sendInvitationMessage
            invitedUserType         = $invitedUserType
        }

        if (-not [string]::IsNullOrWhiteSpace($invitedUserDisplayName)) {
            $invitationBody.invitedUserDisplayName = $invitedUserDisplayName
        }

        $customMessageBody = ([string]$row.CustomMessageBody).Trim()
        $ccRecipients = ConvertTo-Array -Value ([string]$row.CcRecipients)

        if (-not [string]::IsNullOrWhiteSpace($customMessageBody) -or $ccRecipients.Count -gt 0) {
            $invitedUserMessageInfo = @{}

            if (-not [string]::IsNullOrWhiteSpace($customMessageBody)) {
                $invitedUserMessageInfo.customizedMessageBody = $customMessageBody
            }

            if ($ccRecipients.Count -gt 0) {
                $invitedUserMessageInfo.ccRecipients = @(
                    foreach ($ccRecipient in $ccRecipients) {
                        @{ emailAddress = @{ address = $ccRecipient } }
                    }
                )
            }

            $invitationBody.invitedUserMessageInfo = $invitedUserMessageInfo
        }

        if ($PSCmdlet.ShouldProcess($invitedUserEmailAddress, 'Invite Entra ID guest user')) {
            $invitation = Invoke-WithRetry -OperationName "Invite guest $invitedUserEmailAddress" -ScriptBlock {
                New-MgInvitation -BodyParameter $invitationBody -ErrorAction Stop
            }

            $invitedUserId = ''
            if ($invitation.PSObject.Properties.Name -contains 'InvitedUser') {
                if ($invitation.InvitedUser -and $invitation.InvitedUser.PSObject.Properties.Name -contains 'Id') {
                    $invitedUserId = ([string]$invitation.InvitedUser.Id).Trim()
                }
            }

            if ([string]::IsNullOrWhiteSpace($invitedUserId) -and $invitation.PSObject.Properties.Name -contains 'InvitedUserId') {
                $invitedUserId = ([string]$invitation.InvitedUserId).Trim()
            }

            $inviteRedeemUrl = ''
            if ($invitation.PSObject.Properties.Name -contains 'InviteRedeemUrl') {
                $inviteRedeemUrl = ([string]$invitation.InviteRedeemUrl).Trim()
            }

            $successMessage = 'Guest invitation created successfully.'
            if (-not [string]::IsNullOrWhiteSpace($invitedUserId)) {
                $successMessage = "$successMessage InvitedUserId: $invitedUserId."
            }
            if (-not [string]::IsNullOrWhiteSpace($inviteRedeemUrl)) {
                $successMessage = "$successMessage InviteRedeemUrl returned."
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $invitedUserEmailAddress -Action 'InviteGuestUser' -Status 'Invited' -Message $successMessage))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $invitedUserEmailAddress -Action 'InviteGuestUser' -Status 'WhatIf' -Message 'Invitation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($invitedUserEmailAddress) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $invitedUserEmailAddress -Action 'InviteGuestUser' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Entra ID guest invitation script completed.' -Level SUCCESS

