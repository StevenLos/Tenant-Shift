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
Microsoft.Graph.Authentication
Microsoft.Graph.Users
Microsoft.Graph.Identity.SignIns

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Provisions EntraGuestUsers in Microsoft 365.

.DESCRIPTION
    Creates EntraGuestUsers in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P3002-Invite-EntraGuestUsers.ps1 -InputCsvPath .\3002.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P3002-Invite-EntraGuestUsers.ps1 -InputCsvPath .\3002.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Identity.SignIns
    Required roles:   Global Administrator or appropriate workload-specific role
    Limitations:      None known.

    CSV Fields:
    Column                   Type      Required  Description
    -----------------------  ----      --------  -----------
    InvitedUserEmailAddress  String    Yes       <fill in description>
    InvitedUserDisplayName   String    Yes       <fill in description>
    InviteRedirectUrl        String    Yes       <fill in description>
    SendInvitationMessage    String    Yes       <fill in description>
    InvitedUserType          String    Yes       <fill in description>
    CustomMessageBody        String    Yes       <fill in description>
    CcRecipients             String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P3002-Invite-EntraGuestUsers_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {


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
}
finally {
    Stop-RunTranscript
}







