# Modify Detailed Catalog

Detailed catalog for modify/change scripts in `SharedModule/Online/Modify/`.

Current implementation status: foundational update scripts are implemented across Entra, Teams, Exchange Online, and SharePoint/OneDrive, including Entra group-creator policy control, accepted-domain management, and the planned OneDrive/Teams update wave.

## Script Contract

All update scripts should:

- Run on PowerShell 7+
- Support `-WhatIf` and `ShouldProcess`
- Be idempotent when practical
- Validate CSV headers and required values
- Export per-record `Status` and `Message`
- Write a required per-run transcript log in the output folder
- Include rollback/remediation notes for high-impact operations

## ID Ranges

- `30xx`: Entra
- `31xx`: Exchange Online
- `32xx`: SharePoint/OneDrive
- `33xx`: Teams

## Catalog

| ID | Script | Input Template | Workload | Primary Change Scope | Status |
|---|---|---|---|---|---|
| M3001 | `SM-M3001-Update-EntraUsers.ps1` | `SM-M3001-Update-EntraUsers.input.csv` | Entra | User profile/attribute updates with fold-in password reset support | Implemented |
| M3002 | `SM-M3002-Set-EntraUserAccountState.ps1` | `SM-M3002-Set-EntraUserAccountState.input.csv` | Entra | Enable/disable accounts | Implemented |
| M3003 | `SM-M3003-Set-EntraUserLicenses.ps1` | `SM-M3003-Set-EntraUserLicenses.input.csv` | Entra | Add/update license assignments | Implemented |
| M3004 | `SM-M3004-Set-EntraUserPasswordResets.ps1` | `SM-M3004-Set-EntraUserPasswordResets.input.csv` | Entra | Standalone password reset workflow with force-change controls | Implemented |
| M3204 | `SM-M3204-PreProvision-OneDrive.ps1` | `SM-M3204-PreProvision-OneDrive.input.csv` | OneDrive/SharePoint | Trigger OneDrive pre-provisioning for existing users | Implemented |
| M3205 | `SM-M3205-Set-OneDriveStorageQuota.ps1` | `SM-M3205-Set-OneDriveStorageQuota.input.csv` | OneDrive/SharePoint | Update OneDrive storage quota and warning level by user | Implemented |
| M3206 | `SM-M3206-Set-OneDriveSiteCollectionAdmins.ps1` | `SM-M3206-Set-OneDriveSiteCollectionAdmins.input.csv` | OneDrive/SharePoint | Add/remove OneDrive site collection administrators by user | Implemented |
| M3207 | `SM-M3207-Set-OneDriveSharingSettings.ps1` | `SM-M3207-Set-OneDriveSharingSettings.input.csv` | OneDrive/SharePoint | Update OneDrive sharing policy posture by user/site | Implemented |
| M3208 | `SM-M3208-Revoke-OneDriveExternalSharingLinks.ps1` | `SM-M3208-Revoke-OneDriveExternalSharingLinks.input.csv` | OneDrive/SharePoint | Revoke scoped OneDrive external links/principal shares | Implemented |
| M3209 | `SM-M3209-Set-OneDriveSiteLockState.ps1` | `SM-M3209-Set-OneDriveSiteLockState.input.csv` | OneDrive/SharePoint | Set OneDrive lock state for lifecycle/incident controls | Implemented |
| M3005 | `SM-M3005-Update-EntraAssignedSecurityGroups.ps1` | `SM-M3005-Update-EntraAssignedSecurityGroups.input.csv` | Entra | Assigned security group property updates | Implemented |
| M3006 | `SM-M3006-Update-EntraDynamicUserSecurityGroups.ps1` | `SM-M3006-Update-EntraDynamicUserSecurityGroups.input.csv` | Entra | Dynamic rule/processing updates | Implemented |
| M3007 | `SM-M3007-Set-EntraSecurityGroupMembers.ps1` | `SM-M3007-Set-EntraSecurityGroupMembers.input.csv` | Entra | Add/remove group members | Implemented |
| M3008 | `SM-M3008-Update-EntraMicrosoft365Groups.ps1` | `SM-M3008-Update-EntraMicrosoft365Groups.input.csv` | Entra | M365 group settings/visibility updates | Implemented |
| M3009 | `SM-M3009-Set-EntraGroupCreatorsPolicy.ps1` | `SM-M3009-Set-EntraGroupCreatorsPolicy.input.csv` | Entra | Configure tenant-wide Microsoft 365 group creators policy and allowed-creator group | Implemented |
| M3309 | `SM-M3309-Update-MicrosoftTeams.ps1` | `SM-M3309-Update-MicrosoftTeams.input.csv` | Teams | Team settings updates | Implemented |
| M3310 | `SM-M3310-Set-MicrosoftTeamMembers.ps1` | `SM-M3310-Set-MicrosoftTeamMembers.input.csv` | Teams | Add Team owners/members | Implemented |
| M3311 | `SM-M3311-Update-MicrosoftTeamChannels.ps1` | `SM-M3311-Update-MicrosoftTeamChannels.input.csv` | Teams | Add channels to existing Teams | Implemented |
| M3312 | `SM-M3312-Set-MicrosoftTeamChannelMembers.ps1` | `SM-M3312-Set-MicrosoftTeamChannelMembers.input.csv` | Teams | Add/update private/shared channel members | Implemented |
| M3113 | `SM-M3113-Update-ExchangeOnlineMailContacts.ps1` | `SM-M3113-Update-ExchangeOnlineMailContacts.input.csv` | Exchange Online | Mail contact property updates | Implemented |
| M3114 | `SM-M3114-Update-ExchangeOnlineDistributionLists.ps1` | `SM-M3114-Update-ExchangeOnlineDistributionLists.input.csv` | Exchange Online | Distribution list property updates | Implemented |
| M3115 | `SM-M3115-Set-ExchangeOnlineDistributionListMembers.ps1` | `SM-M3115-Set-ExchangeOnlineDistributionListMembers.input.csv` | Exchange Online | Add/remove distribution list members | Implemented |
| M3116 | `SM-M3116-Update-ExchangeOnlineSharedMailboxes.ps1` | `SM-M3116-Update-ExchangeOnlineSharedMailboxes.input.csv` | Exchange Online | Shared mailbox property updates | Implemented |
| M3117 | `SM-M3117-Set-ExchangeOnlineSharedMailboxPermissions.ps1` | `SM-M3117-Set-ExchangeOnlineSharedMailboxPermissions.input.csv` | Exchange Online | Configure shared mailbox permissions | Implemented |
| M3118 | `SM-M3118-Update-ExchangeOnlineResourceMailboxes.ps1` | `SM-M3118-Update-ExchangeOnlineResourceMailboxes.input.csv` | Exchange Online | Resource mailbox settings updates | Implemented |
| M3119 | `SM-M3119-Set-ExchangeOnlineResourceMailboxBookingDelegates.ps1` | `SM-M3119-Set-ExchangeOnlineResourceMailboxBookingDelegates.input.csv` | Exchange Online | Configure booking delegates/policy flags | Implemented |
| M3120 | `SM-M3120-Set-ExchangeOnlineMailboxDelegations.ps1` | `SM-M3120-Set-ExchangeOnlineMailboxDelegations.input.csv` | Exchange Online | Configure mailbox delegation rights | Implemented |
| M3121 | `SM-M3121-Set-ExchangeOnlineMailboxFolderPermissions.ps1` | `SM-M3121-Set-ExchangeOnlineMailboxFolderPermissions.input.csv` | Exchange Online | Configure folder permissions/delegate flags | Implemented |
| M3122 | `SM-M3122-Update-ExchangeOnlineMailEnabledSecurityGroups.ps1` | `SM-M3122-Update-ExchangeOnlineMailEnabledSecurityGroups.input.csv` | Exchange Online | Mail-enabled security group property updates | Implemented |
| M3123 | `SM-M3123-Update-ExchangeOnlineDynamicDistributionGroups.ps1` | `SM-M3123-Update-ExchangeOnlineDynamicDistributionGroups.input.csv` | Exchange Online | Dynamic distribution group property/filter updates | Implemented |
| M3124 | `SM-M3124-Set-ExchangeOnlineUserMailboxForwarding.ps1` | `SM-M3124-Set-ExchangeOnlineUserMailboxForwarding.input.csv` | Exchange Online | Set per-user mailbox forwarding mode and delivery behavior | Implemented |
| M3125 | `SM-M3125-Update-ExchangeOnlineProxyAddresses.ps1` | `SM-M3125-Update-ExchangeOnlineProxyAddresses.input.csv` | Exchange Online | Add/remove/replace mailbox proxy addresses with clear-secondary support | Implemented |
| M3126 | `SM-M3126-Set-ExchangeOnlineMailboxAllowedAuthenticationMethods.ps1` | `SM-M3126-Set-ExchangeOnlineMailboxAllowedAuthenticationMethods.input.csv` | Exchange Online | Configure mailbox authentication/protocol access profile through CAS settings | Implemented |
| M3127 | `SM-M3127-Set-ExchangeOnlineUserPhotos.ps1` | `SM-M3127-Set-ExchangeOnlineUserPhotos.input.csv` | Exchange Online | Set/remove user mailbox photos from file path inputs | Implemented |
| M3128 | `SM-M3128-Set-ExchangeOnlineSafeSenderDomains.ps1` | `SM-M3128-Set-ExchangeOnlineSafeSenderDomains.input.csv` | Exchange Online | Add/remove/replace mailbox trusted sender/domain lists | Implemented |
| M3129 | `SM-M3129-Restore-ExchangeOnlineRecoverableItems.ps1` | `SM-M3129-Restore-ExchangeOnlineRecoverableItems.input.csv` | Exchange Online | Restore recoverable mailbox items in controlled batches | Implemented |
| M3130 | `SM-M3130-Verify-ExchangeOnlineAcceptedDomains.ps1` | `SM-M3130-Verify-ExchangeOnlineAcceptedDomains.input.csv` | Exchange Online | Confirm domain verification state and accepted-domain readiness | Implemented |
| M3131 | `SM-M3131-Remove-ExchangeOnlineAcceptedDomains.ps1` | `SM-M3131-Remove-ExchangeOnlineAcceptedDomains.input.csv` | Exchange Online | Remove accepted domains and optional Entra tenant domains | Implemented |
| M3241 | `SM-M3241-Set-SharePointSiteAdmins.ps1` | `SM-M3241-Set-SharePointSiteAdmins.input.csv` | SharePoint | Add/remove site collection administrators with last-admin safety guard | Implemented |
| M3243 | `SM-M3243-Associate-SharePointSitesToHub.ps1` | `SM-M3243-Associate-SharePointSitesToHub.input.csv` | SharePoint | Associate existing sites to hub sites (optional reassociation) | Implemented |

## Safety and Sequencing Guidance

Recommended execution phases:

1. Entra and OneDrive changes: `M3001`, `M3002`, `M3003`, `M3004`, `M3005`, `M3006`, `M3007`, `M3008`, `M3009`, `M3204` to `M3209`
2. Teams changes: `M3309` to `M3312`
3. Exchange Online changes: `M3113` to `M3131` (implemented Exchange Online set/update/domain coverage)
4. SharePoint changes: `M3241`, `M3243`

High-impact change classes that should always start with `-WhatIf`:

- Membership removals (`M3007`, `M3115`, and future remove-capable variants of `M3310`, `M3312`)
- Permission changes (`M3117`, `M3120`, `M3121`)
- Recoverable item restore operations (`M3129`)
- Domain removal/change operations (`M3131`, plus `M3130` when `AttemptVerification` is enabled)
- SharePoint admin/hub association changes (`M3241`, `M3243`)
- Account state and password changes (`M3002`, `M3004`)
- Tenant-wide policy changes (`M3009`)

## Standard Result Columns

Recommended baseline columns:

- `RowNumber`
- `PrimaryKey`
- `Action`
- `Status`
- `Message`

## Related Docs

- [Modify README](./README.md)
- [SharedModule README](../../README.md)
- [Provision Detailed Catalog](../Provision/README-Provision-Catalog.md)

