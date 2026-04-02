# Modify Detailed Catalog

Detailed catalog for modify/change scripts in `SharedModule/Online/Modify/`.

Current implementation status: foundational update scripts are implemented across MEID, TEAM, EXOL, ONDR, and SPOL, including Entra group-creator policy control, accepted-domain management, and OneDrive/Teams update coverage.

## Script Contract

All modify scripts should:

- Run on PowerShell 7+
- Support `-WhatIf` and `ShouldProcess`
- Be idempotent when practical
- Validate CSV headers and required values
- Export per-record `Status` and `Message`
- Write a required per-run transcript log in the output folder
- Include rollback/remediation notes for high-impact operations

## Catalog

| ID | Script | Input Template | Workload | Primary Change Scope | Status |
|---|---|---|---|---|---|
| M-MEID-0010 | `M-MEID-0010-Update-EntraUsers.ps1` | `M-MEID-0010-Update-EntraUsers.input.csv` | MEID | User profile/attribute updates with fold-in password reset support | Implemented |
| M-MEID-0030 | `M-MEID-0030-Set-EntraUserLicenses.ps1` | `M-MEID-0030-Set-EntraUserLicenses.input.csv` | MEID | Add/update license assignments | Implemented |
| M-MEID-0040 | `M-MEID-0040-Set-EntraUserAccountState.ps1` | `M-MEID-0040-Set-EntraUserAccountState.input.csv` | MEID | Enable/disable accounts | Implemented |
| M-MEID-0050 | `M-MEID-0050-Set-EntraUserPasswordResets.ps1` | `M-MEID-0050-Set-EntraUserPasswordResets.input.csv` | MEID | Standalone password reset workflow with force-change controls | Implemented |
| M-MEID-0070 | `M-MEID-0070-Update-EntraAssignedSecurityGroups.ps1` | `M-MEID-0070-Update-EntraAssignedSecurityGroups.input.csv` | MEID | Assigned security group property updates | Implemented |
| M-MEID-0080 | `M-MEID-0080-Update-EntraDynamicUserSecurityGroups.ps1` | `M-MEID-0080-Update-EntraDynamicUserSecurityGroups.input.csv` | MEID | Dynamic rule/processing updates | Implemented |
| M-MEID-0090 | `M-MEID-0090-Set-EntraSecurityGroupMembers.ps1` | `M-MEID-0090-Set-EntraSecurityGroupMembers.input.csv` | MEID | Add/remove group members | Implemented |
| M-MEID-0100 | `M-MEID-0100-Update-EntraMicrosoft365Groups.ps1` | `M-MEID-0100-Update-EntraMicrosoft365Groups.input.csv` | MEID | M365 group settings/visibility updates | Implemented |
| M-MEID-0130 | `M-MEID-0130-Set-EntraGroupCreatorsPolicy.ps1` | `M-MEID-0130-Set-EntraGroupCreatorsPolicy.input.csv` | MEID | Configure tenant-wide Microsoft 365 group creators policy and allowed-creator group | Implemented |
| M-EXOL-0010 | `M-EXOL-0010-Verify-ExchangeOnlineAcceptedDomains.ps1` | `M-EXOL-0010-Verify-ExchangeOnlineAcceptedDomains.input.csv` | EXOL | Confirm domain verification state and accepted-domain readiness | Implemented |
| M-EXOL-0020 | `M-EXOL-0020-Remove-ExchangeOnlineAcceptedDomains.ps1` | `M-EXOL-0020-Remove-ExchangeOnlineAcceptedDomains.input.csv` | EXOL | Remove accepted domains and optional Entra tenant domains | Implemented |
| M-EXOL-0030 | `M-EXOL-0030-Update-ExchangeOnlineMailContacts.ps1` | `M-EXOL-0030-Update-ExchangeOnlineMailContacts.input.csv` | EXOL | Mail contact property updates | Implemented |
| M-EXOL-0040 | `M-EXOL-0040-Update-ExchangeOnlineDistributionLists.ps1` | `M-EXOL-0040-Update-ExchangeOnlineDistributionLists.input.csv` | EXOL | Distribution list property updates | Implemented |
| M-EXOL-0050 | `M-EXOL-0050-Update-ExchangeOnlineMailEnabledSecurityGroups.ps1` | `M-EXOL-0050-Update-ExchangeOnlineMailEnabledSecurityGroups.input.csv` | EXOL | Mail-enabled security group property updates | Implemented |
| M-EXOL-0060 | `M-EXOL-0060-Update-ExchangeOnlineDynamicDistributionGroups.ps1` | `M-EXOL-0060-Update-ExchangeOnlineDynamicDistributionGroups.input.csv` | EXOL | Dynamic distribution group property/filter updates | Implemented |
| M-EXOL-0070 | `M-EXOL-0070-Update-ExchangeOnlineSharedMailboxes.ps1` | `M-EXOL-0070-Update-ExchangeOnlineSharedMailboxes.input.csv` | EXOL | Shared mailbox property updates | Implemented |
| M-EXOL-0080 | `M-EXOL-0080-Update-ExchangeOnlineResourceMailboxes.ps1` | `M-EXOL-0080-Update-ExchangeOnlineResourceMailboxes.input.csv` | EXOL | Resource mailbox settings updates | Implemented |
| M-EXOL-0090 | `M-EXOL-0090-Set-ExchangeOnlineDistributionListMembers.ps1` | `M-EXOL-0090-Set-ExchangeOnlineDistributionListMembers.input.csv` | EXOL | Add/remove distribution list members | Implemented |
| M-EXOL-0110 | `M-EXOL-0110-Set-ExchangeOnlineSharedMailboxPermissions.ps1` | `M-EXOL-0110-Set-ExchangeOnlineSharedMailboxPermissions.input.csv` | EXOL | Configure shared mailbox permissions | Implemented |
| M-EXOL-0120 | `M-EXOL-0120-Set-ExchangeOnlineResourceMailboxBookingDelegates.ps1` | `M-EXOL-0120-Set-ExchangeOnlineResourceMailboxBookingDelegates.input.csv` | EXOL | Configure booking delegates/policy flags | Implemented |
| M-EXOL-0130 | `M-EXOL-0130-Set-ExchangeOnlineMailboxDelegations.ps1` | `M-EXOL-0130-Set-ExchangeOnlineMailboxDelegations.input.csv` | EXOL | Configure mailbox delegation rights | Implemented |
| M-EXOL-0140 | `M-EXOL-0140-Set-ExchangeOnlineMailboxFolderPermissions.ps1` | `M-EXOL-0140-Set-ExchangeOnlineMailboxFolderPermissions.input.csv` | EXOL | Configure folder permissions/delegate flags | Implemented |
| M-EXOL-0150 | `M-EXOL-0150-Set-ExchangeOnlineUserMailboxForwarding.ps1` | `M-EXOL-0150-Set-ExchangeOnlineUserMailboxForwarding.input.csv` | EXOL | Set per-user mailbox forwarding mode and delivery behavior | Implemented |
| M-EXOL-0160 | `M-EXOL-0160-Update-ExchangeOnlineProxyAddresses.ps1` | `M-EXOL-0160-Update-ExchangeOnlineProxyAddresses.input.csv` | EXOL | Add/remove/replace mailbox proxy addresses with clear-secondary support | Implemented |
| M-EXOL-0170 | `M-EXOL-0170-Set-ExchangeOnlineMailboxAllowedAuthenticationMethods.ps1` | `M-EXOL-0170-Set-ExchangeOnlineMailboxAllowedAuthenticationMethods.input.csv` | EXOL | Configure mailbox authentication/protocol access profile through CAS settings | Implemented |
| M-EXOL-0180 | `M-EXOL-0180-Set-ExchangeOnlineUserPhotos.ps1` | `M-EXOL-0180-Set-ExchangeOnlineUserPhotos.input.csv` | EXOL | Set/remove user mailbox photos from file path inputs | Implemented |
| M-EXOL-0190 | `M-EXOL-0190-Set-ExchangeOnlineSafeSenderDomains.ps1` | `M-EXOL-0190-Set-ExchangeOnlineSafeSenderDomains.input.csv` | EXOL | Add/remove/replace mailbox trusted sender/domain lists | Implemented |
| M-EXOL-0200 | `M-EXOL-0200-Restore-ExchangeOnlineRecoverableItems.ps1` | `M-EXOL-0200-Restore-ExchangeOnlineRecoverableItems.input.csv` | EXOL | Restore recoverable mailbox items in controlled batches | Implemented |
| M-ONDR-0010 | `M-ONDR-0010-PreProvision-OneDrive.ps1` | `M-ONDR-0010-PreProvision-OneDrive.input.csv` | ONDR | Trigger OneDrive pre-provisioning for existing users | Implemented |
| M-ONDR-0020 | `M-ONDR-0020-Set-OneDriveStorageQuota.ps1` | `M-ONDR-0020-Set-OneDriveStorageQuota.input.csv` | ONDR | Update OneDrive storage quota and warning level by user | Implemented |
| M-ONDR-0030 | `M-ONDR-0030-Set-OneDriveSiteCollectionAdmins.ps1` | `M-ONDR-0030-Set-OneDriveSiteCollectionAdmins.input.csv` | ONDR | Add/remove OneDrive site collection administrators by user | Implemented |
| M-ONDR-0040 | `M-ONDR-0040-Set-OneDriveSharingSettings.ps1` | `M-ONDR-0040-Set-OneDriveSharingSettings.input.csv` | ONDR | Update OneDrive sharing policy posture by user/site | Implemented |
| M-ONDR-0050 | `M-ONDR-0050-Revoke-OneDriveExternalSharingLinks.ps1` | `M-ONDR-0050-Revoke-OneDriveExternalSharingLinks.input.csv` | ONDR | Revoke scoped OneDrive external links/principal shares | Implemented |
| M-ONDR-0060 | `M-ONDR-0060-Set-OneDriveSiteLockState.ps1` | `M-ONDR-0060-Set-OneDriveSiteLockState.input.csv` | ONDR | Set OneDrive lock state for lifecycle/incident controls | Implemented |
| M-SPOL-0030 | `M-SPOL-0030-Set-SharePointSiteAdmins.ps1` | `M-SPOL-0030-Set-SharePointSiteAdmins.input.csv` | SPOL | Add/remove site collection administrators with last-admin safety guard | Implemented |
| M-SPOL-0040 | `M-SPOL-0040-Associate-SharePointSitesToHub.ps1` | `M-SPOL-0040-Associate-SharePointSitesToHub.input.csv` | SPOL | Associate existing sites to hub sites (optional reassociation) | Implemented |
| M-TEAM-0010 | `M-TEAM-0010-Update-MicrosoftTeams.ps1` | `M-TEAM-0010-Update-MicrosoftTeams.input.csv` | TEAM | Team settings updates | Implemented |
| M-TEAM-0020 | `M-TEAM-0020-Set-MicrosoftTeamMembers.ps1` | `M-TEAM-0020-Set-MicrosoftTeamMembers.input.csv` | TEAM | Add Team owners/members | Implemented |
| M-TEAM-0030 | `M-TEAM-0030-Update-MicrosoftTeamChannels.ps1` | `M-TEAM-0030-Update-MicrosoftTeamChannels.input.csv` | TEAM | Add channels to existing Teams | Implemented |
| M-TEAM-0040 | `M-TEAM-0040-Set-MicrosoftTeamChannelMembers.ps1` | `M-TEAM-0040-Set-MicrosoftTeamChannelMembers.input.csv` | TEAM | Add/update private/shared channel members | Implemented |

## Safety and Sequencing Guidance

Recommended execution phases:

1. MEID changes: `M-MEID-0010`, `M-MEID-0030`, `M-MEID-0040`, `M-MEID-0050`, `M-MEID-0070`, `M-MEID-0080`, `M-MEID-0090`, `M-MEID-0100`, `M-MEID-0130`
2. TEAM changes: `M-TEAM-0010`, `M-TEAM-0020`, `M-TEAM-0030`, `M-TEAM-0040`
3. EXOL changes: `M-EXOL-0010` through `M-EXOL-0200`
4. ONDR changes: `M-ONDR-0010` through `M-ONDR-0060`
5. SPOL changes: `M-SPOL-0030`, `M-SPOL-0040`

High-impact change classes that should always start with `-WhatIf`:

- Membership removals (`M-MEID-0090`, `M-EXOL-0090`)
- Permission changes (`M-EXOL-0110`, `M-EXOL-0130`, `M-EXOL-0140`)
- Recoverable item restore operations (`M-EXOL-0200`)
- Domain removal/change operations (`M-EXOL-0020`, plus `M-EXOL-0010` when `AttemptVerification` is enabled)
- SharePoint admin/hub association changes (`M-SPOL-0030`, `M-SPOL-0040`)
- Account state and password changes (`M-MEID-0040`, `M-MEID-0050`)
- Tenant-wide policy changes (`M-MEID-0130`)

## Standard Result Columns

Recommended baseline columns:

- `RowNumber`
- `PrimaryKey`
- `Action`
- `Status`
- `Message`

## Related Docs

- [Modify README](./README.md)
- [Root README](../../README.md)
- [Provision Detailed Catalog](../Provision/README-Provision-Catalog.md)
