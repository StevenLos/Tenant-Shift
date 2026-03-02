# Modify Detailed Catalog

Detailed catalog for modify/change scripts in `Online/Modify/`.

Current implementation status: foundational update scripts are implemented across Entra, Teams, Exchange Online, and SharePoint.

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
| M3001 | `M3001-Update-EntraUsers.ps1` | `M3001-Update-EntraUsers.input.csv` | Entra | User profile/attribute updates | Planned |
| M3002 | `M3002-Set-EntraUserAccountState.ps1` | `M3002-Set-EntraUserAccountState.input.csv` | Entra | Enable/disable accounts | Planned |
| M3003 | `M3003-Set-EntraUserLicenses.ps1` | `M3003-Set-EntraUserLicenses.input.csv` | Entra | Add/update license assignments | Implemented |
| M3204 | `M3204-PreProvision-OneDrive.ps1` | `M3204-PreProvision-OneDrive.input.csv` | OneDrive/SharePoint | Trigger OneDrive pre-provisioning for existing users | Implemented |
| M3205 | `M3205-Set-OneDriveStorageQuota.ps1` | `M3205-Set-OneDriveStorageQuota.input.csv` | OneDrive/SharePoint | Update OneDrive storage quota and warning level by user | Implemented |
| M3206 | `M3206-Set-OneDriveSiteCollectionAdmins.ps1` | `M3206-Set-OneDriveSiteCollectionAdmins.input.csv` | OneDrive/SharePoint | Add/remove OneDrive site collection administrators by user | Implemented |
| M3207 | `M3207-Set-OneDriveSharingSettings.ps1` | `M3207-Set-OneDriveSharingSettings.input.csv` | OneDrive/SharePoint | Update OneDrive sharing policy posture by user/site | Planned |
| M3208 | `M3208-Revoke-OneDriveExternalSharingLinks.ps1` | `M3208-Revoke-OneDriveExternalSharingLinks.input.csv` | OneDrive/SharePoint | Revoke scoped OneDrive external links/principal shares | Planned |
| M3209 | `M3209-Set-OneDriveSiteLockState.ps1` | `M3209-Set-OneDriveSiteLockState.input.csv` | OneDrive/SharePoint | Set OneDrive lock state for lifecycle/incident controls | Planned |
| M3005 | `M3005-Update-EntraAssignedSecurityGroups.ps1` | `M3005-Update-EntraAssignedSecurityGroups.input.csv` | Entra | Assigned security group property updates | Planned |
| M3006 | `M3006-Update-EntraDynamicUserSecurityGroups.ps1` | `M3006-Update-EntraDynamicUserSecurityGroups.input.csv` | Entra | Dynamic rule/processing updates | Planned |
| M3007 | `M3007-Set-EntraSecurityGroupMembers.ps1` | `M3007-Set-EntraSecurityGroupMembers.input.csv` | Entra | Add group members | Implemented |
| M3008 | `M3008-Update-EntraMicrosoft365Groups.ps1` | `M3008-Update-EntraMicrosoft365Groups.input.csv` | Entra | M365 group settings/visibility updates | Planned |
| M3309 | `M3309-Update-MicrosoftTeams.ps1` | `M3309-Update-MicrosoftTeams.input.csv` | Teams | Team settings updates | Planned |
| M3310 | `M3310-Set-MicrosoftTeamMembers.ps1` | `M3310-Set-MicrosoftTeamMembers.input.csv` | Teams | Add Team owners/members | Implemented |
| M3311 | `M3311-Update-MicrosoftTeamChannels.ps1` | `M3311-Update-MicrosoftTeamChannels.input.csv` | Teams | Add channels to existing Teams | Implemented |
| M3312 | `M3312-Set-MicrosoftTeamChannelMembers.ps1` | `M3312-Set-MicrosoftTeamChannelMembers.input.csv` | Teams | Add/update private/shared channel members | Implemented |
| M3113 | `M3113-Update-ExchangeOnlineMailContacts.ps1` | `M3113-Update-ExchangeOnlineMailContacts.input.csv` | Exchange Online | Mail contact property updates | Implemented |
| M3114 | `M3114-Update-ExchangeOnlineDistributionLists.ps1` | `M3114-Update-ExchangeOnlineDistributionLists.input.csv` | Exchange Online | Distribution list property updates | Implemented |
| M3115 | `M3115-Set-ExchangeOnlineDistributionListMembers.ps1` | `M3115-Set-ExchangeOnlineDistributionListMembers.input.csv` | Exchange Online | Add distribution list members | Implemented |
| M3116 | `M3116-Update-ExchangeOnlineSharedMailboxes.ps1` | `M3116-Update-ExchangeOnlineSharedMailboxes.input.csv` | Exchange Online | Shared mailbox property updates | Implemented |
| M3117 | `M3117-Set-ExchangeOnlineSharedMailboxPermissions.ps1` | `M3117-Set-ExchangeOnlineSharedMailboxPermissions.input.csv` | Exchange Online | Configure shared mailbox permissions | Implemented |
| M3118 | `M3118-Update-ExchangeOnlineResourceMailboxes.ps1` | `M3118-Update-ExchangeOnlineResourceMailboxes.input.csv` | Exchange Online | Resource mailbox settings updates | Implemented |
| M3119 | `M3119-Set-ExchangeOnlineResourceMailboxBookingDelegates.ps1` | `M3119-Set-ExchangeOnlineResourceMailboxBookingDelegates.input.csv` | Exchange Online | Configure booking delegates/policy flags | Implemented |
| M3120 | `M3120-Set-ExchangeOnlineMailboxDelegations.ps1` | `M3120-Set-ExchangeOnlineMailboxDelegations.input.csv` | Exchange Online | Configure mailbox delegation rights | Implemented |
| M3121 | `M3121-Set-ExchangeOnlineMailboxFolderPermissions.ps1` | `M3121-Set-ExchangeOnlineMailboxFolderPermissions.input.csv` | Exchange Online | Configure folder permissions/delegate flags | Implemented |
| M3122 | `M3122-Update-ExchangeOnlineMailEnabledSecurityGroups.ps1` | `M3122-Update-ExchangeOnlineMailEnabledSecurityGroups.input.csv` | Exchange Online | Mail-enabled security group property updates | Implemented |
| M3123 | `M3123-Update-ExchangeOnlineDynamicDistributionGroups.ps1` | `M3123-Update-ExchangeOnlineDynamicDistributionGroups.input.csv` | Exchange Online | Dynamic distribution group property/filter updates | Implemented |
| M3241 | `M3241-Set-SharePointSiteAdmins.ps1` | `M3241-Set-SharePointSiteAdmins.input.csv` | SharePoint | Add/remove site collection administrators with last-admin safety guard | Implemented |
| M3243 | `M3243-Associate-SharePointSitesToHub.ps1` | `M3243-Associate-SharePointSitesToHub.input.csv` | SharePoint | Associate existing sites to hub sites (optional reassociation) | Implemented |

## Safety and Sequencing Guidance

Recommended execution phases:

1. Entra and OneDrive changes: `M3003`, `M3204`, `M3205`, `M3206`, `M3007`
2. Teams changes: `M3310` to `M3312`
3. Exchange Online changes: `M3113` to `M3123` (implemented Exchange Online set/update coverage)
4. SharePoint changes: `M3241`, `M3243`

High-impact change classes that should always start with `-WhatIf`:

- Membership removals (future remove-capable variants of `M3007`, `M3310`, `M3312`, `M3115`)
- Permission changes (`M3117`, `M3120`, `M3121`)
- SharePoint admin/hub association changes (`M3241`, `M3243`)
- Account state changes (`M3002`)

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










