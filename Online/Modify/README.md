# Modify Folder

`Modify` is for scripts that modify existing tenant objects/configuration.

Operational label: **Modify**.

Current status: update baseline scripts are implemented for Entra licensing/membership, Teams settings/membership/channels, Exchange Online recipient/group updates and permission/delegation, and SharePoint/OneDrive admin/sharing/lifecycle controls.
The provided sample `.input.csv` files are aligned to the Provision sample build set so post-build modify and verification runs are consistent.

## Purpose

Use this folder for controlled change operations after initial provisioning:

- Attribute updates
- Membership changes
- Policy/permission changes
- Lifecycle changes on existing objects

Do not use this folder for:
- Initial object creation (use `Online/Provision`)
- Read-only reporting (use `Online/InventoryAndReport`)

## Naming Standard

- Script: `MWWNN-<Action>-<Target>.ps1`
- Input template: `MWWNN-<Action>-<Target>.input.csv`
- Output pattern: `Results_MWWNN-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Transcript log pattern: `Transcript_MWWNN-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`
- Default output directory (unless `-OutputCsvPath` is supplied): `./Online/Modify/Modify_OutputCsvPath/`

Workload code allocation (`WW` in `<Prefix><WW><NN>`):
- `30xx`: Entra
- `31xx`: Exchange Online
- `32xx`: SharePoint/OneDrive
- `33xx`: Teams

Example:
- `M3310-Set-MicrosoftTeamMembers.ps1`
- `M3310-Set-MicrosoftTeamMembers.input.csv`

## Implemented Scripts

| ID | Script | Workload | Purpose |
|---|---|---|---|
| M3001 | `M3001-Update-EntraUsers.ps1` | Entra | Update user profile attributes using the expanded Entra field model and clear/reset semantics. |
| M3002 | `M3002-Set-EntraUserAccountState.ps1` | Entra | Enable/disable user accounts. |
| M3003 | `M3003-Set-EntraUserLicenses.ps1` | Entra | Add/update user license assignments. |
| M3005 | `M3005-Update-EntraAssignedSecurityGroups.ps1` | Entra | Update assigned security group properties with clear/reset support. |
| M3006 | `M3006-Update-EntraDynamicUserSecurityGroups.ps1` | Entra | Update dynamic membership rules/settings with clear/reset support. |
| M3204 | `M3204-PreProvision-OneDrive.ps1` | OneDrive/SharePoint | Trigger OneDrive pre-provisioning for existing users. |
| M3205 | `M3205-Set-OneDriveStorageQuota.ps1` | OneDrive/SharePoint | Update OneDrive storage quota and warning level by user. |
| M3206 | `M3206-Set-OneDriveSiteCollectionAdmins.ps1` | OneDrive/SharePoint | Add/remove OneDrive site collection administrators by user. |
| M3207 | `M3207-Set-OneDriveSharingSettings.ps1` | OneDrive/SharePoint | Update OneDrive sharing settings by user/site. |
| M3208 | `M3208-Revoke-OneDriveExternalSharingLinks.ps1` | OneDrive/SharePoint | Revoke external sharing principals/links on OneDrive sites by user. |
| M3209 | `M3209-Set-OneDriveSiteLockState.ps1` | OneDrive/SharePoint | Set OneDrive site lock state for incident or lifecycle controls. |
| M3007 | `M3007-Set-EntraSecurityGroupMembers.ps1` | Entra | Add/remove users in existing security groups. |
| M3008 | `M3008-Update-EntraMicrosoft365Groups.ps1` | Entra | Update M365 group properties/visibility with clear/reset support. |
| M3309 | `M3309-Update-MicrosoftTeams.ps1` | Teams | Update Team display/profile and settings families. |
| M3310 | `M3310-Set-MicrosoftTeamMembers.ps1` | Teams | Add users to existing Teams as members/owners. |
| M3311 | `M3311-Update-MicrosoftTeamChannels.ps1` | Teams | Add channels to existing Teams. |
| M3312 | `M3312-Set-MicrosoftTeamChannelMembers.ps1` | Teams | Add/update private/shared channel membership. |
| M3113 | `M3113-Update-ExchangeOnlineMailContacts.ps1` | Exchange Online | Update mail contact attributes. |
| M3114 | `M3114-Update-ExchangeOnlineDistributionLists.ps1` | Exchange Online | Update distribution list properties and sender restrictions. |
| M3115 | `M3115-Set-ExchangeOnlineDistributionListMembers.ps1` | Exchange Online | Add/remove members in existing distribution lists. |
| M3116 | `M3116-Update-ExchangeOnlineSharedMailboxes.ps1` | Exchange Online | Update shared mailbox properties, forwarding, and compliance toggles. |
| M3117 | `M3117-Set-ExchangeOnlineSharedMailboxPermissions.ps1` | Exchange Online | Configure shared mailbox permissions. |
| M3118 | `M3118-Update-ExchangeOnlineResourceMailboxes.ps1` | Exchange Online | Update room/equipment mailbox settings and booking controls. |
| M3119 | `M3119-Set-ExchangeOnlineResourceMailboxBookingDelegates.ps1` | Exchange Online | Configure resource mailbox booking delegates/policies. |
| M3120 | `M3120-Set-ExchangeOnlineMailboxDelegations.ps1` | Exchange Online | Configure mailbox delegations. |
| M3121 | `M3121-Set-ExchangeOnlineMailboxFolderPermissions.ps1` | Exchange Online | Configure mailbox folder-level permissions/delegates. |
| M3122 | `M3122-Update-ExchangeOnlineMailEnabledSecurityGroups.ps1` | Exchange Online | Update mail-enabled security group properties and sender restrictions. |
| M3123 | `M3123-Update-ExchangeOnlineDynamicDistributionGroups.ps1` | Exchange Online | Update dynamic distribution group filters and messaging controls. |
| M3124 | `M3124-Set-ExchangeOnlineUserMailboxForwarding.ps1` | Exchange Online | Set per-user mailbox forwarding mode and delivery behavior. |
| M3241 | `M3241-Set-SharePointSiteAdmins.ps1` | SharePoint | Add/remove site collection administrators. |
| M3243 | `M3243-Associate-SharePointSitesToHub.ps1` | SharePoint | Associate existing sites to existing hubs. |

## Run Pattern

Run from repository root:

```powershell
pwsh ./Online/Modify/M3003-Set-EntraUserLicenses.ps1 -InputCsvPath ./Online/Modify/M3003-Set-EntraUserLicenses.input.csv -WhatIf
pwsh ./Online/Modify/M3310-Set-MicrosoftTeamMembers.ps1 -InputCsvPath ./Online/Modify/M3310-Set-MicrosoftTeamMembers.input.csv -WhatIf
pwsh ./Online/Modify/M3205-Set-OneDriveStorageQuota.ps1 -InputCsvPath ./Online/Modify/M3205-Set-OneDriveStorageQuota.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
pwsh ./Online/Modify/M3206-Set-OneDriveSiteCollectionAdmins.ps1 -InputCsvPath ./Online/Modify/M3206-Set-OneDriveSiteCollectionAdmins.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
pwsh ./Online/Modify/M3207-Set-OneDriveSharingSettings.ps1 -InputCsvPath ./Online/Modify/M3207-Set-OneDriveSharingSettings.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
pwsh ./Online/Modify/M3208-Revoke-OneDriveExternalSharingLinks.ps1 -InputCsvPath ./Online/Modify/M3208-Revoke-OneDriveExternalSharingLinks.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
pwsh ./Online/Modify/M3209-Set-OneDriveSiteLockState.ps1 -InputCsvPath ./Online/Modify/M3209-Set-OneDriveSiteLockState.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
pwsh ./Online/Modify/M3309-Update-MicrosoftTeams.ps1 -InputCsvPath ./Online/Modify/M3309-Update-MicrosoftTeams.input.csv -WhatIf
pwsh ./Online/Modify/M3241-Set-SharePointSiteAdmins.ps1 -InputCsvPath ./Online/Modify/M3241-Set-SharePointSiteAdmins.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
```

For copy/paste command building, use `./Online/Modify/Modify-Orchestrator.xlsx`.

## Required Safety Model

All modify scripts should include:

- `-WhatIf` support and clear `ShouldProcess` messaging
- Idempotent behavior where practical
- Per-record validation and error capture
- Timestamped result export with `Status` and `Message`
- Clear rollback or remediation notes for high-impact changes

## Modify Standards

- Keep workload explicit in script names.
- Include matched `.input.csv` templates for repeatable change sets.
- Reuse `./Common/Online/M365.Common.psm1` (repository-root path) for shared validation, connectivity, and result handling.
- Support bulk operations with per-record outcome tracking.

## References

- [Root README](../../README.md)
- [Provision README](../Provision/README.md)
- [Modify Detailed Catalog](./README-Modify-Catalog.md)
- [Entra User Field Contract](../README-Entra-User-Field-Contract.md)









