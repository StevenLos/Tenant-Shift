# Modify Folder

`Modify` is for scripts that modify existing tenant objects/configuration.

Operational label: **Modify**.

Current status: update baseline scripts are implemented for Entra licensing/membership/password resets/group-creator policy, Teams settings/membership/channels, Exchange Online recipient/group updates plus accepted-domain controls, mailbox auth/profile operations, safe-sender management, and recoverable-item restore operations, and SharePoint/OneDrive admin/sharing/lifecycle controls.
The provided sample `.input.csv` files are aligned to the Provision sample build set so post-build modify and verification runs are consistent.

## Purpose

Use this folder for controlled change operations after initial provisioning:

- Attribute updates
- Membership changes
- Policy/permission changes
- Lifecycle changes on existing objects

Do not use this folder for:
- Initial object creation (use `TenantShift/Online/Provision`)
- Read-only reporting (use `TenantShift/Online/Discover`)

## Naming Standard

- Script: `M-<WW>-<NNNN>-<Action>-<Target>.ps1`
- Input template: `M-<WW>-<NNNN>-<Action>-<Target>.input.csv`
- Output pattern: `Results_M-<WW>-<NNNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Transcript log pattern: `Transcript_M-<WW>-<NNNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`
- Default output directory (unless `-OutputCsvPath` is supplied): `./TenantShift/Online/Modify/Modify_OutputCsvPath/`

Workload code allocation (`WW` in `M-<WW>-<NNNN>`):
- `MEID`: Entra
- `EXOL`: Exchange Online
- `ONDR`: OneDrive
- `SPOL`: SharePoint
- `TEAM`: Teams

Example:
- `M-TEAM-0020-Set-MicrosoftTeamMembers.ps1`
- `M-TEAM-0020-Set-MicrosoftTeamMembers.input.csv`

## Implemented Scripts

| ID | Script | Workload | Purpose |
|---|---|---|---|
| M-MEID-0010 | `M-MEID-0010-Update-EntraUsers.ps1` | Entra | Update user profile attributes using the expanded Entra field model and clear/reset semantics, including fold-in password reset support. |
| M-MEID-0040 | `M-MEID-0040-Set-EntraUserAccountState.ps1` | Entra | Enable/disable user accounts. |
| M-MEID-0030 | `M-MEID-0030-Set-EntraUserLicenses.ps1` | Entra | Add/update user license assignments. |
| M-MEID-0050 | `M-MEID-0050-Set-EntraUserPasswordResets.ps1` | Entra | Standalone user password reset workflow with force-change controls. |
| M-MEID-0070 | `M-MEID-0070-Update-EntraAssignedSecurityGroups.ps1` | Entra | Update assigned security group properties with clear/reset support. |
| M-MEID-0080 | `M-MEID-0080-Update-EntraDynamicUserSecurityGroups.ps1` | Entra | Update dynamic membership rules/settings with clear/reset support. |
| M-ONDR-0010 | `M-ONDR-0010-PreProvision-OneDrive.ps1` | OneDrive | Trigger OneDrive pre-provisioning for existing users. |
| M-ONDR-0020 | `M-ONDR-0020-Set-OneDriveStorageQuota.ps1` | OneDrive | Update OneDrive storage quota and warning level by user. |
| M-ONDR-0030 | `M-ONDR-0030-Set-OneDriveSiteCollectionAdmins.ps1` | OneDrive | Add/remove OneDrive site collection administrators by user. |
| M-ONDR-0040 | `M-ONDR-0040-Set-OneDriveSharingSettings.ps1` | OneDrive | Update OneDrive sharing settings by user/site. |
| M-ONDR-0050 | `M-ONDR-0050-Revoke-OneDriveExternalSharingLinks.ps1` | OneDrive | Revoke external sharing principals/links on OneDrive sites by user. |
| M-ONDR-0060 | `M-ONDR-0060-Set-OneDriveSiteLockState.ps1` | OneDrive | Set OneDrive site lock state for incident or lifecycle controls. |
| M-MEID-0090 | `M-MEID-0090-Set-EntraSecurityGroupMembers.ps1` | Entra | Add/remove users in existing security groups. |
| M-MEID-0100 | `M-MEID-0100-Update-EntraMicrosoft365Groups.ps1` | Entra | Update M365 group properties/visibility with clear/reset support. |
| M-MEID-0130 | `M-MEID-0130-Set-EntraGroupCreatorsPolicy.ps1` | Entra | Configure tenant-wide Microsoft 365 group creator policy and allowed-creator group. |
| M-TEAM-0010 | `M-TEAM-0010-Update-MicrosoftTeams.ps1` | Teams | Update Team display/profile and settings families. |
| M-TEAM-0020 | `M-TEAM-0020-Set-MicrosoftTeamMembers.ps1` | Teams | Add users to existing Teams as members/owners. |
| M-TEAM-0030 | `M-TEAM-0030-Update-MicrosoftTeamChannels.ps1` | Teams | Add channels to existing Teams. |
| M-TEAM-0040 | `M-TEAM-0040-Set-MicrosoftTeamChannelMembers.ps1` | Teams | Add/update private/shared channel membership. |
| M-EXOL-0030 | `M-EXOL-0030-Update-ExchangeOnlineMailContacts.ps1` | Exchange Online | Update mail contact attributes. |
| M-EXOL-0040 | `M-EXOL-0040-Update-ExchangeOnlineDistributionLists.ps1` | Exchange Online | Update distribution list properties and sender restrictions. |
| M-EXOL-0090 | `M-EXOL-0090-Set-ExchangeOnlineDistributionListMembers.ps1` | Exchange Online | Add/remove members in existing distribution lists. |
| M-EXOL-0070 | `M-EXOL-0070-Update-ExchangeOnlineSharedMailboxes.ps1` | Exchange Online | Update shared mailbox properties, forwarding, and compliance toggles. |
| M-EXOL-0110 | `M-EXOL-0110-Set-ExchangeOnlineSharedMailboxPermissions.ps1` | Exchange Online | Configure shared mailbox permissions. |
| M-EXOL-0080 | `M-EXOL-0080-Update-ExchangeOnlineResourceMailboxes.ps1` | Exchange Online | Update room/equipment mailbox settings and booking controls. |
| M-EXOL-0120 | `M-EXOL-0120-Set-ExchangeOnlineResourceMailboxBookingDelegates.ps1` | Exchange Online | Configure resource mailbox booking delegates/policies. |
| M-EXOL-0130 | `M-EXOL-0130-Set-ExchangeOnlineMailboxDelegations.ps1` | Exchange Online | Configure mailbox delegations. |
| M-EXOL-0140 | `M-EXOL-0140-Set-ExchangeOnlineMailboxFolderPermissions.ps1` | Exchange Online | Configure mailbox folder-level permissions/delegates. |
| M-EXOL-0050 | `M-EXOL-0050-Update-ExchangeOnlineMailEnabledSecurityGroups.ps1` | Exchange Online | Update mail-enabled security group properties and sender restrictions. |
| M-EXOL-0060 | `M-EXOL-0060-Update-ExchangeOnlineDynamicDistributionGroups.ps1` | Exchange Online | Update dynamic distribution group filters and messaging controls. |
| M-EXOL-0150 | `M-EXOL-0150-Set-ExchangeOnlineUserMailboxForwarding.ps1` | Exchange Online | Set per-user mailbox forwarding mode and delivery behavior. |
| M-EXOL-0160 | `M-EXOL-0160-Update-ExchangeOnlineProxyAddresses.ps1` | Exchange Online | Add/remove/replace mailbox proxy addresses, including clear-secondary behavior. |
| M-EXOL-0170 | `M-EXOL-0170-Set-ExchangeOnlineMailboxAllowedAuthenticationMethods.ps1` | Exchange Online | Set mailbox protocol/authentication method posture via CAS mailbox controls. |
| M-EXOL-0180 | `M-EXOL-0180-Set-ExchangeOnlineUserPhotos.ps1` | Exchange Online | Set/remove mailbox photos from file paths with preview workflow support where available. |
| M-EXOL-0190 | `M-EXOL-0190-Set-ExchangeOnlineSafeSenderDomains.ps1` | Exchange Online | Add/remove/replace mailbox safe sender/trusted domain entries. |
| M-EXOL-0200 | `M-EXOL-0200-Restore-ExchangeOnlineRecoverableItems.ps1` | Exchange Online | Run recoverable-item restore workflow per mailbox with preview support. |
| M-EXOL-0010 | `M-EXOL-0010-Verify-ExchangeOnlineAcceptedDomains.ps1` | Exchange Online | Attempt/confirm Entra domain verification and validate Exchange accepted-domain presence. |
| M-EXOL-0020 | `M-EXOL-0020-Remove-ExchangeOnlineAcceptedDomains.ps1` | Exchange Online | Remove accepted domains and optional Entra tenant domains with protection controls. |
| M-SPOL-0030 | `M-SPOL-0030-Set-SharePointSiteAdmins.ps1` | SharePoint | Add/remove site collection administrators. |
| M-SPOL-0040 | `M-SPOL-0040-Associate-SharePointSitesToHub.ps1` | SharePoint | Associate existing sites to existing hubs. |

## Run Pattern

Run from repository root:

```powershell
pwsh ./TenantShift/Online/Modify/M-MEID-0030-Set-EntraUserLicenses.ps1 -InputCsvPath ./TenantShift/Online/Modify/M-MEID-0030-Set-EntraUserLicenses.input.csv -WhatIf
pwsh ./TenantShift/Online/Modify/M-MEID-0050-Set-EntraUserPasswordResets.ps1 -InputCsvPath ./TenantShift/Online/Modify/M-MEID-0050-Set-EntraUserPasswordResets.input.csv -WhatIf
pwsh ./TenantShift/Online/Modify/M-MEID-0130-Set-EntraGroupCreatorsPolicy.ps1 -InputCsvPath ./TenantShift/Online/Modify/M-MEID-0130-Set-EntraGroupCreatorsPolicy.input.csv -WhatIf
pwsh ./TenantShift/Online/Modify/M-TEAM-0020-Set-MicrosoftTeamMembers.ps1 -InputCsvPath ./TenantShift/Online/Modify/M-TEAM-0020-Set-MicrosoftTeamMembers.input.csv -WhatIf
pwsh ./TenantShift/Online/Modify/M-ONDR-0020-Set-OneDriveStorageQuota.ps1 -InputCsvPath ./TenantShift/Online/Modify/M-ONDR-0020-Set-OneDriveStorageQuota.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
pwsh ./TenantShift/Online/Modify/M-ONDR-0030-Set-OneDriveSiteCollectionAdmins.ps1 -InputCsvPath ./TenantShift/Online/Modify/M-ONDR-0030-Set-OneDriveSiteCollectionAdmins.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
pwsh ./TenantShift/Online/Modify/M-ONDR-0040-Set-OneDriveSharingSettings.ps1 -InputCsvPath ./TenantShift/Online/Modify/M-ONDR-0040-Set-OneDriveSharingSettings.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
pwsh ./TenantShift/Online/Modify/M-ONDR-0050-Revoke-OneDriveExternalSharingLinks.ps1 -InputCsvPath ./TenantShift/Online/Modify/M-ONDR-0050-Revoke-OneDriveExternalSharingLinks.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
pwsh ./TenantShift/Online/Modify/M-ONDR-0060-Set-OneDriveSiteLockState.ps1 -InputCsvPath ./TenantShift/Online/Modify/M-ONDR-0060-Set-OneDriveSiteLockState.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
pwsh ./TenantShift/Online/Modify/M-TEAM-0010-Update-MicrosoftTeams.ps1 -InputCsvPath ./TenantShift/Online/Modify/M-TEAM-0010-Update-MicrosoftTeams.input.csv -WhatIf
pwsh ./TenantShift/Online/Modify/M-SPOL-0030-Set-SharePointSiteAdmins.ps1 -InputCsvPath ./TenantShift/Online/Modify/M-SPOL-0030-Set-SharePointSiteAdmins.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
```

For copy/paste command building, use `./TenantShift/Online/Modify/Modify-Orchestrator.xlsx`.

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
- Reuse `./TenantShift/Common/Online/M365.Common.psm1` (repository-root path) for shared validation, connectivity, and result handling.
- Support bulk operations with per-record outcome tracking.

## References

- [Root README](../../README.md)
- [Provision README](../Provision/README.md)
- [Modify Detailed Catalog](./README-Modify-Catalog.md)
- [Entra User Field Contract](../README-Entra-User-Field-Contract.md)
- [Operator Runbook](./RUNBOOK-Modify.md)
