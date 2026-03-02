# InventoryAndReport Folder

`InventoryAndReport` is for read-only inventory and reporting scripts.

Operational label: **Inventory and Report**.

Current status: inventory/report coverage includes Entra, Exchange Online, OneDrive, SharePoint, and Teams with mandatory CSV input and shared scope CSVs seeded from the Provision sample build set.

## Purpose

Use this folder for:
- State baselines before/after provision or modify runs
- Compliance and audit exports
- Environment inventory snapshots by workload

Do not use this folder for:
- Creation, modification, or deletion operations

## Naming Standard

- Script: `IRWWNN-<Action>-<Target>.ps1`
- Input CSV: shared `Scope-*.input.csv` (preferred) or script-specific `IRWWNN-<Action>-<Target>.input.csv` when needed
- Output pattern: `Results_IRWWNN-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Transcript log pattern: `Transcript_IRWWNN-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`
- Default output directory (unless `-OutputCsvPath` is supplied): `./Online/InventoryAndReport/InventoryAndReport_OutputCsvPath/`

Workload code allocation (`WW` in `<Prefix><WW><NN>`):
- `30xx`: Entra
- `31xx`: Exchange Online
- `32xx`: SharePoint/OneDrive
- `33xx`: Teams

Example:
- `IR3001-Get-EntraUsers.ps1`
- `Scope-Users.input.csv`

## Run Pattern

Run from repository root:

```powershell
pwsh ./Online/InventoryAndReport/IR3001-Get-EntraUsers.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Users.input.csv
pwsh ./Online/InventoryAndReport/IR3204-Get-OneDriveProvisioningStatus.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./Online/InventoryAndReport/IR3205-Get-OneDriveStorageAndQuota.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./Online/InventoryAndReport/IR3206-Get-OneDriveSiteCollectionAdmins.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./Online/InventoryAndReport/IR3113-Get-ExchangeOnlineMailContacts.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-MailContacts.input.csv
pwsh ./Online/InventoryAndReport/IR3002-Get-EntraGuestUsers.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-GuestUsers.input.csv
pwsh ./Online/InventoryAndReport/IR3006-Get-EntraDynamicUserSecurityGroups.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-EntraDynamicUserSecurityGroups.input.csv
pwsh ./Online/InventoryAndReport/IR3118-Get-ExchangeOnlineResourceMailboxes.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-ResourceMailboxes.input.csv
pwsh ./Online/InventoryAndReport/IR3120-Get-ExchangeOnlineMailboxDelegations.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./Online/InventoryAndReport/IR3309-Get-MicrosoftTeams.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Teams.input.csv
pwsh ./Online/InventoryAndReport/IR3240-Get-SharePointSites.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-SharePointSites.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
```

For copy/paste command building, use `./Online/InventoryAndReport/InventoryAndReport-Orchestrator.xlsx`.

## Scoped Input Files

Use shared scope CSVs when multiple scripts use the same object key:

| Scope File | Key Column(s) | Reused By |
|---|---|---|
| `Scope-Users.input.csv` | `UserPrincipalName` | `IR3001`, `IR3003`, `IR3204`, `IR3205`, `IR3206` |
| `Scope-GuestUsers.input.csv` | `UserPrincipalName` | `IR3002` |
| `Scope-EntraSecurityGroups.input.csv` | `GroupDisplayName` | `IR3005`, `IR3007` |
| `Scope-EntraDynamicUserSecurityGroups.input.csv` | `GroupDisplayName` | `IR3006` |
| `Scope-M365Groups.input.csv` | `GroupMailNickname` | `IR3008` |
| `Scope-Teams.input.csv` | `TeamMailNickname` | `IR3309`, `IR3310`, `IR3311` |
| `Scope-MailContacts.input.csv` | `MailContactIdentity` | `IR3113` |
| `Scope-DistributionLists.input.csv` | `DistributionGroupIdentity` | `IR3114`, `IR3115` |
| `Scope-MailEnabledSecurityGroups.input.csv` | `SecurityGroupIdentity` | `IR3122` |
| `Scope-DynamicDistributionGroups.input.csv` | `DynamicDistributionGroupIdentity` | `IR3123` |
| `Scope-SharedMailboxes.input.csv` | `SharedMailboxIdentity` | `IR3116`, `IR3117` |
| `Scope-ResourceMailboxes.input.csv` | `ResourceMailboxIdentity` | `IR3118`, `IR3119` |
| `Scope-Mailboxes.input.csv` | `MailboxIdentity` | `IR3120`, `IR3121` |
| `Scope-SharePointSites.input.csv` | `SiteUrl` | `IR3240` |

Use `*` in scope files to inventory all objects for that key type (for example `Scope-GuestUsers.input.csv` defaults to `*` because guest UPN formats vary by tenant).

## Sample Build Set Verification

The scope CSVs are preloaded with sample objects from `Online/Provision/*.input.csv`, so you can:

1. Run the Provision scripts with the provided sample set.
2. Run the IR scripts with the matching `Scope-*.input.csv` files to verify created objects and relationships quickly.

Example verification run set:

```powershell
pwsh ./Online/InventoryAndReport/IR3001-Get-EntraUsers.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Users.input.csv
pwsh ./Online/InventoryAndReport/IR3002-Get-EntraGuestUsers.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-GuestUsers.input.csv
pwsh ./Online/InventoryAndReport/IR3003-Get-EntraUserLicenses.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Users.input.csv
pwsh ./Online/InventoryAndReport/IR3204-Get-OneDriveProvisioningStatus.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./Online/InventoryAndReport/IR3205-Get-OneDriveStorageAndQuota.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./Online/InventoryAndReport/IR3206-Get-OneDriveSiteCollectionAdmins.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./Online/InventoryAndReport/IR3005-Get-EntraSecurityGroups.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-EntraSecurityGroups.input.csv
pwsh ./Online/InventoryAndReport/IR3006-Get-EntraDynamicUserSecurityGroups.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-EntraDynamicUserSecurityGroups.input.csv
pwsh ./Online/InventoryAndReport/IR3007-Get-EntraSecurityGroupMembers.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-EntraSecurityGroups.input.csv
pwsh ./Online/InventoryAndReport/IR3008-Get-EntraMicrosoft365Groups.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-M365Groups.input.csv
pwsh ./Online/InventoryAndReport/IR3113-Get-ExchangeOnlineMailContacts.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-MailContacts.input.csv
pwsh ./Online/InventoryAndReport/IR3114-Get-ExchangeOnlineDistributionLists.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-DistributionLists.input.csv
pwsh ./Online/InventoryAndReport/IR3115-Get-ExchangeOnlineDistributionListMembers.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-DistributionLists.input.csv
pwsh ./Online/InventoryAndReport/IR3116-Get-ExchangeOnlineSharedMailboxes.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-SharedMailboxes.input.csv
pwsh ./Online/InventoryAndReport/IR3117-Get-ExchangeOnlineSharedMailboxPermissions.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-SharedMailboxes.input.csv
pwsh ./Online/InventoryAndReport/IR3118-Get-ExchangeOnlineResourceMailboxes.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-ResourceMailboxes.input.csv
pwsh ./Online/InventoryAndReport/IR3119-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-ResourceMailboxes.input.csv
pwsh ./Online/InventoryAndReport/IR3120-Get-ExchangeOnlineMailboxDelegations.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./Online/InventoryAndReport/IR3121-Get-ExchangeOnlineMailboxFolderPermissions.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./Online/InventoryAndReport/IR3122-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-MailEnabledSecurityGroups.input.csv
pwsh ./Online/InventoryAndReport/IR3123-Get-ExchangeOnlineDynamicDistributionGroups.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-DynamicDistributionGroups.input.csv
pwsh ./Online/InventoryAndReport/IR3240-Get-SharePointSites.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-SharePointSites.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./Online/InventoryAndReport/IR3309-Get-MicrosoftTeams.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Teams.input.csv
pwsh ./Online/InventoryAndReport/IR3310-Get-MicrosoftTeamMembers.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Teams.input.csv
pwsh ./Online/InventoryAndReport/IR3311-Get-MicrosoftTeamChannels.ps1 -InputCsvPath ./Online/InventoryAndReport/Scope-Teams.input.csv
```

## InventoryAndReport Output Standard

Discovery scripts should export consistent, easy-to-diff output:

- Primary object key columns (for example: `UserPrincipalName`, `GroupId`)
- Workload/object metadata columns
- `Status` and `Message` columns for per-record operation logging
- Timestamped output file names

## InventoryAndReport Catalog

| ID | Script | Workload | Purpose | Status |
|---|---|---|---|---|
| IR3001 | `IR3001-Get-EntraUsers.ps1` | Entra | Export tenant users. | Implemented |
| IR3002 | `IR3002-Get-EntraGuestUsers.ps1` | Entra | Export guest users. | Implemented |
| IR3003 | `IR3003-Get-EntraUserLicenses.ps1` | Entra | Export assigned licenses. | Implemented |
| IR3204 | `IR3204-Get-OneDriveProvisioningStatus.ps1` | OneDrive/SharePoint | Report OneDrive URL and provisioning status by user. | Implemented |
| IR3205 | `IR3205-Get-OneDriveStorageAndQuota.ps1` | OneDrive/SharePoint | Report OneDrive storage usage and quota settings by user. | Implemented |
| IR3206 | `IR3206-Get-OneDriveSiteCollectionAdmins.ps1` | OneDrive/SharePoint | Report OneDrive site collection admins by user. | Implemented |
| IR3207 | `IR3207-Get-OneDriveSharingSettings.ps1` | OneDrive/SharePoint | Report OneDrive sharing policy posture by user/site. | Planned |
| IR3208 | `IR3208-Get-OneDriveExternalSharingLinks.ps1` | OneDrive/SharePoint | Report OneDrive external sharing links and principal access. | Planned |
| IR3209 | `IR3209-Get-OneDriveSiteLockState.ps1` | OneDrive/SharePoint | Report OneDrive lock state and access freeze posture. | Planned |
| IR3005 | `IR3005-Get-EntraSecurityGroups.ps1` | Entra | Export assigned security groups. | Implemented |
| IR3006 | `IR3006-Get-EntraDynamicUserSecurityGroups.ps1` | Entra | Export dynamic user groups and rules. | Implemented |
| IR3007 | `IR3007-Get-EntraSecurityGroupMembers.ps1` | Entra | Export security group membership. | Implemented |
| IR3008 | `IR3008-Get-EntraMicrosoft365Groups.ps1` | Entra | Export Microsoft 365 groups. | Implemented |
| IR3309 | `IR3309-Get-MicrosoftTeams.ps1` | Teams | Export Teams and core settings. | Implemented |
| IR3310 | `IR3310-Get-MicrosoftTeamMembers.ps1` | Teams | Export Team membership. | Implemented |
| IR3311 | `IR3311-Get-MicrosoftTeamChannels.ps1` | Teams | Export channels by Team. | Implemented |
| IR3312 | `IR3312-Get-MicrosoftTeamChannelMembers.ps1` | Teams | Export private/shared channel membership. | Planned |
| IR3113 | `IR3113-Get-ExchangeOnlineMailContacts.ps1` | Exchange Online | Export mail contacts. | Implemented |
| IR3114 | `IR3114-Get-ExchangeOnlineDistributionLists.ps1` | Exchange Online | Export distribution lists. | Implemented |
| IR3115 | `IR3115-Get-ExchangeOnlineDistributionListMembers.ps1` | Exchange Online | Export DL membership. | Implemented |
| IR3116 | `IR3116-Get-ExchangeOnlineSharedMailboxes.ps1` | Exchange Online | Export shared mailboxes. | Implemented |
| IR3117 | `IR3117-Get-ExchangeOnlineSharedMailboxPermissions.ps1` | Exchange Online | Export mailbox permissions. | Implemented |
| IR3118 | `IR3118-Get-ExchangeOnlineResourceMailboxes.ps1` | Exchange Online | Export room/equipment mailboxes. | Implemented |
| IR3119 | `IR3119-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1` | Exchange Online | Export booking delegate settings. | Implemented |
| IR3120 | `IR3120-Get-ExchangeOnlineMailboxDelegations.ps1` | Exchange Online | Export mailbox delegations. | Implemented |
| IR3121 | `IR3121-Get-ExchangeOnlineMailboxFolderPermissions.ps1` | Exchange Online | Export folder-level permissions. | Implemented |
| IR3122 | `IR3122-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1` | Exchange Online | Export mail-enabled security groups. | Implemented |
| IR3123 | `IR3123-Get-ExchangeOnlineDynamicDistributionGroups.ps1` | Exchange Online | Export dynamic distribution groups. | Implemented |
| IR3240 | `IR3240-Get-SharePointSites.ps1` | SharePoint | Export SharePoint site inventory and core metadata. | Implemented |

## InventoryAndReport Standards

- Keep scripts read-only.
- Require `-InputCsvPath` for every `IR` script.
- Prefer shared scope CSV files where key columns overlap across scripts.
- Use script-specific `IRWWNN-...input.csv` templates only when a script requires unique scope shape not covered by shared scope files.
- Keep workload explicit in script names.
- Reuse `./Common/Online/M365.Common.psm1` (repository-root path) where common validation and result formatting helps.
- Prefer deterministic column ordering for easier diffing between snapshots.

## References

- [Root README](../../README.md)
- [Provision README](../Provision/README.md)
- [InventoryAndReport Detailed Catalog](./README-InventoryAndReport-Catalog.md)













