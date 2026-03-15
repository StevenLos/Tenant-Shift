# InventoryAndReport Folder

`InventoryAndReport` is for read-only inventory and reporting scripts.

Operational label: **Inventory and Report**.

Current status: inventory/report coverage includes Entra, Exchange Online, OneDrive, SharePoint, and Teams with dual discovery scope support (`-InputCsvPath` or `-DiscoverAll`), including Entra privileged-role inventory, Exchange accepted-domain verification records plus mailbox analytics/permission consolidations/retention-tag tests, OneDrive sharing/lock/external-principal reporting, and Teams channel-member reporting.

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
- Default output directory (unless `-OutputCsvPath` is supplied): `./SharedModule/Online/InventoryAndReport/InventoryAndReport_OutputCsvPath/`

Workload code allocation (`WW` in `<Prefix><WW><NN>`):
- `30xx`: Entra
- `31xx`: Exchange Online
- `32xx`: SharePoint/OneDrive
- `33xx`: Teams

Example:
- `SM-IR3001-Get-EntraUsers.ps1`
- `Scope-Users.input.csv`

## Run Pattern

Run from repository root:

```powershell
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3001-Get-EntraUsers.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3204-Get-OneDriveProvisioningStatus.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3205-Get-OneDriveStorageAndQuota.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3206-Get-OneDriveSiteCollectionAdmins.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3207-Get-OneDriveSharingSettings.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3208-Get-OneDriveExternalSharingLinks.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3209-Get-OneDriveSiteLockState.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3113-Get-ExchangeOnlineMailContacts.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-MailContacts.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3002-Get-EntraGuestUsers.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-GuestUsers.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3004-Get-EntraPrivilegedRoles.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-EntraPrivilegedRoles.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3006-Get-EntraDynamicUserSecurityGroups.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-EntraDynamicUserSecurityGroups.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3118-Get-ExchangeOnlineResourceMailboxes.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-ResourceMailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3120-Get-ExchangeOnlineMailboxDelegations.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3124-Get-ExchangeOnlineDomainVerificationRecords.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-AcceptedDomains.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3125-Get-ExchangeOnlineRecipientTypeCounts.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/SM-IR3125-Get-ExchangeOnlineRecipientTypeCounts.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3126-Get-ExchangeOnlineMailboxHighLevelStats.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3127-Get-ExchangeOnlineMailboxSizes.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3128-Get-ExchangeOnlineMailboxStatsPerMailbox.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3129-Get-ExchangeOnlineMailboxPermissionsConsolidated.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3130-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/SM-IR3130-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3309-Get-MicrosoftTeams.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Teams.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3240-Get-SharePointSites.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-SharePointSites.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3312-Get-MicrosoftTeamChannelMembers.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/SM-IR3312-Get-MicrosoftTeamChannelMembers.input.csv
```

Discover-all pattern examples:

```powershell
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3001-Get-EntraUsers.ps1 -DiscoverAll
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3004-Get-EntraPrivilegedRoles.ps1 -DiscoverAll
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3116-Get-ExchangeOnlineSharedMailboxes.ps1 -DiscoverAll
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3124-Get-ExchangeOnlineDomainVerificationRecords.ps1 -DiscoverAll
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3125-Get-ExchangeOnlineRecipientTypeCounts.ps1 -DiscoverAll
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3127-Get-ExchangeOnlineMailboxSizes.ps1 -DiscoverAll
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3130-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.ps1 -DiscoverAll
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3240-Get-SharePointSites.ps1 -DiscoverAll -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3312-Get-MicrosoftTeamChannelMembers.ps1 -DiscoverAll
```

## Discovery Scope Modes

- CSV-bounded mode: `-InputCsvPath` with validated required headers.
- Unbounded mode: `-DiscoverAll` (internally maps required scope headers to `*`).
- Both modes export a `ScopeMode` column in the results (`Csv` or `DiscoverAll`).

## Scoped Input Files

Use shared scope CSVs when multiple scripts use the same object key:

| Scope File | Key Column(s) | Reused By |
|---|---|---|
| `Scope-Users.input.csv` | `UserPrincipalName` | `IR3001`, `IR3003`, `IR3204`, `IR3205`, `IR3206`, `IR3207`, `IR3208`, `IR3209` |
| `Scope-GuestUsers.input.csv` | `UserPrincipalName` | `IR3002` |
| `Scope-EntraPrivilegedRoles.input.csv` | `RoleDisplayName` | `IR3004` |
| `Scope-EntraSecurityGroups.input.csv` | `GroupDisplayName` | `IR3005`, `IR3007` |
| `Scope-EntraDynamicUserSecurityGroups.input.csv` | `GroupDisplayName` | `IR3006` |
| `Scope-M365Groups.input.csv` | `GroupMailNickname` | `IR3008` |
| `Scope-Teams.input.csv` | `TeamMailNickname` | `IR3309`, `IR3310`, `IR3311` |
| `SM-IR3312-Get-MicrosoftTeamChannelMembers.input.csv` | `TeamMailNickname`, `ChannelDisplayName` | `IR3312` |
| `Scope-MailContacts.input.csv` | `MailContactIdentity` | `IR3113` |
| `Scope-DistributionLists.input.csv` | `DistributionGroupIdentity` | `IR3114`, `IR3115` |
| `Scope-MailEnabledSecurityGroups.input.csv` | `SecurityGroupIdentity` | `IR3122` |
| `Scope-DynamicDistributionGroups.input.csv` | `DynamicDistributionGroupIdentity` | `IR3123` |
| `Scope-AcceptedDomains.input.csv` | `DomainName` | `IR3124` |
| `SM-IR3125-Get-ExchangeOnlineRecipientTypeCounts.input.csv` | `RecipientIdentity` | `IR3125` |
| `Scope-SharedMailboxes.input.csv` | `SharedMailboxIdentity` | `IR3116`, `IR3117` |
| `Scope-ResourceMailboxes.input.csv` | `ResourceMailboxIdentity` | `IR3118`, `IR3119` |
| `Scope-Mailboxes.input.csv` | `MailboxIdentity` | `IR3120`, `IR3121`, `IR3126`, `IR3127`, `IR3128`, `IR3129` |
| `SM-IR3130-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.input.csv` | `MailboxIdentity`, `ExpectedTagNames` | `IR3130` |
| `Scope-SharePointSites.input.csv` | `SiteUrl` | `IR3240` |

Use `*` in scope files to inventory all objects for that key type (for example `Scope-GuestUsers.input.csv` defaults to `*` because guest UPN formats vary by tenant).

## Sample Input Verification

The scope CSVs are preloaded with sample objects from `SharedModule/Online/Provision/*.input.csv` so you can:

1. Run the Provision scripts with the provided sample set.
2. Run the IR scripts with the matching `Scope-*.input.csv` files to verify created objects and relationships quickly.

Example verification run set:

```powershell
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3001-Get-EntraUsers.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3002-Get-EntraGuestUsers.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-GuestUsers.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3003-Get-EntraUserLicenses.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3004-Get-EntraPrivilegedRoles.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-EntraPrivilegedRoles.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3204-Get-OneDriveProvisioningStatus.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3205-Get-OneDriveStorageAndQuota.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3206-Get-OneDriveSiteCollectionAdmins.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3207-Get-OneDriveSharingSettings.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3208-Get-OneDriveExternalSharingLinks.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3209-Get-OneDriveSiteLockState.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3005-Get-EntraSecurityGroups.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-EntraSecurityGroups.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3006-Get-EntraDynamicUserSecurityGroups.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-EntraDynamicUserSecurityGroups.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3007-Get-EntraSecurityGroupMembers.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-EntraSecurityGroups.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3008-Get-EntraMicrosoft365Groups.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-M365Groups.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3113-Get-ExchangeOnlineMailContacts.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-MailContacts.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3114-Get-ExchangeOnlineDistributionLists.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-DistributionLists.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3115-Get-ExchangeOnlineDistributionListMembers.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-DistributionLists.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3116-Get-ExchangeOnlineSharedMailboxes.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-SharedMailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3117-Get-ExchangeOnlineSharedMailboxPermissions.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-SharedMailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3118-Get-ExchangeOnlineResourceMailboxes.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-ResourceMailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3119-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-ResourceMailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3120-Get-ExchangeOnlineMailboxDelegations.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3121-Get-ExchangeOnlineMailboxFolderPermissions.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3122-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-MailEnabledSecurityGroups.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3123-Get-ExchangeOnlineDynamicDistributionGroups.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-DynamicDistributionGroups.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3124-Get-ExchangeOnlineDomainVerificationRecords.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-AcceptedDomains.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3125-Get-ExchangeOnlineRecipientTypeCounts.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/SM-IR3125-Get-ExchangeOnlineRecipientTypeCounts.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3126-Get-ExchangeOnlineMailboxHighLevelStats.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3127-Get-ExchangeOnlineMailboxSizes.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3128-Get-ExchangeOnlineMailboxStatsPerMailbox.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3129-Get-ExchangeOnlineMailboxPermissionsConsolidated.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Mailboxes.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3130-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/SM-IR3130-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3240-Get-SharePointSites.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-SharePointSites.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3309-Get-MicrosoftTeams.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Teams.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3310-Get-MicrosoftTeamMembers.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Teams.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3311-Get-MicrosoftTeamChannels.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/Scope-Teams.input.csv
pwsh ./SharedModule/Online/InventoryAndReport/SM-IR3312-Get-MicrosoftTeamChannelMembers.ps1 -InputCsvPath ./SharedModule/Online/InventoryAndReport/SM-IR3312-Get-MicrosoftTeamChannelMembers.input.csv
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
| IR3001 | `SM-IR3001-Get-EntraUsers.ps1` | Entra | Export tenant users with expanded profile/contact/org/extension fields. | Implemented |
| IR3002 | `SM-IR3002-Get-EntraGuestUsers.ps1` | Entra | Export guest users. | Implemented |
| IR3003 | `SM-IR3003-Get-EntraUserLicenses.ps1` | Entra | Export assigned licenses. | Implemented |
| IR3004 | `SM-IR3004-Get-EntraPrivilegedRoles.ps1` | Entra | Export activated directory roles and their assigned members. | Implemented |
| IR3204 | `SM-IR3204-Get-OneDriveProvisioningStatus.ps1` | OneDrive/SharePoint | Report OneDrive URL and provisioning status by user. | Implemented |
| IR3205 | `SM-IR3205-Get-OneDriveStorageAndQuota.ps1` | OneDrive/SharePoint | Report OneDrive storage usage and quota settings by user. | Implemented |
| IR3206 | `SM-IR3206-Get-OneDriveSiteCollectionAdmins.ps1` | OneDrive/SharePoint | Report OneDrive site collection admins by user. | Implemented |
| IR3207 | `SM-IR3207-Get-OneDriveSharingSettings.ps1` | OneDrive/SharePoint | Report OneDrive sharing policy posture by user/site. | Implemented |
| IR3208 | `SM-IR3208-Get-OneDriveExternalSharingLinks.ps1` | OneDrive/SharePoint | Report OneDrive external sharing principals/access by user/site. | Implemented |
| IR3209 | `SM-IR3209-Get-OneDriveSiteLockState.ps1` | OneDrive/SharePoint | Report OneDrive lock state and access freeze posture. | Implemented |
| IR3005 | `SM-IR3005-Get-EntraSecurityGroups.ps1` | Entra | Export assigned security groups. | Implemented |
| IR3006 | `SM-IR3006-Get-EntraDynamicUserSecurityGroups.ps1` | Entra | Export dynamic user groups and rules. | Implemented |
| IR3007 | `SM-IR3007-Get-EntraSecurityGroupMembers.ps1` | Entra | Export security group membership. | Implemented |
| IR3008 | `SM-IR3008-Get-EntraMicrosoft365Groups.ps1` | Entra | Export Microsoft 365 groups. | Implemented |
| IR3309 | `SM-IR3309-Get-MicrosoftTeams.ps1` | Teams | Export Teams and core settings. | Implemented |
| IR3310 | `SM-IR3310-Get-MicrosoftTeamMembers.ps1` | Teams | Export Team membership. | Implemented |
| IR3311 | `SM-IR3311-Get-MicrosoftTeamChannels.ps1` | Teams | Export channels by Team. | Implemented |
| IR3312 | `SM-IR3312-Get-MicrosoftTeamChannelMembers.ps1` | Teams | Export private/shared channel membership. | Implemented |
| IR3113 | `SM-IR3113-Get-ExchangeOnlineMailContacts.ps1` | Exchange Online | Export mail contacts. | Implemented |
| IR3114 | `SM-IR3114-Get-ExchangeOnlineDistributionLists.ps1` | Exchange Online | Export distribution lists. | Implemented |
| IR3115 | `SM-IR3115-Get-ExchangeOnlineDistributionListMembers.ps1` | Exchange Online | Export DL membership. | Implemented |
| IR3116 | `SM-IR3116-Get-ExchangeOnlineSharedMailboxes.ps1` | Exchange Online | Export shared mailboxes. | Implemented |
| IR3117 | `SM-IR3117-Get-ExchangeOnlineSharedMailboxPermissions.ps1` | Exchange Online | Export mailbox permissions. | Implemented |
| IR3118 | `SM-IR3118-Get-ExchangeOnlineResourceMailboxes.ps1` | Exchange Online | Export room/equipment mailboxes. | Implemented |
| IR3119 | `SM-IR3119-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1` | Exchange Online | Export booking delegate settings. | Implemented |
| IR3120 | `SM-IR3120-Get-ExchangeOnlineMailboxDelegations.ps1` | Exchange Online | Export mailbox delegations. | Implemented |
| IR3121 | `SM-IR3121-Get-ExchangeOnlineMailboxFolderPermissions.ps1` | Exchange Online | Export folder-level permissions. | Implemented |
| IR3122 | `SM-IR3122-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1` | Exchange Online | Export mail-enabled security groups. | Implemented |
| IR3123 | `SM-IR3123-Get-ExchangeOnlineDynamicDistributionGroups.ps1` | Exchange Online | Export dynamic distribution groups. | Implemented |
| IR3124 | `SM-IR3124-Get-ExchangeOnlineDomainVerificationRecords.ps1` | Exchange Online | Export accepted-domain verification record requirements and tenant-domain verification state. | Implemented |
| IR3125 | `SM-IR3125-Get-ExchangeOnlineRecipientTypeCounts.ps1` | Exchange Online | Export recipient counts grouped by RecipientTypeDetails for scoped/all recipients. | Implemented |
| IR3126 | `SM-IR3126-Get-ExchangeOnlineMailboxHighLevelStats.ps1` | Exchange Online | Export high-level mailbox size/item distribution statistics for scoped/all mailboxes. | Implemented |
| IR3127 | `SM-IR3127-Get-ExchangeOnlineMailboxSizes.ps1` | Exchange Online | Export per-mailbox main/archive size and quota summary. | Implemented |
| IR3128 | `SM-IR3128-Get-ExchangeOnlineMailboxStatsPerMailbox.ps1` | Exchange Online | Export detailed per-mailbox usage/activity statistics. | Implemented |
| IR3129 | `SM-IR3129-Get-ExchangeOnlineMailboxPermissionsConsolidated.ps1` | Exchange Online | Export one-row-per-mailbox consolidated delegated permission summary. | Implemented |
| IR3130 | `SM-IR3130-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.ps1` | Exchange Online | Detect unexpected/missing mailbox retention policy tags against expected baselines. | Implemented |
| IR3240 | `SM-IR3240-Get-SharePointSites.ps1` | SharePoint | Export SharePoint site inventory and core metadata. | Implemented |

## InventoryAndReport Standards

- Keep scripts read-only.
- Support either `-InputCsvPath` or `-DiscoverAll` for every `IR` script.
- Prefer shared scope CSV files where key columns overlap across scripts.
- Use script-specific `IRWWNN-...input.csv` templates only when a script requires unique scope shape not covered by shared scope files.
- Keep workload explicit in script names.
- Reuse `./SharedModule/Common/Online/M365.Common.psm1` (repository-root path) where common validation and result formatting helps.
- Prefer deterministic column ordering for easier diffing between snapshots.

## References

- [SharedModule README](../../README.md)
- [Provision README](../Provision/README.md)
- [InventoryAndReport Detailed Catalog](./README-InventoryAndReport-Catalog.md)
- [Entra User Field Contract](../README-Entra-User-Field-Contract.md)







