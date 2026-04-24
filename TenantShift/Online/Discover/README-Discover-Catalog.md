# Discover Detailed Catalog

Detailed catalog for discovery/reporting scripts in `TenantShift/Online/Discover/`.

Operational label: **Discover**.

Current implementation status: MEID, TEAM, EXOL, ONDR, and SPOL discover scripts are implemented with expanded coverage including flattened Microsoft 365 group owner/member exports, shared-mailbox SMTP-address exports, accepted-domain verification records, mailbox analytics, consolidated permissions, and retention-tag tests. A user-centric entitlement reconstruction series (0500 sequence) has been added across MEID, EXOL, SPOL, and TEAM workloads — these scripts take a user list as input and produce explainable access footprints with path attribution across all M365 workloads.

## Script Contract

All discover scripts should:

- Be read-only (no create/update/delete actions)
- Run on PowerShell 7+
- Support either:
  - `-InputCsvPath` with validated required CSV headers (CSV-bounded scope)
  - `-DiscoverAll` (unbounded scope)
- Export deterministic CSV output for diff/baselining
- Include `Status` and `Message` per processed item
- Write a required per-run transcript log in the output folder
- Use shared helpers from `./TenantShift/Common/Online/M365.Common.psm1` (repository-root path) where practical
- Reuse shared scope CSV files when key columns overlap

## Catalog

| ID | Script | Input CSV | Workload | Primary Output Focus | Status |
|---|---|---|---|---|---|
| D-MEID-0010 | `D-MEID-0010-Get-EntraUsers.ps1` | `Scope-Users.input.csv` | MEID | User inventory with expanded profile/contact/org/extension fields | Implemented |
| D-MEID-0020 | `D-MEID-0020-Get-EntraGuestUsers.ps1` | `Scope-GuestUsers.input.csv` | MEID | Guest user inventory and invite state | Implemented |
| D-MEID-0030 | `D-MEID-0030-Get-EntraUserLicenses.ps1` | `Scope-Users.input.csv` | MEID | User licensing assignments/plans | Implemented |
| D-MEID-0060 | `D-MEID-0060-Get-EntraPrivilegedRoles.ps1` | `Scope-EntraPrivilegedRoles.input.csv` | MEID | Activated directory roles and assigned members | Implemented |
| D-MEID-0070 | `D-MEID-0070-Get-EntraSecurityGroups.ps1` | `Scope-EntraSecurityGroups.input.csv` | MEID | Assigned security groups with expanded metadata fields | Implemented |
| D-MEID-0080 | `D-MEID-0080-Get-EntraDynamicUserSecurityGroups.ps1` | `Scope-EntraDynamicUserSecurityGroups.input.csv` | MEID | Dynamic groups with expanded metadata and rule fields | Implemented |
| D-MEID-0090 | `D-MEID-0090-Get-EntraSecurityGroupMembers.ps1` | `Scope-EntraSecurityGroups.input.csv` | MEID | Group membership exports | Implemented |
| D-MEID-0100 | `D-MEID-0100-Get-EntraMicrosoft365Groups.ps1` | `Scope-M365Groups.input.csv` | MEID | Microsoft 365 group config with expanded metadata, owners, and members | Implemented |
| D-MEID-0110 | `D-MEID-0110-Get-EntraMicrosoft365GroupMembers.ps1` | `Scope-M365Groups.input.csv` | MEID | Microsoft 365 group members flattened one row per member | Implemented |
| D-MEID-0120 | `D-MEID-0120-Get-EntraMicrosoft365GroupOwners.ps1` | `Scope-M365Groups.input.csv` | MEID | Microsoft 365 group owners flattened one row per owner | Implemented |
| D-EXOL-0010 | `D-EXOL-0010-Get-ExchangeOnlineDomainVerificationRecords.ps1` | `Scope-AcceptedDomains.input.csv` | EXOL | Accepted-domain inventory with Entra verification DNS records | Implemented |
| D-EXOL-0030 | `D-EXOL-0030-Get-ExchangeOnlineMailContacts.ps1` | `Scope-MailContacts.input.csv` | EXOL | Mail contact inventory | Implemented |
| D-EXOL-0040 | `D-EXOL-0040-Get-ExchangeOnlineDistributionLists.ps1` | `Scope-DistributionLists.input.csv` | EXOL | Distribution list inventory | Implemented |
| D-EXOL-0050 | `D-EXOL-0050-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1` | `Scope-MailEnabledSecurityGroups.input.csv` | EXOL | Mail-enabled security group inventory with summarized members | Implemented |
| D-EXOL-0060 | `D-EXOL-0060-Get-ExchangeOnlineDynamicDistributionGroups.ps1` | `Scope-DynamicDistributionGroups.input.csv` | EXOL | Dynamic distribution group inventory | Implemented |
| D-EXOL-0070 | `D-EXOL-0070-Get-ExchangeOnlineSharedMailboxes.ps1` | `Scope-SharedMailboxes.input.csv` | EXOL | Shared mailbox inventory with semicolon-delimited proxy/email addresses | Implemented |
| D-EXOL-0080 | `D-EXOL-0080-Get-ExchangeOnlineResourceMailboxes.ps1` | `Scope-ResourceMailboxes.input.csv` | EXOL | Room/equipment mailbox inventory | Implemented |
| D-EXOL-0090 | `D-EXOL-0090-Get-ExchangeOnlineDistributionListMembers.ps1` | `Scope-DistributionLists.input.csv` | EXOL | Distribution list membership | Implemented |
| D-EXOL-0100 | `D-EXOL-0100-Get-ExchangeOnlineMailEnabledSecurityGroupMembers.ps1` | `Scope-MailEnabledSecurityGroups.input.csv` | EXOL | Mail-enabled security group membership flattened one row per member | Implemented |
| D-EXOL-0110 | `D-EXOL-0110-Get-ExchangeOnlineSharedMailboxPermissions.ps1` | `Scope-SharedMailboxes.input.csv` | EXOL | Shared mailbox permission matrix | Implemented |
| D-EXOL-0120 | `D-EXOL-0120-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1` | `Scope-ResourceMailboxes.input.csv` | EXOL | Resource booking delegate/policy state | Implemented |
| D-EXOL-0130 | `D-EXOL-0130-Get-ExchangeOnlineMailboxDelegations.ps1` | `Scope-Mailboxes.input.csv` | EXOL | Mailbox delegation matrix | Implemented |
| D-EXOL-0140 | `D-EXOL-0140-Get-ExchangeOnlineMailboxFolderPermissions.ps1` | `Scope-Mailboxes.input.csv` | EXOL | Folder permission and delegate flags | Implemented |
| D-EXOL-0210 | `D-EXOL-0210-Get-ExchangeOnlineRecipientTypeCounts.ps1` | `D-EXOL-0210-Get-ExchangeOnlineRecipientTypeCounts.input.csv` | EXOL | Recipient counts grouped by RecipientTypeDetails | Implemented |
| D-EXOL-0220 | `D-EXOL-0220-Get-ExchangeOnlineMailboxHighLevelStats.ps1` | `Scope-Mailboxes.input.csv` | EXOL | Aggregate mailbox size/item distribution metrics | Implemented |
| D-EXOL-0230 | `D-EXOL-0230-Get-ExchangeOnlineMailboxSizes.ps1` | `Scope-Mailboxes.input.csv` | EXOL | Per-mailbox main/archive size and quota export | Implemented |
| D-EXOL-0240 | `D-EXOL-0240-Get-ExchangeOnlineMailboxStatsPerMailbox.ps1` | `Scope-Mailboxes.input.csv` | EXOL | Detailed per-mailbox usage/activity statistics | Implemented |
| D-EXOL-0250 | `D-EXOL-0250-Get-ExchangeOnlineMailboxPermissionsConsolidated.ps1` | `Scope-Mailboxes.input.csv` | EXOL | Consolidated delegated permission summary by mailbox | Implemented |
| D-EXOL-0260 | `D-EXOL-0260-Get-ExchangeOnlineUserMailboxSmtpAddresses.ps1` | `Scope-Mailboxes.input.csv` | EXOL | User mailbox SMTP/proxy-address inventory flattened one row per SMTP address | Implemented |
| D-EXOL-0270 | `D-EXOL-0270-Get-ExchangeOnlineSharedMailboxSmtpAddresses.ps1` | `Scope-SharedMailboxes.input.csv` | EXOL | Shared mailbox SMTP/proxy-address inventory flattened one row per SMTP address | Implemented |
| D-EXOL-0280 | `D-EXOL-0280-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.ps1` | `D-EXOL-0280-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.input.csv` | EXOL | Unexpected/missing mailbox retention tag detection | Implemented |
| D-ONDR-0010 | `D-ONDR-0010-Get-OneDriveProvisioningStatus.ps1` | `Scope-Users.input.csv` | ONDR | OneDrive URL and provisioning state by user | Implemented |
| D-ONDR-0020 | `D-ONDR-0020-Get-OneDriveStorageAndQuota.ps1` | `Scope-Users.input.csv` | ONDR | OneDrive storage usage and quota settings by user | Implemented |
| D-ONDR-0030 | `D-ONDR-0030-Get-OneDriveSiteCollectionAdmins.ps1` | `Scope-Users.input.csv` | ONDR | OneDrive site collection admin inventory by user/site | Implemented |
| D-ONDR-0040 | `D-ONDR-0040-Get-OneDriveSharingSettings.ps1` | `Scope-Users.input.csv` | ONDR | OneDrive sharing policy posture by user/site | Implemented |
| D-ONDR-0050 | `D-ONDR-0050-Get-OneDriveExternalSharingLinks.ps1` | `Scope-Users.input.csv` | ONDR | OneDrive external principal sharing inventory | Implemented |
| D-ONDR-0060 | `D-ONDR-0060-Get-OneDriveSiteLockState.ps1` | `Scope-Users.input.csv` | ONDR | OneDrive lock state and access freeze posture | Implemented |
| D-SPOL-0010 | `D-SPOL-0010-Get-SharePointSites.ps1` | `Scope-SharePointSites.input.csv` | SPOL | Site inventory and core metadata | Implemented |
| D-TEAM-0010 | `D-TEAM-0010-Get-MicrosoftTeams.ps1` | `Scope-Teams.input.csv` | TEAM | Teams inventory and core settings | Implemented |
| D-TEAM-0020 | `D-TEAM-0020-Get-MicrosoftTeamMembers.ps1` | `Scope-Teams.input.csv` | TEAM | Team owner/member assignments | Implemented |
| D-TEAM-0030 | `D-TEAM-0030-Get-MicrosoftTeamChannels.ps1` | `Scope-Teams.input.csv` | TEAM | Channel inventory by Team | Implemented |
| D-TEAM-0040 | `D-TEAM-0040-Get-MicrosoftTeamChannelMembers.ps1` | `D-TEAM-0040-Get-MicrosoftTeamChannelMembers.input.csv` | TEAM | Private/shared channel membership | Implemented |
| D-MEID-0500 | `D-MEID-0500-Get-EntraAssignedGroupsByUser.ps1` | `Scope-Users.input.csv` | MEID | Assigned security group memberships per user — direct and transitive with full membership path attribution | Implemented |
| D-MEID-0510 | `D-MEID-0510-Get-EntraDynamicGroupsByUser.ps1` | `Scope-Users.input.csv` | MEID | Dynamic security group evaluated memberships per user — based on confirmed evaluated membership, not rule text | Implemented |
| D-MEID-0520 | `D-MEID-0520-Get-EntraM365GroupsByUser.ps1` | `Scope-Users.input.csv` | MEID | Microsoft 365 group member and owner relationships per user with Relationship column (Member/Owner/Member+Owner) | Implemented |
| D-EXOL-0500 | `D-EXOL-0500-Get-ExchangeOnlineMailboxAccessByUser.ps1` | `Scope-Users.input.csv` | EXOL | Effective mailbox access per user across shared, resource, and equipment mailboxes — Full Access, ReadOnly, Send As, Send on Behalf — with direct/group-sourced attribution | Implemented |
| D-SPOL-0500 | `D-SPOL-0500-Get-SharePointSitesByUser.ps1` | `Scope-Users.input.csv` | SPOL | SharePoint site access per user across all access paths (SiteCollectionAdmin, SPGroupMember, M365GroupMember, TeamsGroupMember, DirectPermission) with assignment chain | Implemented |
| D-TEAM-0500 | `D-TEAM-0500-Get-MicrosoftTeamsByUser.ps1` | `Scope-Users.input.csv` | TEAM | Microsoft Teams membership and ownership per user with role (Member/Owner) and access type (Direct/GuestInvite) | Implemented |
| D-TEAM-0510 | `D-TEAM-0510-Get-MicrosoftTeamChannelsByUser.ps1` | `Scope-Users.input.csv` | TEAM | Teams channel access per user distinguished by channel type (Standard/Private/Shared) with access mechanism and cross-tenant detection | Implemented |

## Standard Output Columns

Recommended baseline columns:

- `RowNumber`
- `PrimaryKey`
- `Action`
- `Status`
- `Message`
- Workload/object-specific fields

## Suggested Execution Pattern

1. Maintain shared scope files (`Scope-*.input.csv`) for each key type.
2. For each script, choose either CSV-bounded scope (`-InputCsvPath`) or unbounded scope (`-DiscoverAll`).
3. MEID baseline: `D-MEID-0010` through `D-MEID-0120`
4. TEAM baseline: `D-TEAM-0010` through `D-TEAM-0040`
5. EXOL baseline: `D-EXOL-0010` through `D-EXOL-0280`
6. ONDR baseline: `D-ONDR-0010` through `D-ONDR-0060`
7. SPOL baseline: `D-SPOL-0010`

### User-Centric Entitlement Reconstruction (0500 series)

The 0500-sequence scripts are architecturally distinct from the baseline inventory scripts. They take `Scope-Users.input.csv` as input and reconstruct the effective access footprint for each user across workloads with path attribution. Run these when the goal is to answer *why* a user has access, not just *what* exists in the tenant.

8. Entitlement baseline: `D-MEID-0500`, `D-MEID-0510`, `D-MEID-0520`, `D-EXOL-0500`, `D-SPOL-0500`, `D-TEAM-0500`, `D-TEAM-0510`

> **Note:** `D-SPOL-0500` provides full path attribution for user SharePoint access and supersedes `D-SPOL-0030` for user-centric discovery use cases. `D-SPOL-0030` remains in place for resource-centric site-to-user reporting.
>
> **DiscoverAll behaviour:** For 0500-series scripts, `-DiscoverAll` enumerates all **licensed users** in the tenant rather than all resources of a type. For large tenants this can produce substantial output volume and elevated API call counts.

## Related Docs

- [Discover README](./README.md)
- [Root README](../../README.md)
- [Provision Detailed Catalog](../Provision/README-Provision-Catalog.md)
