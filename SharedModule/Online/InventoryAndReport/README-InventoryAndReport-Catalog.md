# InventoryAndReport Detailed Catalog

Detailed catalog for discovery/reporting scripts in `SharedModule/Online/InventoryAndReport/`.

Operational label: **Inventory and Report**.

Current implementation status: Entra, Teams, Exchange Online, OneDrive, and SharePoint scripts are implemented with expanded Entra/Exchange Online scope coverage (including accepted-domain verification records, mailbox analytics, consolidated permissions, and retention-tag tests) plus the planned OneDrive/Teams inventory wave.
Scope CSV templates are seeded with related sample values from the Provision sample input set.

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
- Use shared helpers from `./SharedModule/Common/Online/M365.Common.psm1` (repository-root path) where practical
- Reuse shared scope CSV files when key columns overlap

## ID Ranges

- `30xx`: Entra
- `31xx`: Exchange Online
- `32xx`: SharePoint/OneDrive
- `33xx`: Teams

## Catalog

| ID | Script | Input CSV | Workload | Primary Output Focus | Status |
|---|---|---|---|---|---|
| IR3001 | `SM-IR3001-Get-EntraUsers.ps1` | `Scope-Users.input.csv` | Entra | User inventory with expanded profile/contact/org/extension fields | Implemented |
| IR3002 | `SM-IR3002-Get-EntraGuestUsers.ps1` | `Scope-GuestUsers.input.csv` | Entra | Guest user inventory and invite state | Implemented |
| IR3003 | `SM-IR3003-Get-EntraUserLicenses.ps1` | `Scope-Users.input.csv` | Entra | User licensing assignments/plans | Implemented |
| IR3004 | `SM-IR3004-Get-EntraPrivilegedRoles.ps1` | `Scope-EntraPrivilegedRoles.input.csv` | Entra | Activated directory roles and assigned members | Implemented |
| IR3204 | `SM-IR3204-Get-OneDriveProvisioningStatus.ps1` | `Scope-Users.input.csv` | OneDrive/SharePoint | OneDrive URL and provisioning state by user | Implemented |
| IR3205 | `SM-IR3205-Get-OneDriveStorageAndQuota.ps1` | `Scope-Users.input.csv` | OneDrive/SharePoint | OneDrive storage usage and quota settings by user | Implemented |
| IR3206 | `SM-IR3206-Get-OneDriveSiteCollectionAdmins.ps1` | `Scope-Users.input.csv` | OneDrive/SharePoint | OneDrive site collection admin inventory by user/site | Implemented |
| IR3207 | `SM-IR3207-Get-OneDriveSharingSettings.ps1` | `Scope-Users.input.csv` | OneDrive/SharePoint | OneDrive sharing policy posture by user/site | Implemented |
| IR3208 | `SM-IR3208-Get-OneDriveExternalSharingLinks.ps1` | `Scope-Users.input.csv` | OneDrive/SharePoint | OneDrive external principal sharing inventory | Implemented |
| IR3209 | `SM-IR3209-Get-OneDriveSiteLockState.ps1` | `Scope-Users.input.csv` | OneDrive/SharePoint | OneDrive lock state and access freeze posture | Implemented |
| IR3005 | `SM-IR3005-Get-EntraSecurityGroups.ps1` | `Scope-EntraSecurityGroups.input.csv` | Entra | Assigned security groups with expanded metadata fields | Implemented |
| IR3006 | `SM-IR3006-Get-EntraDynamicUserSecurityGroups.ps1` | `Scope-EntraDynamicUserSecurityGroups.input.csv` | Entra | Dynamic groups with expanded metadata and rule fields | Implemented |
| IR3007 | `SM-IR3007-Get-EntraSecurityGroupMembers.ps1` | `Scope-EntraSecurityGroups.input.csv` | Entra | Group membership exports | Implemented |
| IR3008 | `SM-IR3008-Get-EntraMicrosoft365Groups.ps1` | `Scope-M365Groups.input.csv` | Entra | Microsoft 365 group config with expanded metadata, owners, and members | Implemented |
| IR3309 | `SM-IR3309-Get-MicrosoftTeams.ps1` | `Scope-Teams.input.csv` | Teams | Teams inventory and core settings | Implemented |
| IR3310 | `SM-IR3310-Get-MicrosoftTeamMembers.ps1` | `Scope-Teams.input.csv` | Teams | Team owner/member assignments | Implemented |
| IR3311 | `SM-IR3311-Get-MicrosoftTeamChannels.ps1` | `Scope-Teams.input.csv` | Teams | Channel inventory by Team | Implemented |
| IR3312 | `SM-IR3312-Get-MicrosoftTeamChannelMembers.ps1` | `SM-IR3312-Get-MicrosoftTeamChannelMembers.input.csv` | Teams | Private/shared channel membership | Implemented |
| IR3113 | `SM-IR3113-Get-ExchangeOnlineMailContacts.ps1` | `Scope-MailContacts.input.csv` | Exchange Online | Mail contact inventory | Implemented |
| IR3114 | `SM-IR3114-Get-ExchangeOnlineDistributionLists.ps1` | `Scope-DistributionLists.input.csv` | Exchange Online | Distribution list inventory | Implemented |
| IR3115 | `SM-IR3115-Get-ExchangeOnlineDistributionListMembers.ps1` | `Scope-DistributionLists.input.csv` | Exchange Online | Distribution list membership | Implemented |
| IR3116 | `SM-IR3116-Get-ExchangeOnlineSharedMailboxes.ps1` | `Scope-SharedMailboxes.input.csv` | Exchange Online | Shared mailbox inventory | Implemented |
| IR3117 | `SM-IR3117-Get-ExchangeOnlineSharedMailboxPermissions.ps1` | `Scope-SharedMailboxes.input.csv` | Exchange Online | Shared mailbox permission matrix | Implemented |
| IR3118 | `SM-IR3118-Get-ExchangeOnlineResourceMailboxes.ps1` | `Scope-ResourceMailboxes.input.csv` | Exchange Online | Room/equipment mailbox inventory | Implemented |
| IR3119 | `SM-IR3119-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1` | `Scope-ResourceMailboxes.input.csv` | Exchange Online | Resource booking delegate/policy state | Implemented |
| IR3120 | `SM-IR3120-Get-ExchangeOnlineMailboxDelegations.ps1` | `Scope-Mailboxes.input.csv` | Exchange Online | Mailbox delegation matrix | Implemented |
| IR3121 | `SM-IR3121-Get-ExchangeOnlineMailboxFolderPermissions.ps1` | `Scope-Mailboxes.input.csv` | Exchange Online | Folder permission and delegate flags | Implemented |
| IR3122 | `SM-IR3122-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1` | `Scope-MailEnabledSecurityGroups.input.csv` | Exchange Online | Mail-enabled security group inventory | Implemented |
| IR3123 | `SM-IR3123-Get-ExchangeOnlineDynamicDistributionGroups.ps1` | `Scope-DynamicDistributionGroups.input.csv` | Exchange Online | Dynamic distribution group inventory | Implemented |
| IR3124 | `SM-IR3124-Get-ExchangeOnlineDomainVerificationRecords.ps1` | `Scope-AcceptedDomains.input.csv` | Exchange Online | Accepted-domain inventory with Entra verification DNS records | Implemented |
| IR3125 | `SM-IR3125-Get-ExchangeOnlineRecipientTypeCounts.ps1` | `SM-IR3125-Get-ExchangeOnlineRecipientTypeCounts.input.csv` | Exchange Online | Recipient counts grouped by RecipientTypeDetails | Implemented |
| IR3126 | `SM-IR3126-Get-ExchangeOnlineMailboxHighLevelStats.ps1` | `Scope-Mailboxes.input.csv` | Exchange Online | Aggregate mailbox size/item distribution metrics | Implemented |
| IR3127 | `SM-IR3127-Get-ExchangeOnlineMailboxSizes.ps1` | `Scope-Mailboxes.input.csv` | Exchange Online | Per-mailbox main/archive size and quota export | Implemented |
| IR3128 | `SM-IR3128-Get-ExchangeOnlineMailboxStatsPerMailbox.ps1` | `Scope-Mailboxes.input.csv` | Exchange Online | Detailed per-mailbox usage/activity statistics | Implemented |
| IR3129 | `SM-IR3129-Get-ExchangeOnlineMailboxPermissionsConsolidated.ps1` | `Scope-Mailboxes.input.csv` | Exchange Online | Consolidated delegated permission summary by mailbox | Implemented |
| IR3130 | `SM-IR3130-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.ps1` | `SM-IR3130-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.input.csv` | Exchange Online | Unexpected/missing mailbox retention tag detection | Implemented |
| IR3240 | `SM-IR3240-Get-SharePointSites.ps1` | `Scope-SharePointSites.input.csv` | SharePoint | Site inventory and core metadata | Implemented |

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
3. Entra and OneDrive baseline: `IR3001` to `IR3008`, plus `IR3204` to `IR3209`
4. Teams baseline: `IR3309` to `IR3312`
5. Exchange Online baseline: `IR3113` to `IR3130`
6. SharePoint baseline: `IR3240`

## Related Docs

- [InventoryAndReport README](./README.md)
- [SharedModule README](../../README.md)
- [Provision Detailed Catalog](../Provision/README-Provision-Catalog.md)





