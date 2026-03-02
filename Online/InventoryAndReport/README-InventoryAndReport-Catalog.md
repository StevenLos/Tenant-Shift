# InventoryAndReport Detailed Catalog

Detailed catalog for discovery/reporting scripts in `Online/InventoryAndReport/`.

Operational label: **Inventory and Report**.

Current implementation status: Entra, Teams, Exchange Online, OneDrive, and SharePoint scripts are implemented with expanded Exchange Online scope coverage.
Scope CSV templates are seeded with related sample values from the Provision sample build set.

## Script Contract

All discover scripts should:

- Be read-only (no create/update/delete actions)
- Run on PowerShell 7+
- Require `-InputCsvPath` with validated required CSV headers
- Export deterministic CSV output for diff/baselining
- Include `Status` and `Message` per processed item
- Write a required per-run transcript log in the output folder
- Use shared helpers from `./Common/Online/M365.Common.psm1` (repository-root path) where practical
- Reuse shared scope CSV files when key columns overlap

## ID Ranges

- `30xx`: Entra
- `31xx`: Exchange Online
- `32xx`: SharePoint/OneDrive
- `33xx`: Teams

## Catalog

| ID | Script | Input CSV | Workload | Primary Output Focus | Status |
|---|---|---|---|---|---|
| IR3001 | `IR3001-Get-EntraUsers.ps1` | `Scope-Users.input.csv` | Entra | User inventory and core profile fields | Implemented |
| IR3002 | `IR3002-Get-EntraGuestUsers.ps1` | `Scope-GuestUsers.input.csv` | Entra | Guest user inventory and invite state | Implemented |
| IR3003 | `IR3003-Get-EntraUserLicenses.ps1` | `Scope-Users.input.csv` | Entra | User licensing assignments/plans | Implemented |
| IR3204 | `IR3204-Get-OneDriveProvisioningStatus.ps1` | `Scope-Users.input.csv` | OneDrive/SharePoint | OneDrive URL and provisioning state by user | Implemented |
| IR3205 | `IR3205-Get-OneDriveStorageAndQuota.ps1` | `Scope-Users.input.csv` | OneDrive/SharePoint | OneDrive storage usage and quota settings by user | Implemented |
| IR3206 | `IR3206-Get-OneDriveSiteCollectionAdmins.ps1` | `Scope-Users.input.csv` | OneDrive/SharePoint | OneDrive site collection admin inventory by user/site | Implemented |
| IR3207 | `IR3207-Get-OneDriveSharingSettings.ps1` | `Scope-Users.input.csv` | OneDrive/SharePoint | OneDrive sharing policy posture by user/site | Planned |
| IR3208 | `IR3208-Get-OneDriveExternalSharingLinks.ps1` | `Scope-Users.input.csv` | OneDrive/SharePoint | OneDrive external link and principal sharing inventory | Planned |
| IR3209 | `IR3209-Get-OneDriveSiteLockState.ps1` | `Scope-Users.input.csv` | OneDrive/SharePoint | OneDrive lock state and access freeze posture | Planned |
| IR3005 | `IR3005-Get-EntraSecurityGroups.ps1` | `Scope-EntraSecurityGroups.input.csv` | Entra | Assigned security groups | Implemented |
| IR3006 | `IR3006-Get-EntraDynamicUserSecurityGroups.ps1` | `Scope-EntraDynamicUserSecurityGroups.input.csv` | Entra | Dynamic groups and membership rules | Implemented |
| IR3007 | `IR3007-Get-EntraSecurityGroupMembers.ps1` | `Scope-EntraSecurityGroups.input.csv` | Entra | Group membership exports | Implemented |
| IR3008 | `IR3008-Get-EntraMicrosoft365Groups.ps1` | `Scope-M365Groups.input.csv` | Entra | Microsoft 365 group config and ownership | Implemented |
| IR3309 | `IR3309-Get-MicrosoftTeams.ps1` | `Scope-Teams.input.csv` | Teams | Teams inventory and core settings | Implemented |
| IR3310 | `IR3310-Get-MicrosoftTeamMembers.ps1` | `Scope-Teams.input.csv` | Teams | Team owner/member assignments | Implemented |
| IR3311 | `IR3311-Get-MicrosoftTeamChannels.ps1` | `Scope-Teams.input.csv` | Teams | Channel inventory by Team | Implemented |
| IR3312 | `IR3312-Get-MicrosoftTeamChannelMembers.ps1` | `IR3312-Get-MicrosoftTeamChannelMembers.input.csv` | Teams | Private/shared channel membership | Planned |
| IR3113 | `IR3113-Get-ExchangeOnlineMailContacts.ps1` | `Scope-MailContacts.input.csv` | Exchange Online | Mail contact inventory | Implemented |
| IR3114 | `IR3114-Get-ExchangeOnlineDistributionLists.ps1` | `Scope-DistributionLists.input.csv` | Exchange Online | Distribution list inventory | Implemented |
| IR3115 | `IR3115-Get-ExchangeOnlineDistributionListMembers.ps1` | `Scope-DistributionLists.input.csv` | Exchange Online | Distribution list membership | Implemented |
| IR3116 | `IR3116-Get-ExchangeOnlineSharedMailboxes.ps1` | `Scope-SharedMailboxes.input.csv` | Exchange Online | Shared mailbox inventory | Implemented |
| IR3117 | `IR3117-Get-ExchangeOnlineSharedMailboxPermissions.ps1` | `Scope-SharedMailboxes.input.csv` | Exchange Online | Shared mailbox permission matrix | Implemented |
| IR3118 | `IR3118-Get-ExchangeOnlineResourceMailboxes.ps1` | `Scope-ResourceMailboxes.input.csv` | Exchange Online | Room/equipment mailbox inventory | Implemented |
| IR3119 | `IR3119-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1` | `Scope-ResourceMailboxes.input.csv` | Exchange Online | Resource booking delegate/policy state | Implemented |
| IR3120 | `IR3120-Get-ExchangeOnlineMailboxDelegations.ps1` | `Scope-Mailboxes.input.csv` | Exchange Online | Mailbox delegation matrix | Implemented |
| IR3121 | `IR3121-Get-ExchangeOnlineMailboxFolderPermissions.ps1` | `Scope-Mailboxes.input.csv` | Exchange Online | Folder permission and delegate flags | Implemented |
| IR3122 | `IR3122-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1` | `Scope-MailEnabledSecurityGroups.input.csv` | Exchange Online | Mail-enabled security group inventory | Implemented |
| IR3123 | `IR3123-Get-ExchangeOnlineDynamicDistributionGroups.ps1` | `Scope-DynamicDistributionGroups.input.csv` | Exchange Online | Dynamic distribution group inventory | Implemented |
| IR3240 | `IR3240-Get-SharePointSites.ps1` | `Scope-SharePointSites.input.csv` | SharePoint | Site inventory and core metadata | Implemented |

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
2. Entra and OneDrive baseline: `IR3001` to `IR3008`, plus `IR3204` to `IR3206`
3. Teams baseline: `IR3309` to `IR3312`
4. Exchange Online baseline: `IR3113` to `IR3123`
5. SharePoint baseline: `IR3240`

## Related Docs

- [InventoryAndReport README](./README.md)
- [Root README](../../README.md)
- [Provision Detailed Catalog](../Provision/README-Provision-Catalog.md)











