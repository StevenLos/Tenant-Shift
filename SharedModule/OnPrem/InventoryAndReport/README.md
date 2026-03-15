# OnPrem InventoryAndReport

`SharedModule/OnPrem/InventoryAndReport` is for planned read-only inventory and reporting scripts targeting on-prem workloads.

Operational label: **Inventory and Report**.

Current status: ActiveDirectory inventory baseline is implemented with `SM-IR0001-Get-ActiveDirectoryUsers.ps1`, `SM-IR0002-Get-ActiveDirectoryContacts.ps1`, `SM-IR0005-Get-ActiveDirectorySecurityGroups.ps1`, `SM-IR0006-Get-ActiveDirectoryDistributionGroups.ps1`, `SM-IR0007-Get-ActiveDirectorySecurityGroupMembers.ps1`, `SM-IR0008-Get-ActiveDirectoryDistributionGroupMembers.ps1`, `SM-IR0009-Get-ActiveDirectoryOrganizationalUnits.ps1`, `SM-IR0010-Get-ActiveDirectoryUserRecursiveGroupMemberships.ps1`, `SM-IR0011-Get-ActiveDirectoryUsersWithoutGroupMemberships.ps1`, and `SM-IR0012-Get-ActiveDirectoryGroupsWithoutMembers.ps1`. ExchangeOnPrem inventory scripts are implemented with `SM-IR0213-Get-ExchangeOnPremMailContacts.ps1`, `SM-IR0214-Get-ExchangeOnPremDistributionLists.ps1`, `SM-IR0215-Get-ExchangeOnPremDistributionListMembers.ps1`, `SM-IR0216-Get-ExchangeOnPremSharedMailboxes.ps1`, `SM-IR0217-Get-ExchangeOnPremSharedMailboxPermissions.ps1`, `SM-IR0218-Get-ExchangeOnPremResourceMailboxes.ps1`, `SM-IR0219-Get-ExchangeOnPremResourceMailboxBookingDelegates.ps1`, `SM-IR0220-Get-ExchangeOnPremMailboxDelegations.ps1`, `SM-IR0221-Get-ExchangeOnPremMailboxFolderPermissions.ps1`, `SM-IR0222-Get-ExchangeOnPremMailEnabledSecurityGroups.ps1`, `SM-IR0223-Get-ExchangeOnPremDynamicDistributionGroups.ps1`, `SM-IR0224-Get-ExchangeOnPremInboundConnectorDetails.ps1`, `SM-IR0225-Get-ExchangeOnPremOutlookClientVersionsFromRpcLogs.ps1`, and `SM-IR0226-Get-ExchangeOnPremRpcLogExport.ps1`; remaining scripts are planned.

## Purpose

Use this folder for:

- Baseline snapshots before and after provision/modify runs
- Audit and compliance exports
- Object and permissions discovery for AD, ExchangeOnPrem, and FileServices

Do not use this folder for:

- Creation, modification, or deletion operations

## Naming Standard

- Script: `IRWWNN-<Action>-<Target>.ps1`
- Input CSV: shared `Scope-*.input.csv` (preferred) or script-specific `IRWWNN-<Action>-<Target>.input.csv` when needed
- Output pattern: `Results_IRWWNN-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Transcript log pattern: `Transcript_IRWWNN-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`

## Runtime Assumptions

- ActiveDirectory inventory scripts (`00xx`) run in Windows PowerShell `5.1`.
- ExchangeOnPrem inventory scripts (`02xx`) run in Exchange Management Shell (Windows PowerShell `5.1`).

## Discovery Scope Modes

- Default discovery model is CSV scope input via `-InputCsvPath`.
- Implemented ActiveDirectory and ExchangeOnPrem inventory scripts (`IR0001`, `IR0002`, `IR0005`, `IR0006`, `IR0007`, `IR0008`, `IR0009`, `IR0010`, `IR0011`, `IR0012`, `IR0213`, `IR0214`, `IR0215`, `IR0216`, `IR0217`, `IR0218`, `IR0219`, `IR0220`, `IR0221`, `IR0222`, `IR0223`, `IR0224`, `IR0225`, `IR0226`) also support unbounded discovery via `-DiscoverAll`.
- `-DiscoverAll` scope controls are script-specific:
  - Directory/object-focused scripts use controls such as `-SearchBase`, `-Server`, and `-MaxObjects`.
  - RPC-log-focused scripts (`IR0225`, `IR0226`) use controls such as `-LogPath`, `-LookbackDays`, and `-MaxObjects`.

Workload code allocation (`WW` in `<Prefix><WW><NN>`):

- `00xx`: ActiveDirectory
- `01xx`: GroupPolicy
- `02xx`: ExchangeOnPrem
- `03xx`: FileServices

## Script Matrix (Current Status)

| ID | Script | Workload | Primary Output Focus |
|---|---|---|---|
| IR0001 | `SM-IR0001-Get-ActiveDirectoryUsers.ps1` | ActiveDirectory | AD user inventory and key identity attributes. |
| IR0002 | `SM-IR0002-Get-ActiveDirectoryContacts.ps1` | ActiveDirectory | AD contact inventory and mail routing attributes. |
| IR0005 | `SM-IR0005-Get-ActiveDirectorySecurityGroups.ps1` | ActiveDirectory | AD security group inventory and scope/type metadata. |
| IR0006 | `SM-IR0006-Get-ActiveDirectoryDistributionGroups.ps1` | ActiveDirectory | AD distribution group inventory and scope/type metadata. |
| IR0007 | `SM-IR0007-Get-ActiveDirectorySecurityGroupMembers.ps1` | ActiveDirectory | Group membership exports for AD security groups. |
| IR0008 | `SM-IR0008-Get-ActiveDirectoryDistributionGroupMembers.ps1` | ActiveDirectory | Group membership exports for AD distribution groups. |
| IR0009 | `SM-IR0009-Get-ActiveDirectoryOrganizationalUnits.ps1` | ActiveDirectory | OU hierarchy and delegation boundary inventory. |
| IR0010 | `SM-IR0010-Get-ActiveDirectoryUserRecursiveGroupMemberships.ps1` | ActiveDirectory | Recursive group memberships per user, including primary group coverage. |
| IR0011 | `SM-IR0011-Get-ActiveDirectoryUsersWithoutGroupMemberships.ps1` | ActiveDirectory | AD users with no recursive group memberships. |
| IR0012 | `SM-IR0012-Get-ActiveDirectoryGroupsWithoutMembers.ps1` | ActiveDirectory | AD groups with no direct members and no primary-group users. |
| IR0101 | `SM-IR0101-Get-GroupPolicyObjects.ps1` | GroupPolicy | GPO inventory including status, links, and metadata. |
| IR0213 | `SM-IR0213-Get-ExchangeOnPremMailContacts.ps1` | ExchangeOnPrem | Mail contact inventory. |
| IR0214 | `SM-IR0214-Get-ExchangeOnPremDistributionLists.ps1` | ExchangeOnPrem | Distribution list inventory. |
| IR0215 | `SM-IR0215-Get-ExchangeOnPremDistributionListMembers.ps1` | ExchangeOnPrem | Distribution list membership exports. |
| IR0216 | `SM-IR0216-Get-ExchangeOnPremSharedMailboxes.ps1` | ExchangeOnPrem | Shared mailbox inventory. |
| IR0217 | `SM-IR0217-Get-ExchangeOnPremSharedMailboxPermissions.ps1` | ExchangeOnPrem | Shared mailbox permission matrix. |
| IR0218 | `SM-IR0218-Get-ExchangeOnPremResourceMailboxes.ps1` | ExchangeOnPrem | Room and equipment mailbox inventory. |
| IR0219 | `SM-IR0219-Get-ExchangeOnPremResourceMailboxBookingDelegates.ps1` | ExchangeOnPrem | Resource booking delegate and policy state. |
| IR0220 | `SM-IR0220-Get-ExchangeOnPremMailboxDelegations.ps1` | ExchangeOnPrem | Mailbox delegation matrix. |
| IR0221 | `SM-IR0221-Get-ExchangeOnPremMailboxFolderPermissions.ps1` | ExchangeOnPrem | Mailbox folder permission matrix. |
| IR0222 | `SM-IR0222-Get-ExchangeOnPremMailEnabledSecurityGroups.ps1` | ExchangeOnPrem | Mail-enabled security group inventory. |
| IR0223 | `SM-IR0223-Get-ExchangeOnPremDynamicDistributionGroups.ps1` | ExchangeOnPrem | Dynamic distribution group inventory and filters. |
| IR0224 | `SM-IR0224-Get-ExchangeOnPremInboundConnectorDetails.ps1` | ExchangeOnPrem | Inbound receive connector inventory and transport settings. |
| IR0225 | `SM-IR0225-Get-ExchangeOnPremOutlookClientVersionsFromRpcLogs.ps1` | ExchangeOnPrem | Outlook client version aggregates parsed from RPC client access logs. |
| IR0226 | `SM-IR0226-Get-ExchangeOnPremRpcLogExport.ps1` | ExchangeOnPrem | Parsed RPC client access log row export for detailed analysis. |
| IR0301 | `SM-IR0301-Get-FileServicesShares.ps1` | FileServices | Share inventory and UNC/path metadata. |
| IR0302 | `SM-IR0302-Get-FileServicesSharePermissions.ps1` | FileServices | Share ACL export. |
| IR0303 | `SM-IR0303-Get-FileServicesNtfsPermissions.ps1` | FileServices | NTFS ACL export. |
| IR0304 | `SM-IR0304-Get-FileServicesHomeDriveUsage.ps1` | FileServices | Home drive location and utilization export. |

Implemented now:

- `SM-IR0001-Get-ActiveDirectoryUsers.ps1`
- `Scope-ActiveDirectoryUsers.input.csv`
- `SM-IR0002-Get-ActiveDirectoryContacts.ps1`
- `Scope-ActiveDirectoryContacts.input.csv`
- `SM-IR0005-Get-ActiveDirectorySecurityGroups.ps1`
- `SM-IR0006-Get-ActiveDirectoryDistributionGroups.ps1`
- `SM-IR0007-Get-ActiveDirectorySecurityGroupMembers.ps1`
- `SM-IR0008-Get-ActiveDirectoryDistributionGroupMembers.ps1`
- `SM-IR0009-Get-ActiveDirectoryOrganizationalUnits.ps1`
- `SM-IR0010-Get-ActiveDirectoryUserRecursiveGroupMemberships.ps1`
- `SM-IR0011-Get-ActiveDirectoryUsersWithoutGroupMemberships.ps1`
- `SM-IR0012-Get-ActiveDirectoryGroupsWithoutMembers.ps1`
- `Scope-ActiveDirectorySecurityGroups.input.csv`
- `Scope-ActiveDirectoryDistributionGroups.input.csv`
- `Scope-ActiveDirectoryOUs.input.csv`
- `SM-IR0213-Get-ExchangeOnPremMailContacts.ps1`
- `SM-IR0214-Get-ExchangeOnPremDistributionLists.ps1`
- `SM-IR0215-Get-ExchangeOnPremDistributionListMembers.ps1`
- `SM-IR0216-Get-ExchangeOnPremSharedMailboxes.ps1`
- `SM-IR0217-Get-ExchangeOnPremSharedMailboxPermissions.ps1`
- `SM-IR0218-Get-ExchangeOnPremResourceMailboxes.ps1`
- `SM-IR0219-Get-ExchangeOnPremResourceMailboxBookingDelegates.ps1`
- `SM-IR0220-Get-ExchangeOnPremMailboxDelegations.ps1`
- `SM-IR0221-Get-ExchangeOnPremMailboxFolderPermissions.ps1`
- `SM-IR0222-Get-ExchangeOnPremMailEnabledSecurityGroups.ps1`
- `SM-IR0223-Get-ExchangeOnPremDynamicDistributionGroups.ps1`
- `SM-IR0224-Get-ExchangeOnPremInboundConnectorDetails.ps1`
- `SM-IR0225-Get-ExchangeOnPremOutlookClientVersionsFromRpcLogs.ps1`
- `SM-IR0226-Get-ExchangeOnPremRpcLogExport.ps1`
- `Scope-ExchangeOnPremMailContacts.input.csv`
- `Scope-ExchangeOnPremDistributionLists.input.csv`
- `Scope-ExchangeOnPremSharedMailboxes.input.csv`
- `Scope-ExchangeOnPremResourceMailboxes.input.csv`
- `Scope-ExchangeOnPremMailboxes.input.csv`
- `Scope-ExchangeOnPremMailEnabledSecurityGroups.input.csv`
- `Scope-ExchangeOnPremDynamicDistributionGroups.input.csv`
- `Scope-ExchangeOnPremInboundConnectors.input.csv`
- `Scope-ExchangeOnPremRpcLogs.input.csv`

## References

- [OnPrem InventoryAndReport Detailed Catalog](./README-InventoryAndReport-Catalog.md)
- [OnPrem README](../README.md)
- [SharedModule README](../../README.md)
