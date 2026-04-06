# OnPrem Discover

`TenantShift/OnPrem/Discover` is for planned read-only inventory and reporting scripts targeting on-prem workloads.

Operational label: **Discover**.

Current status: ActiveDirectory inventory baseline is implemented with `D-ADUC-0020-Get-ActiveDirectoryUsers.ps1`, `D-ADUC-0030-Get-ActiveDirectoryContacts.ps1`, `D-ADUC-0040-Get-ActiveDirectorySecurityGroups.ps1`, `D-ADUC-0050-Get-ActiveDirectoryDistributionGroups.ps1`, `D-ADUC-0060-Get-ActiveDirectorySecurityGroupMembers.ps1`, `D-ADUC-0070-Get-ActiveDirectoryDistributionGroupMembers.ps1`, `D-ADUC-0010-Get-ActiveDirectoryOrganizationalUnits.ps1`, `D-ADUC-0100-Get-ActiveDirectoryUserRecursiveGroupMemberships.ps1`, `D-ADUC-0110-Get-ActiveDirectoryUsersWithoutGroupMemberships.ps1`, and `D-ADUC-0120-Get-ActiveDirectoryGroupsWithoutMembers.ps1`. ExchangeOnPrem inventory scripts are implemented with `D-EXOP-0010-Get-ExchangeOnPremMailContacts.ps1`, `D-EXOP-0020-Get-ExchangeOnPremDistributionLists.ps1`, `D-EXOP-0070-Get-ExchangeOnPremDistributionListMembers.ps1`, `D-EXOP-0050-Get-ExchangeOnPremSharedMailboxes.ps1`, `D-EXOP-0080-Get-ExchangeOnPremSharedMailboxPermissions.ps1`, `D-EXOP-0060-Get-ExchangeOnPremResourceMailboxes.ps1`, `D-EXOP-0090-Get-ExchangeOnPremResourceMailboxBookingDelegates.ps1`, `D-EXOP-0100-Get-ExchangeOnPremMailboxDelegations.ps1`, `D-EXOP-0110-Get-ExchangeOnPremMailboxFolderPermissions.ps1`, `D-EXOP-0030-Get-ExchangeOnPremMailEnabledSecurityGroups.ps1`, `D-EXOP-0040-Get-ExchangeOnPremDynamicDistributionGroups.ps1`, `D-EXOP-0150-Get-ExchangeOnPremInboundConnectorDetails.ps1`, `D-EXOP-0160-Get-ExchangeOnPremOutlookClientVersionsFromRpcLogs.ps1`, and `D-EXOP-0170-Get-ExchangeOnPremRpcLogExport.ps1`; remaining scripts are planned.

## Purpose

Use this folder for:

- Baseline snapshots before and after provision/modify runs
- Audit and compliance exports
- Object and permissions discovery for AD, ExchangeOnPrem, and FileServices

Do not use this folder for:

- Creation, modification, or deletion operations

## Naming Standard

- Script: `D-<WW>-<NNNN>-<Action>-<Target>.ps1`
- Input CSV: shared `Scope-*.input.csv` (preferred) or script-specific `D-<WW>-<NNNN>-<Action>-<Target>.input.csv` when needed
- Output pattern: `Results_D-<WW>-<NNNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Transcript log pattern: `Transcript_D-<WW>-<NNNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`

## Runtime Assumptions

- ActiveDirectory inventory scripts (ADUC) run in Windows PowerShell `5.1`.
- ExchangeOnPrem inventory scripts (EXOP) run in Exchange Management Shell (Windows PowerShell `5.1`).

## Discovery Scope Modes

- Default discovery model is CSV scope input via `-InputCsvPath`.
- Implemented ActiveDirectory and ExchangeOnPrem inventory scripts also support unbounded discovery via `-DiscoverAll`.
- `-DiscoverAll` scope controls are script-specific:
  - Directory/object-focused scripts use controls such as `-SearchBase`, `-Server`, and `-MaxObjects`.
  - RPC-log-focused scripts (`D-EXOP-0160`, `D-EXOP-0170`) use controls such as `-LogPath`, `-LookbackDays`, and `-MaxObjects`.

Workload code allocation (`WW` in `D-<WW>-<NNNN>`):

- `ADUC`: ActiveDirectory
- `GPOL`: GroupPolicy
- `EXOP`: ExchangeOnPrem
- `FILE`: FileServices

## Script Matrix (Current Status)

| ID | Script | Workload | Primary Output Focus |
|---|---|---|---|
| D-ADUC-0010 | `D-ADUC-0010-Get-ActiveDirectoryOrganizationalUnits.ps1` | ActiveDirectory | OU hierarchy and delegation boundary inventory. |
| D-ADUC-0020 | `D-ADUC-0020-Get-ActiveDirectoryUsers.ps1` | ActiveDirectory | AD user inventory and key identity attributes. |
| D-ADUC-0030 | `D-ADUC-0030-Get-ActiveDirectoryContacts.ps1` | ActiveDirectory | AD contact inventory and mail routing attributes. |
| D-ADUC-0040 | `D-ADUC-0040-Get-ActiveDirectorySecurityGroups.ps1` | ActiveDirectory | AD security group inventory and scope/type metadata. |
| D-ADUC-0050 | `D-ADUC-0050-Get-ActiveDirectoryDistributionGroups.ps1` | ActiveDirectory | AD distribution group inventory and scope/type metadata. |
| D-ADUC-0060 | `D-ADUC-0060-Get-ActiveDirectorySecurityGroupMembers.ps1` | ActiveDirectory | Group membership exports for AD security groups. |
| D-ADUC-0070 | `D-ADUC-0070-Get-ActiveDirectoryDistributionGroupMembers.ps1` | ActiveDirectory | Group membership exports for AD distribution groups. |
| D-ADUC-0100 | `D-ADUC-0100-Get-ActiveDirectoryUserRecursiveGroupMemberships.ps1` | ActiveDirectory | Recursive group memberships per user, including primary group coverage. |
| D-ADUC-0110 | `D-ADUC-0110-Get-ActiveDirectoryUsersWithoutGroupMemberships.ps1` | ActiveDirectory | AD users with no recursive group memberships. |
| D-ADUC-0120 | `D-ADUC-0120-Get-ActiveDirectoryGroupsWithoutMembers.ps1` | ActiveDirectory | AD groups with no direct members and no primary-group users. |
| D-GPOL-0010 | `D-GPOL-0010-Get-GroupPolicyObjects.ps1` | GroupPolicy | GPO inventory including status, links, and metadata. (**Planned**) |
| D-EXOP-0010 | `D-EXOP-0010-Get-ExchangeOnPremMailContacts.ps1` | ExchangeOnPrem | Mail contact inventory. |
| D-EXOP-0020 | `D-EXOP-0020-Get-ExchangeOnPremDistributionLists.ps1` | ExchangeOnPrem | Distribution list inventory. |
| D-EXOP-0030 | `D-EXOP-0030-Get-ExchangeOnPremMailEnabledSecurityGroups.ps1` | ExchangeOnPrem | Mail-enabled security group inventory. |
| D-EXOP-0040 | `D-EXOP-0040-Get-ExchangeOnPremDynamicDistributionGroups.ps1` | ExchangeOnPrem | Dynamic distribution group inventory and filters. |
| D-EXOP-0050 | `D-EXOP-0050-Get-ExchangeOnPremSharedMailboxes.ps1` | ExchangeOnPrem | Shared mailbox inventory. |
| D-EXOP-0060 | `D-EXOP-0060-Get-ExchangeOnPremResourceMailboxes.ps1` | ExchangeOnPrem | Room and equipment mailbox inventory. |
| D-EXOP-0070 | `D-EXOP-0070-Get-ExchangeOnPremDistributionListMembers.ps1` | ExchangeOnPrem | Distribution list membership exports. |
| D-EXOP-0080 | `D-EXOP-0080-Get-ExchangeOnPremSharedMailboxPermissions.ps1` | ExchangeOnPrem | Shared mailbox permission matrix. |
| D-EXOP-0090 | `D-EXOP-0090-Get-ExchangeOnPremResourceMailboxBookingDelegates.ps1` | ExchangeOnPrem | Resource booking delegate and policy state. |
| D-EXOP-0100 | `D-EXOP-0100-Get-ExchangeOnPremMailboxDelegations.ps1` | ExchangeOnPrem | Mailbox delegation matrix. |
| D-EXOP-0110 | `D-EXOP-0110-Get-ExchangeOnPremMailboxFolderPermissions.ps1` | ExchangeOnPrem | Mailbox folder permission matrix. |
| D-EXOP-0150 | `D-EXOP-0150-Get-ExchangeOnPremInboundConnectorDetails.ps1` | ExchangeOnPrem | Inbound receive connector inventory and transport settings. |
| D-EXOP-0160 | `D-EXOP-0160-Get-ExchangeOnPremOutlookClientVersionsFromRpcLogs.ps1` | ExchangeOnPrem | Outlook client version aggregates parsed from RPC client access logs. |
| D-EXOP-0170 | `D-EXOP-0170-Get-ExchangeOnPremRpcLogExport.ps1` | ExchangeOnPrem | Parsed RPC client access log row export for detailed analysis. |
| D-FILE-0010 | `D-FILE-0010-Get-FileServicesShares.ps1` | FileServices | Share inventory and UNC/path metadata. (**Planned**) |
| D-FILE-0020 | `D-FILE-0020-Get-FileServicesSharePermissions.ps1` | FileServices | Share ACL export. (**Planned**) |
| D-FILE-0030 | `D-FILE-0030-Get-FileServicesNtfsPermissions.ps1` | FileServices | NTFS ACL export. (**Planned**) |
| D-FILE-0040 | `D-FILE-0040-Get-FileServicesHomeDriveUsage.ps1` | FileServices | Home drive location and utilization export. (**Planned**) |

Implemented now:

- `D-ADUC-0010-Get-ActiveDirectoryOrganizationalUnits.ps1`
- `D-ADUC-0020-Get-ActiveDirectoryUsers.ps1`
- `Scope-ActiveDirectoryUsers.input.csv`
- `D-ADUC-0030-Get-ActiveDirectoryContacts.ps1`
- `Scope-ActiveDirectoryContacts.input.csv`
- `D-ADUC-0040-Get-ActiveDirectorySecurityGroups.ps1`
- `D-ADUC-0050-Get-ActiveDirectoryDistributionGroups.ps1`
- `D-ADUC-0060-Get-ActiveDirectorySecurityGroupMembers.ps1`
- `D-ADUC-0070-Get-ActiveDirectoryDistributionGroupMembers.ps1`
- `D-ADUC-0100-Get-ActiveDirectoryUserRecursiveGroupMemberships.ps1`
- `D-ADUC-0110-Get-ActiveDirectoryUsersWithoutGroupMemberships.ps1`
- `D-ADUC-0120-Get-ActiveDirectoryGroupsWithoutMembers.ps1`
- `Scope-ActiveDirectorySecurityGroups.input.csv`
- `Scope-ActiveDirectoryDistributionGroups.input.csv`
- `Scope-ActiveDirectoryOUs.input.csv`
- `D-EXOP-0010-Get-ExchangeOnPremMailContacts.ps1`
- `D-EXOP-0020-Get-ExchangeOnPremDistributionLists.ps1`
- `D-EXOP-0030-Get-ExchangeOnPremMailEnabledSecurityGroups.ps1`
- `D-EXOP-0040-Get-ExchangeOnPremDynamicDistributionGroups.ps1`
- `D-EXOP-0050-Get-ExchangeOnPremSharedMailboxes.ps1`
- `D-EXOP-0060-Get-ExchangeOnPremResourceMailboxes.ps1`
- `D-EXOP-0070-Get-ExchangeOnPremDistributionListMembers.ps1`
- `D-EXOP-0080-Get-ExchangeOnPremSharedMailboxPermissions.ps1`
- `D-EXOP-0090-Get-ExchangeOnPremResourceMailboxBookingDelegates.ps1`
- `D-EXOP-0100-Get-ExchangeOnPremMailboxDelegations.ps1`
- `D-EXOP-0110-Get-ExchangeOnPremMailboxFolderPermissions.ps1`
- `D-EXOP-0150-Get-ExchangeOnPremInboundConnectorDetails.ps1`
- `D-EXOP-0160-Get-ExchangeOnPremOutlookClientVersionsFromRpcLogs.ps1`
- `D-EXOP-0170-Get-ExchangeOnPremRpcLogExport.ps1`
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

- [OnPrem Discover Detailed Catalog](./README-Discover-Catalog.md)
- [OnPrem README](../README.md)
- [Root README](../../../README.md)
- [Operator Runbook](./RUNBOOK-Discover.md)
