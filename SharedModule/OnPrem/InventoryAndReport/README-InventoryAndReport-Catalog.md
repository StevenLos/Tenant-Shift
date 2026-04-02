# InventoryAndReport Detailed Catalog

Detailed catalog for discovery/reporting scripts in `SharedModule/OnPrem/InventoryAndReport/`.

Operational label: **Discover**.

Current implementation status: partial. ADUC discover baseline scripts (`D-ADUC-0010` through `D-ADUC-0120`) and EXOP discover scripts (`D-EXOP-0010` through `D-EXOP-0170`) are implemented; GPOL and FILE remain planned.

## Script Contract

All discover scripts should:

- Be read-only (no create/update/delete actions)
- Run in native on-prem shells:
  - ADUC: Windows PowerShell `5.1`
  - EXOP: Exchange Management Shell (Windows PowerShell `5.1`)
- Support CSV-bounded scope via `-InputCsvPath` with validated required CSV headers (default model)
- Export deterministic CSV output for diff/baselining
- Include `Status` and `Message` per processed item
- Write a required per-run transcript log in the output folder
- Reuse shared scope CSV files when key columns overlap

## Catalog

| ID | Script | Input CSV | Workload | Primary Output Focus | Status |
|---|---|---|---|---|---|
| D-ADUC-0010 | `D-ADUC-0010-Get-ActiveDirectoryOrganizationalUnits.ps1` | `Scope-ActiveDirectoryOUs.input.csv` | ADUC | AD OU hierarchy export | Implemented |
| D-ADUC-0020 | `D-ADUC-0020-Get-ActiveDirectoryUsers.ps1` | `Scope-ActiveDirectoryUsers.input.csv` | ADUC | AD user inventory and core profile fields | Implemented |
| D-ADUC-0030 | `D-ADUC-0030-Get-ActiveDirectoryContacts.ps1` | `Scope-ActiveDirectoryContacts.input.csv` | ADUC | AD contact inventory and mail-routing fields | Implemented |
| D-ADUC-0040 | `D-ADUC-0040-Get-ActiveDirectorySecurityGroups.ps1` | `Scope-ActiveDirectorySecurityGroups.input.csv` | ADUC | AD security group inventory | Implemented |
| D-ADUC-0050 | `D-ADUC-0050-Get-ActiveDirectoryDistributionGroups.ps1` | `Scope-ActiveDirectoryDistributionGroups.input.csv` | ADUC | AD distribution group inventory | Implemented |
| D-ADUC-0060 | `D-ADUC-0060-Get-ActiveDirectorySecurityGroupMembers.ps1` | `Scope-ActiveDirectorySecurityGroups.input.csv` | ADUC | AD security group membership export | Implemented |
| D-ADUC-0070 | `D-ADUC-0070-Get-ActiveDirectoryDistributionGroupMembers.ps1` | `Scope-ActiveDirectoryDistributionGroups.input.csv` | ADUC | AD distribution group membership export | Implemented |
| D-ADUC-0100 | `D-ADUC-0100-Get-ActiveDirectoryUserRecursiveGroupMemberships.ps1` | `Scope-ActiveDirectoryUsers.input.csv` | ADUC | Recursive group memberships by user (including primary group) | Implemented |
| D-ADUC-0110 | `D-ADUC-0110-Get-ActiveDirectoryUsersWithoutGroupMemberships.ps1` | `Scope-ActiveDirectoryUsers.input.csv` | ADUC | Users with no recursive group memberships | Implemented |
| D-ADUC-0120 | `D-ADUC-0120-Get-ActiveDirectoryGroupsWithoutMembers.ps1` | `Scope-ActiveDirectorySecurityGroups.input.csv` | ADUC | Groups with no direct members and no primary-group users | Implemented |
| D-GPOL-0010 | `D-GPOL-0010-Get-GroupPolicyObjects.ps1` | `Scope-GroupPolicyObjects.input.csv` | GPOL | GPO inventory, link targets, and enforcement state | **Planned** |
| D-EXOP-0010 | `D-EXOP-0010-Get-ExchangeOnPremMailContacts.ps1` | `Scope-ExchangeOnPremMailContacts.input.csv` | EXOP | Mail contact inventory | Implemented |
| D-EXOP-0020 | `D-EXOP-0020-Get-ExchangeOnPremDistributionLists.ps1` | `Scope-ExchangeOnPremDistributionLists.input.csv` | EXOP | Distribution list inventory | Implemented |
| D-EXOP-0030 | `D-EXOP-0030-Get-ExchangeOnPremMailEnabledSecurityGroups.ps1` | `Scope-ExchangeOnPremMailEnabledSecurityGroups.input.csv` | EXOP | Mail-enabled security group inventory | Implemented |
| D-EXOP-0040 | `D-EXOP-0040-Get-ExchangeOnPremDynamicDistributionGroups.ps1` | `Scope-ExchangeOnPremDynamicDistributionGroups.input.csv` | EXOP | Dynamic distribution group inventory | Implemented |
| D-EXOP-0050 | `D-EXOP-0050-Get-ExchangeOnPremSharedMailboxes.ps1` | `Scope-ExchangeOnPremSharedMailboxes.input.csv` | EXOP | Shared mailbox inventory | Implemented |
| D-EXOP-0060 | `D-EXOP-0060-Get-ExchangeOnPremResourceMailboxes.ps1` | `Scope-ExchangeOnPremResourceMailboxes.input.csv` | EXOP | Room/equipment mailbox inventory | Implemented |
| D-EXOP-0070 | `D-EXOP-0070-Get-ExchangeOnPremDistributionListMembers.ps1` | `Scope-ExchangeOnPremDistributionLists.input.csv` | EXOP | Distribution list membership export | Implemented |
| D-EXOP-0080 | `D-EXOP-0080-Get-ExchangeOnPremSharedMailboxPermissions.ps1` | `Scope-ExchangeOnPremSharedMailboxes.input.csv` | EXOP | Shared mailbox permission matrix | Implemented |
| D-EXOP-0090 | `D-EXOP-0090-Get-ExchangeOnPremResourceMailboxBookingDelegates.ps1` | `Scope-ExchangeOnPremResourceMailboxes.input.csv` | EXOP | Resource booking delegate/policy state | Implemented |
| D-EXOP-0100 | `D-EXOP-0100-Get-ExchangeOnPremMailboxDelegations.ps1` | `Scope-ExchangeOnPremMailboxes.input.csv` | EXOP | Mailbox delegation matrix | Implemented |
| D-EXOP-0110 | `D-EXOP-0110-Get-ExchangeOnPremMailboxFolderPermissions.ps1` | `Scope-ExchangeOnPremMailboxes.input.csv` | EXOP | Folder permission/delegate flag export | Implemented |
| D-EXOP-0150 | `D-EXOP-0150-Get-ExchangeOnPremInboundConnectorDetails.ps1` | `Scope-ExchangeOnPremInboundConnectors.input.csv` | EXOP | Inbound receive connector inventory and transport settings | Implemented |
| D-EXOP-0160 | `D-EXOP-0160-Get-ExchangeOnPremOutlookClientVersionsFromRpcLogs.ps1` | `Scope-ExchangeOnPremRpcLogs.input.csv` | EXOP | Outlook client version aggregates parsed from RPC client access logs | Implemented |
| D-EXOP-0170 | `D-EXOP-0170-Get-ExchangeOnPremRpcLogExport.ps1` | `Scope-ExchangeOnPremRpcLogs.input.csv` | EXOP | Parsed RPC client access log row export for detailed analysis | Implemented |
| D-FILE-0010 | `D-FILE-0010-Get-FileServicesShares.ps1` | `Scope-FileServicesShares.input.csv` | FILE | File share inventory | **Planned** |
| D-FILE-0020 | `D-FILE-0020-Get-FileServicesSharePermissions.ps1` | `Scope-FileServicesShares.input.csv` | FILE | Share ACL export | **Planned** |
| D-FILE-0030 | `D-FILE-0030-Get-FileServicesNtfsPermissions.ps1` | `Scope-FileServicesPaths.input.csv` | FILE | NTFS ACL export | **Planned** |
| D-FILE-0040 | `D-FILE-0040-Get-FileServicesHomeDriveUsage.ps1` | `Scope-ActiveDirectoryUsers.input.csv` | FILE | Home drive location and utilization export | **Planned** |

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
3. ADUC baseline: `D-ADUC-0010`, `D-ADUC-0020`, `D-ADUC-0030`, `D-ADUC-0040`, `D-ADUC-0050`, `D-ADUC-0060`, `D-ADUC-0070`, `D-ADUC-0100`, `D-ADUC-0110`, `D-ADUC-0120`
4. GPOL baseline: `D-GPOL-0010`
5. EXOP baseline: `D-EXOP-0010` through `D-EXOP-0170`
6. FILE baseline: `D-FILE-0010` through `D-FILE-0040`

## Related Docs

- [OnPrem InventoryAndReport README](./README.md)
- [OnPrem README](../README.md)
- [Root README](../../../README.md)
