# InventoryAndReport Detailed Catalog

Detailed catalog for discovery/reporting scripts in `OnPrem/InventoryAndReport/`.

Operational label: **Inventory and Report**.

Current implementation status: partial. ActiveDirectory inventory baseline scripts (`IR0001`, `IR0002`, `IR0005`, `IR0006`, `IR0007`, `IR0008`, `IR0009`) and ExchangeOnPrem inventory scripts (`IR0213` through `IR0223`) are implemented; remaining scripts are planned.

## Script Contract

All planned discover scripts should:

- Be read-only (no create/update/delete actions)
- Run in native on-prem shells:
  - ActiveDirectory (`00xx`): Windows PowerShell `5.1`
  - ExchangeOnPrem (`02xx`): Exchange Management Shell (Windows PowerShell `5.1`)
- Support CSV-bounded scope via `-InputCsvPath` with validated required CSV headers (default model)
- Export deterministic CSV output for diff/baselining
- Include `Status` and `Message` per processed item
- Write a required per-run transcript log in the output folder
- Reuse shared scope CSV files when key columns overlap

## ID Ranges

- `00xx`: ActiveDirectory
- `01xx`: GroupPolicy
- `02xx`: ExchangeOnPrem
- `03xx`: FileServices

## Catalog

| ID | Script | Input CSV | Workload | Primary Output Focus | Status |
|---|---|---|---|---|---|
| IR0001 | `IR0001-Get-ActiveDirectoryUsers.ps1` | `Scope-ActiveDirectoryUsers.input.csv` | ActiveDirectory | AD user inventory and core profile fields | Implemented |
| IR0002 | `IR0002-Get-ActiveDirectoryContacts.ps1` | `Scope-ActiveDirectoryContacts.input.csv` | ActiveDirectory | AD contact inventory and mail-routing fields | Implemented |
| IR0005 | `IR0005-Get-ActiveDirectorySecurityGroups.ps1` | `Scope-ActiveDirectorySecurityGroups.input.csv` | ActiveDirectory | AD security group inventory | Implemented |
| IR0006 | `IR0006-Get-ActiveDirectoryDistributionGroups.ps1` | `Scope-ActiveDirectoryDistributionGroups.input.csv` | ActiveDirectory | AD distribution group inventory | Implemented |
| IR0007 | `IR0007-Get-ActiveDirectorySecurityGroupMembers.ps1` | `Scope-ActiveDirectorySecurityGroups.input.csv` | ActiveDirectory | AD security group membership export | Implemented |
| IR0008 | `IR0008-Get-ActiveDirectoryDistributionGroupMembers.ps1` | `Scope-ActiveDirectoryDistributionGroups.input.csv` | ActiveDirectory | AD distribution group membership export | Implemented |
| IR0009 | `IR0009-Get-ActiveDirectoryOrganizationalUnits.ps1` | `Scope-ActiveDirectoryOUs.input.csv` | ActiveDirectory | AD OU hierarchy export | Implemented |
| IR0101 | `IR0101-Get-GroupPolicyObjects.ps1` | `Scope-GroupPolicyObjects.input.csv` | GroupPolicy | GPO inventory, link targets, and enforcement state | Planned |
| IR0213 | `IR0213-Get-ExchangeOnPremMailContacts.ps1` | `Scope-ExchangeOnPremMailContacts.input.csv` | ExchangeOnPrem | Mail contact inventory | Implemented |
| IR0214 | `IR0214-Get-ExchangeOnPremDistributionLists.ps1` | `Scope-ExchangeOnPremDistributionLists.input.csv` | ExchangeOnPrem | Distribution list inventory | Implemented |
| IR0215 | `IR0215-Get-ExchangeOnPremDistributionListMembers.ps1` | `Scope-ExchangeOnPremDistributionLists.input.csv` | ExchangeOnPrem | Distribution list membership export | Implemented |
| IR0216 | `IR0216-Get-ExchangeOnPremSharedMailboxes.ps1` | `Scope-ExchangeOnPremSharedMailboxes.input.csv` | ExchangeOnPrem | Shared mailbox inventory | Implemented |
| IR0217 | `IR0217-Get-ExchangeOnPremSharedMailboxPermissions.ps1` | `Scope-ExchangeOnPremSharedMailboxes.input.csv` | ExchangeOnPrem | Shared mailbox permission matrix | Implemented |
| IR0218 | `IR0218-Get-ExchangeOnPremResourceMailboxes.ps1` | `Scope-ExchangeOnPremResourceMailboxes.input.csv` | ExchangeOnPrem | Room/equipment mailbox inventory | Implemented |
| IR0219 | `IR0219-Get-ExchangeOnPremResourceMailboxBookingDelegates.ps1` | `Scope-ExchangeOnPremResourceMailboxes.input.csv` | ExchangeOnPrem | Resource booking delegate/policy state | Implemented |
| IR0220 | `IR0220-Get-ExchangeOnPremMailboxDelegations.ps1` | `Scope-ExchangeOnPremMailboxes.input.csv` | ExchangeOnPrem | Mailbox delegation matrix | Implemented |
| IR0221 | `IR0221-Get-ExchangeOnPremMailboxFolderPermissions.ps1` | `Scope-ExchangeOnPremMailboxes.input.csv` | ExchangeOnPrem | Folder permission/delegate flag export | Implemented |
| IR0222 | `IR0222-Get-ExchangeOnPremMailEnabledSecurityGroups.ps1` | `Scope-ExchangeOnPremMailEnabledSecurityGroups.input.csv` | ExchangeOnPrem | Mail-enabled security group inventory | Implemented |
| IR0223 | `IR0223-Get-ExchangeOnPremDynamicDistributionGroups.ps1` | `Scope-ExchangeOnPremDynamicDistributionGroups.input.csv` | ExchangeOnPrem | Dynamic distribution group inventory | Implemented |
| IR0301 | `IR0301-Get-FileServicesShares.ps1` | `Scope-FileServicesShares.input.csv` | FileServices | File share inventory | Planned |
| IR0302 | `IR0302-Get-FileServicesSharePermissions.ps1` | `Scope-FileServicesShares.input.csv` | FileServices | Share ACL export | Planned |
| IR0303 | `IR0303-Get-FileServicesNtfsPermissions.ps1` | `Scope-FileServicesPaths.input.csv` | FileServices | NTFS ACL export | Planned |
| IR0304 | `IR0304-Get-FileServicesHomeDriveUsage.ps1` | `Scope-ActiveDirectoryUsers.input.csv` | FileServices | Home drive location and utilization export | Planned |

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
2. For implemented ActiveDirectory and ExchangeOnPrem scripts (`IR0001`, `IR0002`, `IR0005`, `IR0006`, `IR0007`, `IR0008`, `IR0009`, `IR0213`, `IR0214`, `IR0215`, `IR0216`, `IR0217`, `IR0218`, `IR0219`, `IR0220`, `IR0221`, `IR0222`, `IR0223`), use either CSV scope (`-InputCsvPath`) or unbounded scope (`-DiscoverAll` with optional `-SearchBase`, `-Server`, `-MaxObjects`).
3. ActiveDirectory baseline target set: `IR0001`, `IR0002`, `IR0005`, `IR0006`, `IR0007`, `IR0008`, `IR0009`
4. GroupPolicy baseline: `IR0101`
5. ExchangeOnPrem baseline: `IR0213` to `IR0223`
6. FileServices baseline: `IR0301` to `IR0304`

## Related Docs

- [OnPrem InventoryAndReport README](./README.md)
- [OnPrem README](../README.md)
- [Root README](../../README.md)
