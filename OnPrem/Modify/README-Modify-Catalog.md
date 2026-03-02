# Modify Detailed Catalog

Detailed catalog for planned modify/change scripts in `OnPrem/Modify/`.

Operational label: **Modify**.

Current implementation status: planning only. No OnPrem modify scripts are implemented yet.

## Script Contract

All planned update scripts should:

- Run on PowerShell 7+
- Support `-WhatIf` and `ShouldProcess`
- Be idempotent when practical
- Validate CSV headers and required values
- Export per-record `Status` and `Message`
- Write a required per-run transcript log in the output folder
- Include rollback/remediation notes for high-impact operations

## ID Ranges

- `00xx`: ActiveDirectory
- `01xx`: GroupPolicy
- `02xx`: ExchangeOnPrem
- `03xx`: FileServices

## Catalog

| ID | Script | Input Template | Workload | Primary Change Scope | Status |
|---|---|---|---|---|---|
| M0001 | `M0001-Update-ActiveDirectoryUsers.ps1` | `M0001-Update-ActiveDirectoryUsers.input.csv` | ActiveDirectory | User attribute updates | Planned |
| M0002 | `M0002-Update-ActiveDirectoryContacts.ps1` | `M0002-Update-ActiveDirectoryContacts.input.csv` | ActiveDirectory | Contact attribute updates | Planned |
| M0005 | `M0005-Update-ActiveDirectorySecurityGroups.ps1` | `M0005-Update-ActiveDirectorySecurityGroups.input.csv` | ActiveDirectory | Security group property updates | Planned |
| M0007 | `M0007-Set-ActiveDirectorySecurityGroupMembers.ps1` | `M0007-Set-ActiveDirectorySecurityGroupMembers.input.csv` | ActiveDirectory | Add/remove security group members | Planned |
| M0009 | `M0009-Move-ActiveDirectoryObjects.ps1` | `M0009-Move-ActiveDirectoryObjects.input.csv` | ActiveDirectory | Move users/groups/contacts between OUs | Planned |
| M0101 | `M0101-Set-GroupPolicyLinks.ps1` | `M0101-Set-GroupPolicyLinks.input.csv` | GroupPolicy | Create/update GPO links and enforcement/order state | Planned |
| M0213 | `M0213-Update-ExchangeOnPremMailContacts.ps1` | `M0213-Update-ExchangeOnPremMailContacts.input.csv` | ExchangeOnPrem | Mail contact property updates | Planned |
| M0214 | `M0214-Update-ExchangeOnPremDistributionLists.ps1` | `M0214-Update-ExchangeOnPremDistributionLists.input.csv` | ExchangeOnPrem | Distribution list property updates | Planned |
| M0215 | `M0215-Set-ExchangeOnPremDistributionListMembers.ps1` | `M0215-Set-ExchangeOnPremDistributionListMembers.input.csv` | ExchangeOnPrem | Add/remove distribution list members | Planned |
| M0216 | `M0216-Update-ExchangeOnPremSharedMailboxes.ps1` | `M0216-Update-ExchangeOnPremSharedMailboxes.input.csv` | ExchangeOnPrem | Shared mailbox property updates | Planned |
| M0217 | `M0217-Set-ExchangeOnPremSharedMailboxPermissions.ps1` | `M0217-Set-ExchangeOnPremSharedMailboxPermissions.input.csv` | ExchangeOnPrem | Configure shared mailbox permissions | Planned |
| M0218 | `M0218-Update-ExchangeOnPremResourceMailboxes.ps1` | `M0218-Update-ExchangeOnPremResourceMailboxes.input.csv` | ExchangeOnPrem | Resource mailbox settings updates | Planned |
| M0219 | `M0219-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1` | `M0219-Set-ExchangeOnPremResourceMailboxBookingDelegates.input.csv` | ExchangeOnPrem | Configure booking delegates/policies | Planned |
| M0220 | `M0220-Set-ExchangeOnPremMailboxDelegations.ps1` | `M0220-Set-ExchangeOnPremMailboxDelegations.input.csv` | ExchangeOnPrem | Configure mailbox delegation rights | Planned |
| M0221 | `M0221-Set-ExchangeOnPremMailboxFolderPermissions.ps1` | `M0221-Set-ExchangeOnPremMailboxFolderPermissions.input.csv` | ExchangeOnPrem | Configure folder permissions/delegate flags | Planned |
| M0222 | `M0222-Update-ExchangeOnPremMailEnabledSecurityGroups.ps1` | `M0222-Update-ExchangeOnPremMailEnabledSecurityGroups.input.csv` | ExchangeOnPrem | Mail-enabled security group updates | Planned |
| M0223 | `M0223-Update-ExchangeOnPremDynamicDistributionGroups.ps1` | `M0223-Update-ExchangeOnPremDynamicDistributionGroups.input.csv` | ExchangeOnPrem | Dynamic distribution group filter/property updates | Planned |
| M0301 | `M0301-Update-FileServicesShares.ps1` | `M0301-Update-FileServicesShares.input.csv` | FileServices | Share property/path updates | Planned |
| M0302 | `M0302-Set-FileServicesSharePermissions.ps1` | `M0302-Set-FileServicesSharePermissions.input.csv` | FileServices | Share ACL updates | Planned |
| M0303 | `M0303-Set-FileServicesNtfsPermissions.ps1` | `M0303-Set-FileServicesNtfsPermissions.input.csv` | FileServices | NTFS ACL updates | Planned |
| M0304 | `M0304-Update-FileServicesHomeDrives.ps1` | `M0304-Update-FileServicesHomeDrives.input.csv` | FileServices | Home drive path/quota updates | Planned |
| M0305 | `M0305-Set-FileServicesOwnerAndFullControlBySid.ps1` | `M0305-Set-FileServicesOwnerAndFullControlBySid.input.csv` | FileServices | Set owner to provided SID and grant full control to provided SID (icacls-based) | Planned |
| M0306 | `M0306-Grant-FileServicesFullControlBySid.ps1` | `M0306-Grant-FileServicesFullControlBySid.input.csv` | FileServices | Grant full control to provided SID without ownership change (icacls-based) | Planned |

## Safety and Sequencing Guidance

Recommended execution phases:

1. ActiveDirectory changes: `M0001`, `M0002`, `M0005`, `M0007`, `M0009`
2. GroupPolicy changes: `M0101`
3. ExchangeOnPrem changes: `M0213` to `M0223`
4. FileServices changes: `M0301` to `M0306`

High-impact change classes that should always start with `-WhatIf`:

- Membership removals (`M0007`, `M0215`)
- Permission and delegation changes (`M0217`, `M0220`, `M0221`, `M0302`, `M0303`, `M0305`, `M0306`)
- OU moves (`M0009`)

FileServices-specific planned note for `M0305`:

- Script should support path-level targeting for files or folders, optional recursion, and preserve `-WhatIf`.
- Execution is expected to require elevated rights and privileges required to change ownership.

FileServices-specific planned note for `M0306`:

- Script should support path-level targeting for files or folders, optional recursion, and preserve `-WhatIf`.

## Standard Result Columns

Recommended baseline columns:

- `RowNumber`
- `PrimaryKey`
- `Action`
- `Status`
- `Message`

## Related Docs

- [OnPrem Modify README](./README.md)
- [OnPrem README](../README.md)
- [Root README](../../README.md)


