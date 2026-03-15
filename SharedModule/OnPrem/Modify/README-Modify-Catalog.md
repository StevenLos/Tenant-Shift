# Modify Detailed Catalog

Detailed catalog for modify/change scripts in `SharedModule/OnPrem/Modify/`.

Operational label: **Modify**.

Current implementation status: partial. ActiveDirectory modify baseline scripts (`M0001`, `M0002`, `M0005`, `M0006`, `M0007`, `M0008`, `M0009`, `M0010`) and ExchangeOnPrem modify scripts (`M0213` through `M0226`) are implemented; remaining scripts are planned.

## Script Contract

All planned update scripts should:

- Run in native on-prem shells:
  - ActiveDirectory (`00xx`): Windows PowerShell `5.1`
  - ExchangeOnPrem (`02xx`): Exchange Management Shell (Windows PowerShell `5.1`)
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
| M0001 | `SM-M0001-Update-ActiveDirectoryUsers.ps1` | `SM-M0001-Update-ActiveDirectoryUsers.input.csv` | ActiveDirectory | User attribute updates with fold-in password reset support | Implemented |
| M0002 | `SM-M0002-Update-ActiveDirectoryContacts.ps1` | `SM-M0002-Update-ActiveDirectoryContacts.input.csv` | ActiveDirectory | Contact attribute updates | Implemented |
| M0005 | `SM-M0005-Update-ActiveDirectorySecurityGroups.ps1` | `SM-M0005-Update-ActiveDirectorySecurityGroups.input.csv` | ActiveDirectory | Security group property updates | Implemented |
| M0006 | `SM-M0006-Update-ActiveDirectoryDistributionGroups.ps1` | `SM-M0006-Update-ActiveDirectoryDistributionGroups.input.csv` | ActiveDirectory | Distribution group property updates | Implemented |
| M0007 | `SM-M0007-Set-ActiveDirectorySecurityGroupMembers.ps1` | `SM-M0007-Set-ActiveDirectorySecurityGroupMembers.input.csv` | ActiveDirectory | Add/remove security group members | Implemented |
| M0008 | `SM-M0008-Set-ActiveDirectoryDistributionGroupMembers.ps1` | `SM-M0008-Set-ActiveDirectoryDistributionGroupMembers.input.csv` | ActiveDirectory | Add/remove distribution group members | Implemented |
| M0009 | `SM-M0009-Move-ActiveDirectoryObjects.ps1` | `SM-M0009-Move-ActiveDirectoryObjects.input.csv` | ActiveDirectory | Move users/groups/contacts between OUs | Implemented |
| M0010 | `SM-M0010-Set-ActiveDirectoryTemporaryPasswords.ps1` | `SM-M0010-Set-ActiveDirectoryTemporaryPasswords.input.csv` | ActiveDirectory | Standalone temporary password reset workflow (optional unlock/enable controls) | Implemented |
| M0101 | `SM-M0101-Set-GroupPolicyLinks.ps1` | `SM-M0101-Set-GroupPolicyLinks.input.csv` | GroupPolicy | Create/update GPO links and enforcement/order state | Planned |
| M0213 | `SM-M0213-Update-ExchangeOnPremMailContacts.ps1` | `SM-M0213-Update-ExchangeOnPremMailContacts.input.csv` | ExchangeOnPrem | Mail contact property updates | Implemented |
| M0214 | `SM-M0214-Update-ExchangeOnPremDistributionLists.ps1` | `SM-M0214-Update-ExchangeOnPremDistributionLists.input.csv` | ExchangeOnPrem | Distribution list property updates | Implemented |
| M0215 | `SM-M0215-Set-ExchangeOnPremDistributionListMembers.ps1` | `SM-M0215-Set-ExchangeOnPremDistributionListMembers.input.csv` | ExchangeOnPrem | Add/remove distribution list members | Implemented |
| M0216 | `SM-M0216-Update-ExchangeOnPremSharedMailboxes.ps1` | `SM-M0216-Update-ExchangeOnPremSharedMailboxes.input.csv` | ExchangeOnPrem | Shared mailbox property updates | Implemented |
| M0217 | `SM-M0217-Set-ExchangeOnPremSharedMailboxPermissions.ps1` | `SM-M0217-Set-ExchangeOnPremSharedMailboxPermissions.input.csv` | ExchangeOnPrem | Configure shared mailbox permissions | Implemented |
| M0218 | `SM-M0218-Update-ExchangeOnPremResourceMailboxes.ps1` | `SM-M0218-Update-ExchangeOnPremResourceMailboxes.input.csv` | ExchangeOnPrem | Resource mailbox settings updates | Implemented |
| M0219 | `SM-M0219-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1` | `SM-M0219-Set-ExchangeOnPremResourceMailboxBookingDelegates.input.csv` | ExchangeOnPrem | Configure booking delegates/policies | Implemented |
| M0220 | `SM-M0220-Set-ExchangeOnPremMailboxDelegations.ps1` | `SM-M0220-Set-ExchangeOnPremMailboxDelegations.input.csv` | ExchangeOnPrem | Configure mailbox delegation rights | Implemented |
| M0221 | `SM-M0221-Set-ExchangeOnPremMailboxFolderPermissions.ps1` | `SM-M0221-Set-ExchangeOnPremMailboxFolderPermissions.input.csv` | ExchangeOnPrem | Configure folder permissions/delegate flags | Implemented |
| M0222 | `SM-M0222-Update-ExchangeOnPremMailEnabledSecurityGroups.ps1` | `SM-M0222-Update-ExchangeOnPremMailEnabledSecurityGroups.input.csv` | ExchangeOnPrem | Mail-enabled security group updates | Implemented |
| M0223 | `SM-M0223-Update-ExchangeOnPremDynamicDistributionGroups.ps1` | `SM-M0223-Update-ExchangeOnPremDynamicDistributionGroups.input.csv` | ExchangeOnPrem | Dynamic distribution group filter/property updates | Implemented |
| M0224 | `SM-M0224-Set-ExchangeOnPremUserMailboxForwarding.ps1` | `SM-M0224-Set-ExchangeOnPremUserMailboxForwarding.input.csv` | ExchangeOnPrem | Set per-user mailbox forwarding mode and delivery behavior | Implemented |
| M0225 | `SM-M0225-Convert-ExchangeOnPremMailboxToMailEnabledUser.ps1` | `SM-M0225-Convert-ExchangeOnPremMailboxToMailEnabledUser.input.csv` | ExchangeOnPrem | Convert mailbox objects to mail-enabled users while preserving routing addresses | Implemented |
| M0226 | `SM-M0226-Set-ExchangeOnPremMigrationWizDelegation.ps1` | `SM-M0226-Set-ExchangeOnPremMigrationWizDelegation.input.csv` | ExchangeOnPrem | Grant/remove migration service delegation (FullAccess/SendAs) on target mailboxes | Implemented |
| M0301 | `SM-M0301-Update-FileServicesShares.ps1` | `SM-M0301-Update-FileServicesShares.input.csv` | FileServices | Share property/path updates | Planned |
| M0302 | `SM-M0302-Set-FileServicesSharePermissions.ps1` | `SM-M0302-Set-FileServicesSharePermissions.input.csv` | FileServices | Share ACL updates | Planned |
| M0303 | `SM-M0303-Set-FileServicesNtfsPermissions.ps1` | `SM-M0303-Set-FileServicesNtfsPermissions.input.csv` | FileServices | NTFS ACL updates | Planned |
| M0304 | `SM-M0304-Update-FileServicesHomeDrives.ps1` | `SM-M0304-Update-FileServicesHomeDrives.input.csv` | FileServices | Home drive path/quota updates | Planned |
| M0305 | `SM-M0305-Set-FileServicesOwnerAndFullControlBySid.ps1` | `SM-M0305-Set-FileServicesOwnerAndFullControlBySid.input.csv` | FileServices | Set owner to provided SID and grant full control to provided SID (icacls-based) | Planned |
| M0306 | `SM-M0306-Grant-FileServicesFullControlBySid.ps1` | `SM-M0306-Grant-FileServicesFullControlBySid.input.csv` | FileServices | Grant full control to provided SID without ownership change (icacls-based) | Planned |

## Safety and Sequencing Guidance

Recommended execution phases:

1. ActiveDirectory changes: `M0001`, `M0002`, `M0005`, `M0006`, `M0007`, `M0008`, `M0009`, `M0010`
2. GroupPolicy changes: `M0101`
3. ExchangeOnPrem changes: `M0213` to `M0226`
4. FileServices changes: `M0301` to `M0306`

High-impact change classes that should always start with `-WhatIf`:

- Membership removals (`M0007`, `M0008`, `M0215`)
- Permission and delegation changes (`M0217`, `M0220`, `M0221`, `M0226`, `M0302`, `M0303`, `M0305`, `M0306`)
- OU moves (`M0009`)
- Password resets (`M0010`)
- Mailbox-to-mail-user conversion (`M0225`)

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
- [SharedModule README](../../README.md)
