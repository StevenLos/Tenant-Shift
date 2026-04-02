# Modify Detailed Catalog

Detailed catalog for modify/change scripts in `SharedModule/OnPrem/Modify/`.

Operational label: **Modify**.

Current implementation status: partial. ADUC modify baseline scripts (`M-ADUC-0020` through `M-ADUC-0090`) and EXOP modify scripts (`M-EXOP-0010` through `M-EXOP-0140`) are implemented; GPOL and FILE remain planned.

## Script Contract

All modify scripts should:

- Run in native on-prem shells:
  - ADUC: Windows PowerShell `5.1`
  - EXOP: Exchange Management Shell (Windows PowerShell `5.1`)
- Support `-WhatIf` and `ShouldProcess`
- Be idempotent when practical
- Validate CSV headers and required values
- Export per-record `Status` and `Message`
- Write a required per-run transcript log in the output folder
- Include rollback/remediation notes for high-impact operations

## Catalog

| ID | Script | Input Template | Workload | Primary Change Scope | Status |
|---|---|---|---|---|---|
| M-ADUC-0020 | `M-ADUC-0020-Update-ActiveDirectoryUsers.ps1` | `M-ADUC-0020-Update-ActiveDirectoryUsers.input.csv` | ADUC | User attribute updates with fold-in password reset support | Implemented |
| M-ADUC-0030 | `M-ADUC-0030-Update-ActiveDirectoryContacts.ps1` | `M-ADUC-0030-Update-ActiveDirectoryContacts.input.csv` | ADUC | Contact attribute updates | Implemented |
| M-ADUC-0040 | `M-ADUC-0040-Update-ActiveDirectorySecurityGroups.ps1` | `M-ADUC-0040-Update-ActiveDirectorySecurityGroups.input.csv` | ADUC | Security group property updates | Implemented |
| M-ADUC-0050 | `M-ADUC-0050-Update-ActiveDirectoryDistributionGroups.ps1` | `M-ADUC-0050-Update-ActiveDirectoryDistributionGroups.input.csv` | ADUC | Distribution group property updates | Implemented |
| M-ADUC-0060 | `M-ADUC-0060-Set-ActiveDirectorySecurityGroupMembers.ps1` | `M-ADUC-0060-Set-ActiveDirectorySecurityGroupMembers.input.csv` | ADUC | Add/remove security group members | Implemented |
| M-ADUC-0070 | `M-ADUC-0070-Set-ActiveDirectoryDistributionGroupMembers.ps1` | `M-ADUC-0070-Set-ActiveDirectoryDistributionGroupMembers.input.csv` | ADUC | Add/remove distribution group members | Implemented |
| M-ADUC-0080 | `M-ADUC-0080-Set-ActiveDirectoryTemporaryPasswords.ps1` | `M-ADUC-0080-Set-ActiveDirectoryTemporaryPasswords.input.csv` | ADUC | Standalone temporary password reset workflow (optional unlock/enable controls) | Implemented |
| M-ADUC-0090 | `M-ADUC-0090-Move-ActiveDirectoryObjects.ps1` | `M-ADUC-0090-Move-ActiveDirectoryObjects.input.csv` | ADUC | Move users/groups/contacts between OUs | Implemented |
| M-GPOL-0010 | `M-GPOL-0010-Set-GroupPolicyLinks.ps1` | `M-GPOL-0010-Set-GroupPolicyLinks.input.csv` | GPOL | Create/update GPO links and enforcement/order state | **Planned** |
| M-EXOP-0010 | `M-EXOP-0010-Update-ExchangeOnPremMailContacts.ps1` | `M-EXOP-0010-Update-ExchangeOnPremMailContacts.input.csv` | EXOP | Mail contact property updates | Implemented |
| M-EXOP-0020 | `M-EXOP-0020-Update-ExchangeOnPremDistributionLists.ps1` | `M-EXOP-0020-Update-ExchangeOnPremDistributionLists.input.csv` | EXOP | Distribution list property updates | Implemented |
| M-EXOP-0030 | `M-EXOP-0030-Update-ExchangeOnPremMailEnabledSecurityGroups.ps1` | `M-EXOP-0030-Update-ExchangeOnPremMailEnabledSecurityGroups.input.csv` | EXOP | Mail-enabled security group updates | Implemented |
| M-EXOP-0040 | `M-EXOP-0040-Update-ExchangeOnPremDynamicDistributionGroups.ps1` | `M-EXOP-0040-Update-ExchangeOnPremDynamicDistributionGroups.input.csv` | EXOP | Dynamic distribution group filter/property updates | Implemented |
| M-EXOP-0050 | `M-EXOP-0050-Update-ExchangeOnPremSharedMailboxes.ps1` | `M-EXOP-0050-Update-ExchangeOnPremSharedMailboxes.input.csv` | EXOP | Shared mailbox property updates | Implemented |
| M-EXOP-0060 | `M-EXOP-0060-Update-ExchangeOnPremResourceMailboxes.ps1` | `M-EXOP-0060-Update-ExchangeOnPremResourceMailboxes.input.csv` | EXOP | Resource mailbox settings updates | Implemented |
| M-EXOP-0070 | `M-EXOP-0070-Set-ExchangeOnPremDistributionListMembers.ps1` | `M-EXOP-0070-Set-ExchangeOnPremDistributionListMembers.input.csv` | EXOP | Add/remove distribution list members | Implemented |
| M-EXOP-0080 | `M-EXOP-0080-Set-ExchangeOnPremSharedMailboxPermissions.ps1` | `M-EXOP-0080-Set-ExchangeOnPremSharedMailboxPermissions.input.csv` | EXOP | Configure shared mailbox permissions | Implemented |
| M-EXOP-0090 | `M-EXOP-0090-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1` | `M-EXOP-0090-Set-ExchangeOnPremResourceMailboxBookingDelegates.input.csv` | EXOP | Configure booking delegates/policies | Implemented |
| M-EXOP-0100 | `M-EXOP-0100-Set-ExchangeOnPremMailboxDelegations.ps1` | `M-EXOP-0100-Set-ExchangeOnPremMailboxDelegations.input.csv` | EXOP | Configure mailbox delegation rights | Implemented |
| M-EXOP-0110 | `M-EXOP-0110-Set-ExchangeOnPremMailboxFolderPermissions.ps1` | `M-EXOP-0110-Set-ExchangeOnPremMailboxFolderPermissions.input.csv` | EXOP | Configure folder permissions/delegate flags | Implemented |
| M-EXOP-0120 | `M-EXOP-0120-Set-ExchangeOnPremUserMailboxForwarding.ps1` | `M-EXOP-0120-Set-ExchangeOnPremUserMailboxForwarding.input.csv` | EXOP | Set per-user mailbox forwarding mode and delivery behavior | Implemented |
| M-EXOP-0130 | `M-EXOP-0130-Convert-ExchangeOnPremMailboxToMailEnabledUser.ps1` | `M-EXOP-0130-Convert-ExchangeOnPremMailboxToMailEnabledUser.input.csv` | EXOP | Convert mailbox objects to mail-enabled users while preserving routing addresses | Implemented |
| M-EXOP-0140 | `M-EXOP-0140-Set-ExchangeOnPremMigrationWizDelegation.ps1` | `M-EXOP-0140-Set-ExchangeOnPremMigrationWizDelegation.input.csv` | EXOP | Grant/remove migration service delegation (FullAccess/SendAs) on target mailboxes | Implemented |
| M-FILE-0010 | `M-FILE-0010-Update-FileServicesShares.ps1` | `M-FILE-0010-Update-FileServicesShares.input.csv` | FILE | Share property/path updates | **Planned** |
| M-FILE-0020 | `M-FILE-0020-Set-FileServicesSharePermissions.ps1` | `M-FILE-0020-Set-FileServicesSharePermissions.input.csv` | FILE | Share ACL updates | **Planned** |
| M-FILE-0030 | `M-FILE-0030-Set-FileServicesNtfsPermissions.ps1` | `M-FILE-0030-Set-FileServicesNtfsPermissions.input.csv` | FILE | NTFS ACL updates | **Planned** |
| M-FILE-0040 | `M-FILE-0040-Update-FileServicesHomeDrives.ps1` | `M-FILE-0040-Update-FileServicesHomeDrives.input.csv` | FILE | Home drive path/quota updates | **Planned** |
| M-FILE-0050 | `M-FILE-0050-Set-FileServicesOwnerAndFullControlBySid.ps1` | `M-FILE-0050-Set-FileServicesOwnerAndFullControlBySid.input.csv` | FILE | Set owner to provided SID and grant full control to provided SID (icacls-based) | **Planned** |
| M-FILE-0060 | `M-FILE-0060-Grant-FileServicesFullControlBySid.ps1` | `M-FILE-0060-Grant-FileServicesFullControlBySid.input.csv` | FILE | Grant full control to provided SID without ownership change (icacls-based) | **Planned** |

## Safety and Sequencing Guidance

Recommended execution phases:

1. ADUC changes: `M-ADUC-0020`, `M-ADUC-0030`, `M-ADUC-0040`, `M-ADUC-0050`, `M-ADUC-0060`, `M-ADUC-0070`, `M-ADUC-0080`, `M-ADUC-0090`
2. GPOL changes: `M-GPOL-0010`
3. EXOP changes: `M-EXOP-0010` through `M-EXOP-0140`
4. FILE changes: `M-FILE-0010` through `M-FILE-0060`

High-impact change classes that should always start with `-WhatIf`:

- Membership removals (`M-ADUC-0060`, `M-ADUC-0070`, `M-EXOP-0070`)
- Permission and delegation changes (`M-EXOP-0080`, `M-EXOP-0100`, `M-EXOP-0110`, `M-EXOP-0140`, `M-FILE-0020`, `M-FILE-0030`, `M-FILE-0050`, `M-FILE-0060`)
- OU moves (`M-ADUC-0090`)
- Password resets (`M-ADUC-0080`)
- Mailbox-to-mail-user conversion (`M-EXOP-0130`)

FILE-specific planned note for `M-FILE-0050`:

- Script should support path-level targeting for files or folders, optional recursion, and preserve `-WhatIf`.
- Execution is expected to require elevated rights and privileges required to change ownership.

FILE-specific planned note for `M-FILE-0060`:

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
- [Root README](../../../README.md)
