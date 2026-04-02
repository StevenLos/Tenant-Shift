# OnPrem Modify

`SharedModule/OnPrem/Modify` is for planned update/change scripts targeting existing on-prem objects.

Operational label: **Modify**.

Current status: ActiveDirectory modify baseline is implemented with `M-ADUC-0020-Update-ActiveDirectoryUsers.ps1`, `M-ADUC-0030-Update-ActiveDirectoryContacts.ps1`, `M-ADUC-0040-Update-ActiveDirectorySecurityGroups.ps1`, `M-ADUC-0050-Update-ActiveDirectoryDistributionGroups.ps1`, `M-ADUC-0060-Set-ActiveDirectorySecurityGroupMembers.ps1`, `M-ADUC-0070-Set-ActiveDirectoryDistributionGroupMembers.ps1`, `M-ADUC-0090-Move-ActiveDirectoryObjects.ps1`, and `M-ADUC-0080-Set-ActiveDirectoryTemporaryPasswords.ps1`. ExchangeOnPrem modify scripts are implemented with `M-EXOP-0010-Update-ExchangeOnPremMailContacts.ps1`, `M-EXOP-0020-Update-ExchangeOnPremDistributionLists.ps1`, `M-EXOP-0070-Set-ExchangeOnPremDistributionListMembers.ps1`, `M-EXOP-0050-Update-ExchangeOnPremSharedMailboxes.ps1`, `M-EXOP-0080-Set-ExchangeOnPremSharedMailboxPermissions.ps1`, `M-EXOP-0060-Update-ExchangeOnPremResourceMailboxes.ps1`, `M-EXOP-0090-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1`, `M-EXOP-0100-Set-ExchangeOnPremMailboxDelegations.ps1`, `M-EXOP-0110-Set-ExchangeOnPremMailboxFolderPermissions.ps1`, `M-EXOP-0030-Update-ExchangeOnPremMailEnabledSecurityGroups.ps1`, `M-EXOP-0040-Update-ExchangeOnPremDynamicDistributionGroups.ps1`, `M-EXOP-0120-Set-ExchangeOnPremUserMailboxForwarding.ps1`, `M-EXOP-0130-Convert-ExchangeOnPremMailboxToMailEnabledUser.ps1`, and `M-EXOP-0140-Set-ExchangeOnPremMigrationWizDelegation.ps1`; remaining scripts are planned.

## Purpose

Use this folder for controlled change operations after initial provisioning:

- Attribute and policy updates
- Membership and permission changes
- Lifecycle changes for existing objects

Do not use this folder for:

- Initial object creation (use `SharedModule/OnPrem/Provision`)
- Read-only reporting (use `SharedModule/OnPrem/InventoryAndReport`)

## Naming Standard

- Script: `M-<WW>-<NNNN>-<Action>-<Target>.ps1`
- Input template: `M-<WW>-<NNNN>-<Action>-<Target>.input.csv`
- Output pattern: `Results_M-<WW>-<NNNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Transcript log pattern: `Transcript_M-<WW>-<NNNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`

## Runtime Assumptions

- ActiveDirectory modify scripts (ADUC) run in Windows PowerShell `5.1`.
- ExchangeOnPrem modify scripts (EXOP) run in Exchange Management Shell (Windows PowerShell `5.1`).

Workload code allocation (`WW` in `M-<WW>-<NNNN>`):

- `ADUC`: ActiveDirectory
- `GPOL`: GroupPolicy
- `EXOP`: ExchangeOnPrem
- `FILE`: FileServices

## Script Matrix (Current Status)

| ID | Script | Workload | Primary Change Scope |
|---|---|---|---|
| M-ADUC-0020 | `M-ADUC-0020-Update-ActiveDirectoryUsers.ps1` | ActiveDirectory | User attribute updates (display, UPN, manager, office, phone), including fold-in password reset support. |
| M-ADUC-0030 | `M-ADUC-0030-Update-ActiveDirectoryContacts.ps1` | ActiveDirectory | Contact attribute updates and target address alignment. |
| M-ADUC-0040 | `M-ADUC-0040-Update-ActiveDirectorySecurityGroups.ps1` | ActiveDirectory | Group property updates (name, description, scope). |
| M-ADUC-0050 | `M-ADUC-0050-Update-ActiveDirectoryDistributionGroups.ps1` | ActiveDirectory | Distribution group property updates (name, description, scope). |
| M-ADUC-0060 | `M-ADUC-0060-Set-ActiveDirectorySecurityGroupMembers.ps1` | ActiveDirectory | Add/remove security group members. |
| M-ADUC-0070 | `M-ADUC-0070-Set-ActiveDirectoryDistributionGroupMembers.ps1` | ActiveDirectory | Add/remove distribution group members. |
| M-ADUC-0080 | `M-ADUC-0080-Set-ActiveDirectoryTemporaryPasswords.ps1` | ActiveDirectory | Standalone temporary password reset workflow (optional unlock/enable controls). |
| M-ADUC-0090 | `M-ADUC-0090-Move-ActiveDirectoryObjects.ps1` | ActiveDirectory | OU move operations for users, groups, and contacts. |
| M-GPOL-0010 | `M-GPOL-0010-Set-GroupPolicyLinks.ps1` | GroupPolicy | Create/update GPO links and enforcement/order settings. (**Planned**) |
| M-EXOP-0010 | `M-EXOP-0010-Update-ExchangeOnPremMailContacts.ps1` | ExchangeOnPrem | Mail contact property updates. |
| M-EXOP-0020 | `M-EXOP-0020-Update-ExchangeOnPremDistributionLists.ps1` | ExchangeOnPrem | Distribution list property updates. |
| M-EXOP-0030 | `M-EXOP-0030-Update-ExchangeOnPremMailEnabledSecurityGroups.ps1` | ExchangeOnPrem | Mail-enabled security group updates. |
| M-EXOP-0040 | `M-EXOP-0040-Update-ExchangeOnPremDynamicDistributionGroups.ps1` | ExchangeOnPrem | Dynamic distribution group filter/property updates. |
| M-EXOP-0050 | `M-EXOP-0050-Update-ExchangeOnPremSharedMailboxes.ps1` | ExchangeOnPrem | Shared mailbox property updates. |
| M-EXOP-0060 | `M-EXOP-0060-Update-ExchangeOnPremResourceMailboxes.ps1` | ExchangeOnPrem | Resource mailbox settings updates. |
| M-EXOP-0070 | `M-EXOP-0070-Set-ExchangeOnPremDistributionListMembers.ps1` | ExchangeOnPrem | Add/remove distribution list members. |
| M-EXOP-0080 | `M-EXOP-0080-Set-ExchangeOnPremSharedMailboxPermissions.ps1` | ExchangeOnPrem | Configure full access/send rights. |
| M-EXOP-0090 | `M-EXOP-0090-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1` | ExchangeOnPrem | Configure booking delegates and processing flags. |
| M-EXOP-0100 | `M-EXOP-0100-Set-ExchangeOnPremMailboxDelegations.ps1` | ExchangeOnPrem | Configure mailbox delegation rights. |
| M-EXOP-0110 | `M-EXOP-0110-Set-ExchangeOnPremMailboxFolderPermissions.ps1` | ExchangeOnPrem | Configure folder-level mailbox permissions. |
| M-EXOP-0120 | `M-EXOP-0120-Set-ExchangeOnPremUserMailboxForwarding.ps1` | ExchangeOnPrem | Set per-user mailbox forwarding mode and delivery behavior. |
| M-EXOP-0130 | `M-EXOP-0130-Convert-ExchangeOnPremMailboxToMailEnabledUser.ps1` | ExchangeOnPrem | Convert mailbox objects to mail-enabled users while preserving routing addresses. |
| M-EXOP-0140 | `M-EXOP-0140-Set-ExchangeOnPremMigrationWizDelegation.ps1` | ExchangeOnPrem | Grant/remove migration service delegation (FullAccess/SendAs) on target mailboxes. |
| M-FILE-0010 | `M-FILE-0010-Update-FileServicesShares.ps1` | FileServices | Share property/path updates. (**Planned**) |
| M-FILE-0020 | `M-FILE-0020-Set-FileServicesSharePermissions.ps1` | FileServices | Share ACL changes. (**Planned**) |
| M-FILE-0030 | `M-FILE-0030-Set-FileServicesNtfsPermissions.ps1` | FileServices | NTFS ACL changes. (**Planned**) |
| M-FILE-0040 | `M-FILE-0040-Update-FileServicesHomeDrives.ps1` | FileServices | Home drive path/quota updates. (**Planned**) |
| M-FILE-0050 | `M-FILE-0050-Set-FileServicesOwnerAndFullControlBySid.ps1` | FileServices | Set file/folder owner SID and grant full control to a specified SID (icacls-based). (**Planned**) |
| M-FILE-0060 | `M-FILE-0060-Grant-FileServicesFullControlBySid.ps1` | FileServices | Grant full control to a specified SID for a file/folder path (icacls-based). (**Planned**) |

Implemented now:

- `M-ADUC-0020-Update-ActiveDirectoryUsers.ps1`
- `M-ADUC-0020-Update-ActiveDirectoryUsers.input.csv`
- `M-ADUC-0030-Update-ActiveDirectoryContacts.ps1`
- `M-ADUC-0030-Update-ActiveDirectoryContacts.input.csv`
- `M-ADUC-0040-Update-ActiveDirectorySecurityGroups.ps1`
- `M-ADUC-0040-Update-ActiveDirectorySecurityGroups.input.csv`
- `M-ADUC-0050-Update-ActiveDirectoryDistributionGroups.ps1`
- `M-ADUC-0050-Update-ActiveDirectoryDistributionGroups.input.csv`
- `M-ADUC-0060-Set-ActiveDirectorySecurityGroupMembers.ps1`
- `M-ADUC-0060-Set-ActiveDirectorySecurityGroupMembers.input.csv`
- `M-ADUC-0070-Set-ActiveDirectoryDistributionGroupMembers.ps1`
- `M-ADUC-0070-Set-ActiveDirectoryDistributionGroupMembers.input.csv`
- `M-ADUC-0080-Set-ActiveDirectoryTemporaryPasswords.ps1`
- `M-ADUC-0080-Set-ActiveDirectoryTemporaryPasswords.input.csv`
- `M-ADUC-0090-Move-ActiveDirectoryObjects.ps1`
- `M-ADUC-0090-Move-ActiveDirectoryObjects.input.csv`
- `M-EXOP-0010-Update-ExchangeOnPremMailContacts.ps1`
- `M-EXOP-0010-Update-ExchangeOnPremMailContacts.input.csv`
- `M-EXOP-0020-Update-ExchangeOnPremDistributionLists.ps1`
- `M-EXOP-0020-Update-ExchangeOnPremDistributionLists.input.csv`
- `M-EXOP-0030-Update-ExchangeOnPremMailEnabledSecurityGroups.ps1`
- `M-EXOP-0030-Update-ExchangeOnPremMailEnabledSecurityGroups.input.csv`
- `M-EXOP-0040-Update-ExchangeOnPremDynamicDistributionGroups.ps1`
- `M-EXOP-0040-Update-ExchangeOnPremDynamicDistributionGroups.input.csv`
- `M-EXOP-0050-Update-ExchangeOnPremSharedMailboxes.ps1`
- `M-EXOP-0050-Update-ExchangeOnPremSharedMailboxes.input.csv`
- `M-EXOP-0060-Update-ExchangeOnPremResourceMailboxes.ps1`
- `M-EXOP-0060-Update-ExchangeOnPremResourceMailboxes.input.csv`
- `M-EXOP-0070-Set-ExchangeOnPremDistributionListMembers.ps1`
- `M-EXOP-0070-Set-ExchangeOnPremDistributionListMembers.input.csv`
- `M-EXOP-0080-Set-ExchangeOnPremSharedMailboxPermissions.ps1`
- `M-EXOP-0080-Set-ExchangeOnPremSharedMailboxPermissions.input.csv`
- `M-EXOP-0090-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1`
- `M-EXOP-0090-Set-ExchangeOnPremResourceMailboxBookingDelegates.input.csv`
- `M-EXOP-0100-Set-ExchangeOnPremMailboxDelegations.ps1`
- `M-EXOP-0100-Set-ExchangeOnPremMailboxDelegations.input.csv`
- `M-EXOP-0110-Set-ExchangeOnPremMailboxFolderPermissions.ps1`
- `M-EXOP-0110-Set-ExchangeOnPremMailboxFolderPermissions.input.csv`
- `M-EXOP-0120-Set-ExchangeOnPremUserMailboxForwarding.ps1`
- `M-EXOP-0120-Set-ExchangeOnPremUserMailboxForwarding.input.csv`
- `M-EXOP-0130-Convert-ExchangeOnPremMailboxToMailEnabledUser.ps1`
- `M-EXOP-0130-Convert-ExchangeOnPremMailboxToMailEnabledUser.input.csv`
- `M-EXOP-0140-Set-ExchangeOnPremMigrationWizDelegation.ps1`
- `M-EXOP-0140-Set-ExchangeOnPremMigrationWizDelegation.input.csv`

## References

- [OnPrem Modify Detailed Catalog](./README-Modify-Catalog.md)
- [OnPrem README](../README.md)
- [Root README](../../../README.md)
- [Operator Runbook](./RUNBOOK-Modify.md)
