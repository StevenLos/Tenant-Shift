# OnPrem Modify

`SharedModule/OnPrem/Modify` is for planned update/change scripts targeting existing on-prem objects.

Operational label: **Modify**.

Current status: ActiveDirectory modify baseline is implemented with `SM-M0001-Update-ActiveDirectoryUsers.ps1`, `SM-M0002-Update-ActiveDirectoryContacts.ps1`, `SM-M0005-Update-ActiveDirectorySecurityGroups.ps1`, `SM-M0006-Update-ActiveDirectoryDistributionGroups.ps1`, `SM-M0007-Set-ActiveDirectorySecurityGroupMembers.ps1`, `SM-M0008-Set-ActiveDirectoryDistributionGroupMembers.ps1`, `SM-M0009-Move-ActiveDirectoryObjects.ps1`, and `SM-M0010-Set-ActiveDirectoryTemporaryPasswords.ps1`. ExchangeOnPrem modify scripts are implemented with `SM-M0213-Update-ExchangeOnPremMailContacts.ps1`, `SM-M0214-Update-ExchangeOnPremDistributionLists.ps1`, `SM-M0215-Set-ExchangeOnPremDistributionListMembers.ps1`, `SM-M0216-Update-ExchangeOnPremSharedMailboxes.ps1`, `SM-M0217-Set-ExchangeOnPremSharedMailboxPermissions.ps1`, `SM-M0218-Update-ExchangeOnPremResourceMailboxes.ps1`, `SM-M0219-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1`, `SM-M0220-Set-ExchangeOnPremMailboxDelegations.ps1`, `SM-M0221-Set-ExchangeOnPremMailboxFolderPermissions.ps1`, `SM-M0222-Update-ExchangeOnPremMailEnabledSecurityGroups.ps1`, `SM-M0223-Update-ExchangeOnPremDynamicDistributionGroups.ps1`, `SM-M0224-Set-ExchangeOnPremUserMailboxForwarding.ps1`, `SM-M0225-Convert-ExchangeOnPremMailboxToMailEnabledUser.ps1`, and `SM-M0226-Set-ExchangeOnPremMigrationWizDelegation.ps1`; remaining scripts are planned.

## Purpose

Use this folder for controlled change operations after initial provisioning:

- Attribute and policy updates
- Membership and permission changes
- Lifecycle changes for existing objects

Do not use this folder for:

- Initial object creation (use `SharedModule/OnPrem/Provision`)
- Read-only reporting (use `SharedModule/OnPrem/InventoryAndReport`)

## Naming Standard

- Script: `MWWNN-<Action>-<Target>.ps1`
- Input template: `MWWNN-<Action>-<Target>.input.csv`
- Output pattern: `Results_MWWNN-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Transcript log pattern: `Transcript_MWWNN-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`

## Runtime Assumptions

- ActiveDirectory modify scripts (`00xx`) run in Windows PowerShell `5.1`.
- ExchangeOnPrem modify scripts (`02xx`) run in Exchange Management Shell (Windows PowerShell `5.1`).

Workload code allocation (`WW` in `<Prefix><WW><NN>`):

- `00xx`: ActiveDirectory
- `01xx`: GroupPolicy
- `02xx`: ExchangeOnPrem
- `03xx`: FileServices

## Script Matrix (Current Status)

| ID | Script | Workload | Primary Change Scope |
|---|---|---|---|
| M0001 | `SM-M0001-Update-ActiveDirectoryUsers.ps1` | ActiveDirectory | User attribute updates (display, UPN, manager, office, phone), including fold-in password reset support. |
| M0002 | `SM-M0002-Update-ActiveDirectoryContacts.ps1` | ActiveDirectory | Contact attribute updates and target address alignment. |
| M0005 | `SM-M0005-Update-ActiveDirectorySecurityGroups.ps1` | ActiveDirectory | Group property updates (name, description, scope). |
| M0006 | `SM-M0006-Update-ActiveDirectoryDistributionGroups.ps1` | ActiveDirectory | Distribution group property updates (name, description, scope). |
| M0007 | `SM-M0007-Set-ActiveDirectorySecurityGroupMembers.ps1` | ActiveDirectory | Add/remove security group members. |
| M0008 | `SM-M0008-Set-ActiveDirectoryDistributionGroupMembers.ps1` | ActiveDirectory | Add/remove distribution group members. |
| M0009 | `SM-M0009-Move-ActiveDirectoryObjects.ps1` | ActiveDirectory | OU move operations for users, groups, and contacts. |
| M0010 | `SM-M0010-Set-ActiveDirectoryTemporaryPasswords.ps1` | ActiveDirectory | Standalone temporary password reset workflow (optional unlock/enable controls). |
| M0101 | `SM-M0101-Set-GroupPolicyLinks.ps1` | GroupPolicy | Create/update GPO links and enforcement/order settings. |
| M0213 | `SM-M0213-Update-ExchangeOnPremMailContacts.ps1` | ExchangeOnPrem | Mail contact property updates. |
| M0214 | `SM-M0214-Update-ExchangeOnPremDistributionLists.ps1` | ExchangeOnPrem | Distribution list property updates. |
| M0215 | `SM-M0215-Set-ExchangeOnPremDistributionListMembers.ps1` | ExchangeOnPrem | Add/remove distribution list members. |
| M0216 | `SM-M0216-Update-ExchangeOnPremSharedMailboxes.ps1` | ExchangeOnPrem | Shared mailbox property updates. |
| M0217 | `SM-M0217-Set-ExchangeOnPremSharedMailboxPermissions.ps1` | ExchangeOnPrem | Configure full access/send rights. |
| M0218 | `SM-M0218-Update-ExchangeOnPremResourceMailboxes.ps1` | ExchangeOnPrem | Resource mailbox settings updates. |
| M0219 | `SM-M0219-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1` | ExchangeOnPrem | Configure booking delegates and processing flags. |
| M0220 | `SM-M0220-Set-ExchangeOnPremMailboxDelegations.ps1` | ExchangeOnPrem | Configure mailbox delegation rights. |
| M0221 | `SM-M0221-Set-ExchangeOnPremMailboxFolderPermissions.ps1` | ExchangeOnPrem | Configure folder-level mailbox permissions. |
| M0222 | `SM-M0222-Update-ExchangeOnPremMailEnabledSecurityGroups.ps1` | ExchangeOnPrem | Mail-enabled security group updates. |
| M0223 | `SM-M0223-Update-ExchangeOnPremDynamicDistributionGroups.ps1` | ExchangeOnPrem | Dynamic distribution group filter/property updates. |
| M0224 | `SM-M0224-Set-ExchangeOnPremUserMailboxForwarding.ps1` | ExchangeOnPrem | Set per-user mailbox forwarding mode and delivery behavior. |
| M0225 | `SM-M0225-Convert-ExchangeOnPremMailboxToMailEnabledUser.ps1` | ExchangeOnPrem | Convert mailbox objects to mail-enabled users while preserving routing addresses. |
| M0226 | `SM-M0226-Set-ExchangeOnPremMigrationWizDelegation.ps1` | ExchangeOnPrem | Grant/remove migration service delegation (FullAccess/SendAs) on target mailboxes. |
| M0301 | `SM-M0301-Update-FileServicesShares.ps1` | FileServices | Share property/path updates. |
| M0302 | `SM-M0302-Set-FileServicesSharePermissions.ps1` | FileServices | Share ACL changes. |
| M0303 | `SM-M0303-Set-FileServicesNtfsPermissions.ps1` | FileServices | NTFS ACL changes. |
| M0304 | `SM-M0304-Update-FileServicesHomeDrives.ps1` | FileServices | Home drive path/quota updates. |
| M0305 | `SM-M0305-Set-FileServicesOwnerAndFullControlBySid.ps1` | FileServices | Set file/folder owner SID and grant full control to a specified SID (icacls-based). |
| M0306 | `SM-M0306-Grant-FileServicesFullControlBySid.ps1` | FileServices | Grant full control to a specified SID for a file/folder path (icacls-based). |

Implemented now:

- `SM-M0001-Update-ActiveDirectoryUsers.ps1`
- `SM-M0001-Update-ActiveDirectoryUsers.input.csv`
- `SM-M0002-Update-ActiveDirectoryContacts.ps1`
- `SM-M0002-Update-ActiveDirectoryContacts.input.csv`
- `SM-M0005-Update-ActiveDirectorySecurityGroups.ps1`
- `SM-M0005-Update-ActiveDirectorySecurityGroups.input.csv`
- `SM-M0006-Update-ActiveDirectoryDistributionGroups.ps1`
- `SM-M0006-Update-ActiveDirectoryDistributionGroups.input.csv`
- `SM-M0007-Set-ActiveDirectorySecurityGroupMembers.ps1`
- `SM-M0007-Set-ActiveDirectorySecurityGroupMembers.input.csv`
- `SM-M0008-Set-ActiveDirectoryDistributionGroupMembers.ps1`
- `SM-M0008-Set-ActiveDirectoryDistributionGroupMembers.input.csv`
- `SM-M0009-Move-ActiveDirectoryObjects.ps1`
- `SM-M0009-Move-ActiveDirectoryObjects.input.csv`
- `SM-M0010-Set-ActiveDirectoryTemporaryPasswords.ps1`
- `SM-M0010-Set-ActiveDirectoryTemporaryPasswords.input.csv`
- `SM-M0213-Update-ExchangeOnPremMailContacts.ps1`
- `SM-M0213-Update-ExchangeOnPremMailContacts.input.csv`
- `SM-M0214-Update-ExchangeOnPremDistributionLists.ps1`
- `SM-M0214-Update-ExchangeOnPremDistributionLists.input.csv`
- `SM-M0215-Set-ExchangeOnPremDistributionListMembers.ps1`
- `SM-M0215-Set-ExchangeOnPremDistributionListMembers.input.csv`
- `SM-M0216-Update-ExchangeOnPremSharedMailboxes.ps1`
- `SM-M0216-Update-ExchangeOnPremSharedMailboxes.input.csv`
- `SM-M0217-Set-ExchangeOnPremSharedMailboxPermissions.ps1`
- `SM-M0217-Set-ExchangeOnPremSharedMailboxPermissions.input.csv`
- `SM-M0218-Update-ExchangeOnPremResourceMailboxes.ps1`
- `SM-M0218-Update-ExchangeOnPremResourceMailboxes.input.csv`
- `SM-M0219-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1`
- `SM-M0219-Set-ExchangeOnPremResourceMailboxBookingDelegates.input.csv`
- `SM-M0220-Set-ExchangeOnPremMailboxDelegations.ps1`
- `SM-M0220-Set-ExchangeOnPremMailboxDelegations.input.csv`
- `SM-M0221-Set-ExchangeOnPremMailboxFolderPermissions.ps1`
- `SM-M0221-Set-ExchangeOnPremMailboxFolderPermissions.input.csv`
- `SM-M0222-Update-ExchangeOnPremMailEnabledSecurityGroups.ps1`
- `SM-M0222-Update-ExchangeOnPremMailEnabledSecurityGroups.input.csv`
- `SM-M0223-Update-ExchangeOnPremDynamicDistributionGroups.ps1`
- `SM-M0223-Update-ExchangeOnPremDynamicDistributionGroups.input.csv`
- `SM-M0224-Set-ExchangeOnPremUserMailboxForwarding.ps1`
- `SM-M0224-Set-ExchangeOnPremUserMailboxForwarding.input.csv`
- `SM-M0225-Convert-ExchangeOnPremMailboxToMailEnabledUser.ps1`
- `SM-M0225-Convert-ExchangeOnPremMailboxToMailEnabledUser.input.csv`
- `SM-M0226-Set-ExchangeOnPremMigrationWizDelegation.ps1`
- `SM-M0226-Set-ExchangeOnPremMigrationWizDelegation.input.csv`

## References

- [OnPrem Modify Detailed Catalog](./README-Modify-Catalog.md)
- [OnPrem README](../README.md)
- [SharedModule README](../../README.md)
