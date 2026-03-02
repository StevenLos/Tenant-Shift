# OnPrem Modify

`OnPrem/Modify` is for planned update/change scripts targeting existing on-prem objects.

Operational label: **Modify**.

Current status: planning matrix is defined; no OnPrem modify scripts are implemented yet.

## Purpose

Use this folder for controlled change operations after initial provisioning:

- Attribute and policy updates
- Membership and permission changes
- Lifecycle changes for existing objects

Do not use this folder for:

- Initial object creation (use `OnPrem/Provision`)
- Read-only reporting (use `OnPrem/InventoryAndReport`)

## Naming Standard

- Script: `MWWNN-<Action>-<Target>.ps1`
- Input template: `MWWNN-<Action>-<Target>.input.csv`
- Output pattern: `Results_MWWNN-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Transcript log pattern: `Transcript_MWWNN-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`

Workload code allocation (`WW` in `<Prefix><WW><NN>`):

- `00xx`: ActiveDirectory
- `01xx`: GroupPolicy
- `02xx`: ExchangeOnPrem
- `03xx`: FileServices

## Planned Script Matrix

| ID | Script | Workload | Primary Change Scope |
|---|---|---|---|
| M0001 | `M0001-Update-ActiveDirectoryUsers.ps1` | ActiveDirectory | User attribute updates (display, UPN, manager, office, phone). |
| M0002 | `M0002-Update-ActiveDirectoryContacts.ps1` | ActiveDirectory | Contact attribute updates and target address alignment. |
| M0005 | `M0005-Update-ActiveDirectorySecurityGroups.ps1` | ActiveDirectory | Group property updates (name, description, scope). |
| M0007 | `M0007-Set-ActiveDirectorySecurityGroupMembers.ps1` | ActiveDirectory | Add/remove security group members. |
| M0009 | `M0009-Move-ActiveDirectoryObjects.ps1` | ActiveDirectory | OU move operations for users, groups, and contacts. |
| M0101 | `M0101-Set-GroupPolicyLinks.ps1` | GroupPolicy | Create/update GPO links and enforcement/order settings. |
| M0213 | `M0213-Update-ExchangeOnPremMailContacts.ps1` | ExchangeOnPrem | Mail contact property updates. |
| M0214 | `M0214-Update-ExchangeOnPremDistributionLists.ps1` | ExchangeOnPrem | Distribution list property updates. |
| M0215 | `M0215-Set-ExchangeOnPremDistributionListMembers.ps1` | ExchangeOnPrem | Add/remove distribution list members. |
| M0216 | `M0216-Update-ExchangeOnPremSharedMailboxes.ps1` | ExchangeOnPrem | Shared mailbox property updates. |
| M0217 | `M0217-Set-ExchangeOnPremSharedMailboxPermissions.ps1` | ExchangeOnPrem | Configure full access/send rights. |
| M0218 | `M0218-Update-ExchangeOnPremResourceMailboxes.ps1` | ExchangeOnPrem | Resource mailbox settings updates. |
| M0219 | `M0219-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1` | ExchangeOnPrem | Configure booking delegates and processing flags. |
| M0220 | `M0220-Set-ExchangeOnPremMailboxDelegations.ps1` | ExchangeOnPrem | Configure mailbox delegation rights. |
| M0221 | `M0221-Set-ExchangeOnPremMailboxFolderPermissions.ps1` | ExchangeOnPrem | Configure folder-level mailbox permissions. |
| M0222 | `M0222-Update-ExchangeOnPremMailEnabledSecurityGroups.ps1` | ExchangeOnPrem | Mail-enabled security group updates. |
| M0223 | `M0223-Update-ExchangeOnPremDynamicDistributionGroups.ps1` | ExchangeOnPrem | Dynamic distribution group filter/property updates. |
| M0301 | `M0301-Update-FileServicesShares.ps1` | FileServices | Share property/path updates. |
| M0302 | `M0302-Set-FileServicesSharePermissions.ps1` | FileServices | Share ACL changes. |
| M0303 | `M0303-Set-FileServicesNtfsPermissions.ps1` | FileServices | NTFS ACL changes. |
| M0304 | `M0304-Update-FileServicesHomeDrives.ps1` | FileServices | Home drive path/quota updates. |
| M0305 | `M0305-Set-FileServicesOwnerAndFullControlBySid.ps1` | FileServices | Set file/folder owner SID and grant full control to a specified SID (icacls-based). |
| M0306 | `M0306-Grant-FileServicesFullControlBySid.ps1` | FileServices | Grant full control to a specified SID for a file/folder path (icacls-based). |

## References

- [OnPrem Modify Detailed Catalog](./README-Modify-Catalog.md)
- [OnPrem README](../README.md)
- [Root README](../../README.md)



