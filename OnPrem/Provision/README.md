# OnPrem Provision

`OnPrem/Provision` is for planned create and initial setup scripts targeting on-prem workloads.

Operational label: **Provision**.

Current status: planning matrix is defined; no OnPrem provision scripts are implemented yet.

## Purpose

Use this folder for first-time object creation in on-prem environments:

- ActiveDirectory object creation baselines
- ExchangeOnPrem recipient/group creation baselines
- FileServices share/folder creation baselines

Do not use this folder for:

- Ongoing updates to existing objects (use `OnPrem/Modify`)
- Read-only discovery exports (use `OnPrem/InventoryAndReport`)

## Naming Standard

- Script: `PWWNN-<Action>-<Target>.ps1`
- Input template: `PWWNN-<Action>-<Target>.input.csv`
- Output pattern: `Results_PWWNN-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Transcript log pattern: `Transcript_PWWNN-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`

Workload code allocation (`WW` in `<Prefix><WW><NN>`):

- `00xx`: ActiveDirectory
- `01xx`: GroupPolicy
- `02xx`: ExchangeOnPrem
- `03xx`: FileServices

## Planned Script Matrix

| ID | Script | Workload | Purpose |
|---|---|---|---|
| P0001 | `P0001-Create-ActiveDirectoryUsers.ps1` | ActiveDirectory | Create AD user accounts from CSV. |
| P0002 | `P0002-Create-ActiveDirectoryContacts.ps1` | ActiveDirectory | Create AD contact objects from CSV. |
| P0005 | `P0005-Create-ActiveDirectorySecurityGroups.ps1` | ActiveDirectory | Create AD security groups. |
| P0009 | `P0009-Create-ActiveDirectoryOrganizationalUnits.ps1` | ActiveDirectory | Create OUs for identity placement and delegation boundaries. |
| P0101 | `P0101-Import-GroupPolicyBackups.ps1` | GroupPolicy | Import and create GPOs from backup paths. |
| P0213 | `P0213-Create-ExchangeOnPremMailContacts.ps1` | ExchangeOnPrem | Create on-prem mail contacts aligned to AD objects. |
| P0214 | `P0214-Create-ExchangeOnPremDistributionLists.ps1` | ExchangeOnPrem | Create on-prem distribution groups. |
| P0215 | `P0215-Create-ExchangeOnPremMailEnabledSecurityGroups.ps1` | ExchangeOnPrem | Create on-prem mail-enabled security groups. |
| P0216 | `P0216-Create-ExchangeOnPremSharedMailboxes.ps1` | ExchangeOnPrem | Create on-prem shared mailboxes. |
| P0218 | `P0218-Create-ExchangeOnPremResourceMailboxes.ps1` | ExchangeOnPrem | Create on-prem room and equipment mailboxes. |
| P0219 | `P0219-Create-ExchangeOnPremDynamicDistributionGroups.ps1` | ExchangeOnPrem | Create on-prem dynamic distribution groups. |
| P0301 | `P0301-Create-FileServicesShares.ps1` | FileServices | Create file shares from CSV definitions. |
| P0302 | `P0302-Set-FileServicesSharePermissions.ps1` | FileServices | Apply share-level ACL baselines. |
| P0303 | `P0303-Set-FileServicesNtfsPermissions.ps1` | FileServices | Apply NTFS ACL baselines. |
| P0304 | `P0304-Create-FileServicesHomeDrives.ps1` | FileServices | Create user home drive folders and shares. |

## Planned Execution Order

1. ActiveDirectory baseline: `P0001`, `P0002`, `P0005`, `P0009`
2. GroupPolicy baseline: `P0101`
3. ExchangeOnPrem baseline: `P0213`, `P0214`, `P0215`, `P0216`, `P0218`, `P0219`
4. FileServices baseline: `P0301`, `P0302`, `P0303`, `P0304`

## References

- [OnPrem Provision Detailed Catalog](./README-Provision-Catalog.md)
- [OnPrem README](../README.md)
- [Root README](../../README.md)



