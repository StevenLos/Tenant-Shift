# OnPrem Provision

`SharedModule/OnPrem/Provision` is for planned create and initial setup scripts targeting on-prem workloads.

Operational label: **Provision**.

Current status: ActiveDirectory provision baseline is implemented with `P-ADUC-0020-Create-ActiveDirectoryUsers.ps1`, `P-ADUC-0030-Create-ActiveDirectoryContacts.ps1`, `P-ADUC-0040-Create-ActiveDirectorySecurityGroups.ps1`, `P-ADUC-0050-Create-ActiveDirectoryDistributionGroups.ps1`, and `P-ADUC-0010-Create-ActiveDirectoryOrganizationalUnits.ps1`. ExchangeOnPrem provision scripts are implemented with `P-EXOP-0010-Create-ExchangeOnPremMailContacts.ps1`, `P-EXOP-0020-Create-ExchangeOnPremDistributionLists.ps1`, `P-EXOP-0030-Create-ExchangeOnPremMailEnabledSecurityGroups.ps1`, `P-EXOP-0050-Create-ExchangeOnPremSharedMailboxes.ps1`, `P-EXOP-0060-Create-ExchangeOnPremResourceMailboxes.ps1`, and `P-EXOP-0040-Create-ExchangeOnPremDynamicDistributionGroups.ps1`; remaining non-ActiveDirectory/non-ExchangeOnPrem provision scripts remain planned.

## Purpose

Use this folder for first-time object creation in on-prem environments:

- ActiveDirectory object creation baselines
- ExchangeOnPrem recipient/group creation baselines
- FileServices share/folder creation baselines

Do not use this folder for:

- Ongoing updates to existing objects (use `SharedModule/OnPrem/Modify`)
- Read-only discovery exports (use `SharedModule/OnPrem/InventoryAndReport`)

## Naming Standard

- Script: `P-<WW>-<NNNN>-<Action>-<Target>.ps1`
- Input template: `P-<WW>-<NNNN>-<Action>-<Target>.input.csv`
- Output pattern: `Results_P-<WW>-<NNNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Transcript log pattern: `Transcript_P-<WW>-<NNNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`

## Runtime Assumptions

- ActiveDirectory provision scripts (ADUC) run in Windows PowerShell `5.1`.
- ExchangeOnPrem provision scripts (EXOP) run in Exchange Management Shell (Windows PowerShell `5.1`).

Workload code allocation (`WW` in `P-<WW>-<NNNN>`):

- `ADUC`: ActiveDirectory
- `GPOL`: GroupPolicy
- `EXOP`: ExchangeOnPrem
- `FILE`: FileServices

## Script Matrix (Current Status)

| ID | Script | Workload | Purpose |
|---|---|---|---|
| P-ADUC-0010 | `P-ADUC-0010-Create-ActiveDirectoryOrganizationalUnits.ps1` | ActiveDirectory | Create OUs for identity placement and delegation boundaries. |
| P-ADUC-0020 | `P-ADUC-0020-Create-ActiveDirectoryUsers.ps1` | ActiveDirectory | Create AD user accounts from CSV. |
| P-ADUC-0030 | `P-ADUC-0030-Create-ActiveDirectoryContacts.ps1` | ActiveDirectory | Create AD contact objects from CSV. |
| P-ADUC-0040 | `P-ADUC-0040-Create-ActiveDirectorySecurityGroups.ps1` | ActiveDirectory | Create AD security groups. |
| P-ADUC-0050 | `P-ADUC-0050-Create-ActiveDirectoryDistributionGroups.ps1` | ActiveDirectory | Create AD distribution groups. |
| P-GPOL-0010 | `P-GPOL-0010-Import-GroupPolicyBackups.ps1` | GroupPolicy | Import and create GPOs from backup paths. (**Planned**) |
| P-EXOP-0010 | `P-EXOP-0010-Create-ExchangeOnPremMailContacts.ps1` | ExchangeOnPrem | Create on-prem mail contacts aligned to AD objects. |
| P-EXOP-0020 | `P-EXOP-0020-Create-ExchangeOnPremDistributionLists.ps1` | ExchangeOnPrem | Create on-prem distribution groups. |
| P-EXOP-0030 | `P-EXOP-0030-Create-ExchangeOnPremMailEnabledSecurityGroups.ps1` | ExchangeOnPrem | Create on-prem mail-enabled security groups. |
| P-EXOP-0040 | `P-EXOP-0040-Create-ExchangeOnPremDynamicDistributionGroups.ps1` | ExchangeOnPrem | Create on-prem dynamic distribution groups. |
| P-EXOP-0050 | `P-EXOP-0050-Create-ExchangeOnPremSharedMailboxes.ps1` | ExchangeOnPrem | Create on-prem shared mailboxes. |
| P-EXOP-0060 | `P-EXOP-0060-Create-ExchangeOnPremResourceMailboxes.ps1` | ExchangeOnPrem | Create on-prem room and equipment mailboxes. |
| P-FILE-0010 | `P-FILE-0010-Create-FileServicesShares.ps1` | FileServices | Create file shares from CSV definitions. (**Planned**) |
| P-FILE-0020 | `P-FILE-0020-Set-FileServicesSharePermissions.ps1` | FileServices | Apply share-level ACL baselines. (**Planned**) |
| P-FILE-0030 | `P-FILE-0030-Set-FileServicesNtfsPermissions.ps1` | FileServices | Apply NTFS ACL baselines. (**Planned**) |
| P-FILE-0040 | `P-FILE-0040-Create-FileServicesHomeDrives.ps1` | FileServices | Create user home drive folders and shares. (**Planned**) |

Implemented now:

- `P-ADUC-0010-Create-ActiveDirectoryOrganizationalUnits.ps1`
- `P-ADUC-0010-Create-ActiveDirectoryOrganizationalUnits.input.csv`
- `P-ADUC-0020-Create-ActiveDirectoryUsers.ps1`
- `P-ADUC-0020-Create-ActiveDirectoryUsers.input.csv`
- `P-ADUC-0030-Create-ActiveDirectoryContacts.ps1`
- `P-ADUC-0030-Create-ActiveDirectoryContacts.input.csv`
- `P-ADUC-0040-Create-ActiveDirectorySecurityGroups.ps1`
- `P-ADUC-0040-Create-ActiveDirectorySecurityGroups.input.csv`
- `P-ADUC-0050-Create-ActiveDirectoryDistributionGroups.ps1`
- `P-ADUC-0050-Create-ActiveDirectoryDistributionGroups.input.csv`
- `P-EXOP-0010-Create-ExchangeOnPremMailContacts.ps1`
- `P-EXOP-0010-Create-ExchangeOnPremMailContacts.input.csv`
- `P-EXOP-0020-Create-ExchangeOnPremDistributionLists.ps1`
- `P-EXOP-0020-Create-ExchangeOnPremDistributionLists.input.csv`
- `P-EXOP-0030-Create-ExchangeOnPremMailEnabledSecurityGroups.ps1`
- `P-EXOP-0030-Create-ExchangeOnPremMailEnabledSecurityGroups.input.csv`
- `P-EXOP-0040-Create-ExchangeOnPremDynamicDistributionGroups.ps1`
- `P-EXOP-0040-Create-ExchangeOnPremDynamicDistributionGroups.input.csv`
- `P-EXOP-0050-Create-ExchangeOnPremSharedMailboxes.ps1`
- `P-EXOP-0050-Create-ExchangeOnPremSharedMailboxes.input.csv`
- `P-EXOP-0060-Create-ExchangeOnPremResourceMailboxes.ps1`
- `P-EXOP-0060-Create-ExchangeOnPremResourceMailboxes.input.csv`

## Planned Execution Order

1. ActiveDirectory baseline: `P-ADUC-0010`, `P-ADUC-0020`, `P-ADUC-0030`, `P-ADUC-0040`, `P-ADUC-0050`
2. GroupPolicy baseline: `P-GPOL-0010`
3. ExchangeOnPrem baseline: `P-EXOP-0010`, `P-EXOP-0020`, `P-EXOP-0030`, `P-EXOP-0040`, `P-EXOP-0050`, `P-EXOP-0060`
4. FileServices baseline: `P-FILE-0010`, `P-FILE-0020`, `P-FILE-0030`, `P-FILE-0040`

## References

- [OnPrem Provision Detailed Catalog](./README-Provision-Catalog.md)
- [OnPrem README](../README.md)
- [Root README](../../../README.md)
- [Operator Runbook](./RUNBOOK-Provision.md)
