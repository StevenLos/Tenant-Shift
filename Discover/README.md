# Discover Folder

`Discover` is for read-only inventory and reporting scripts.

Current status: scaffolded, no `D` scripts added yet.

## Purpose

Use this folder for:
- State baselines before/after build or update runs
- Compliance and audit exports
- Environment inventory snapshots by workload

Do not use this folder for:
- Creation, modification, or deletion operations

## Naming Standard

- Script: `D##-<Action>-<Target>.ps1`
- Optional input template: `D##-<Action>-<Target>.input.csv`
- Output pattern: `Results_D##-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`

Example:
- `D01-Get-EntraUsers.ps1`
- `D01-Get-EntraUsers.input.csv` (only when scoped input is needed)

## Run Pattern

Run from repository root:

```powershell
pwsh ./Discover/D01-Get-EntraUsers.ps1
pwsh ./Discover/D02-Get-EntraGuestUsers.ps1 -InputCsvPath ./Discover/D02-Get-EntraGuestUsers.input.csv
```

## Discover Output Standard

Discovery scripts should export consistent, easy-to-diff output:

- Primary object key columns (for example: `UserPrincipalName`, `GroupId`)
- Workload/object metadata columns
- `Status` and `Message` columns for per-record operation logging
- Timestamped output file names

## Proposed Initial Discover Backlog

| ID | Proposed Script | Workload | Purpose |
|---|---|---|---|
| D01 | `D01-Get-EntraUsers.ps1` | Entra | Export tenant users. |
| D02 | `D02-Get-EntraGuestUsers.ps1` | Entra | Export guest users. |
| D03 | `D03-Get-EntraUserLicenses.ps1` | Entra | Export assigned licenses. |
| D04 | `D04-Get-OneDriveProvisioningStatus.ps1` | OneDrive/SharePoint | Report OneDrive site provisioning status. |
| D05 | `D05-Get-EntraSecurityGroups.ps1` | Entra | Export assigned security groups. |
| D06 | `D06-Get-EntraDynamicUserSecurityGroups.ps1` | Entra | Export dynamic user groups and rules. |
| D07 | `D07-Get-EntraSecurityGroupMembers.ps1` | Entra | Export security group membership. |
| D08 | `D08-Get-EntraMicrosoft365Groups.ps1` | Entra | Export Microsoft 365 groups. |
| D09 | `D09-Get-MicrosoftTeams.ps1` | Teams | Export Teams and core settings. |
| D10 | `D10-Get-MicrosoftTeamMembers.ps1` | Teams | Export Team membership. |
| D11 | `D11-Get-MicrosoftTeamChannels.ps1` | Teams | Export channels by Team. |
| D12 | `D12-Get-MicrosoftTeamChannelMembers.ps1` | Teams | Export private/shared channel membership. |
| D13 | `D13-Get-ExchangeMailContacts.ps1` | Exchange | Export mail contacts. |
| D14 | `D14-Get-ExchangeDistributionLists.ps1` | Exchange | Export distribution lists. |
| D15 | `D15-Get-ExchangeDistributionListMembers.ps1` | Exchange | Export DL membership. |
| D16 | `D16-Get-ExchangeSharedMailboxes.ps1` | Exchange | Export shared mailboxes. |
| D17 | `D17-Get-ExchangeSharedMailboxPermissions.ps1` | Exchange | Export mailbox permissions. |
| D18 | `D18-Get-ExchangeResourceMailboxes.ps1` | Exchange | Export room/equipment mailboxes. |
| D19 | `D19-Get-ExchangeResourceMailboxBookingDelegates.ps1` | Exchange | Export booking delegate settings. |
| D20 | `D20-Get-ExchangeMailboxDelegations.ps1` | Exchange | Export mailbox delegations. |
| D21 | `D21-Get-ExchangeMailboxFolderPermissions.ps1` | Exchange | Export folder-level permissions. |

## Discover Standards

- Keep scripts read-only.
- Keep workload explicit in script names.
- Reuse `M365.Common.psm1` where common validation and result formatting helps.
- Prefer deterministic column ordering for easier diffing between snapshots.

## References

- [Root README](../README.md)
- [Build README](../Build/README.md)
- [Discover Detailed Catalog](./README-Discover-Catalog.md)
