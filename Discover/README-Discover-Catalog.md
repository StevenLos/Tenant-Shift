# Discover Detailed Catalog

Detailed catalog for discovery/reporting scripts in `Discover/`.

Current implementation status: planned catalog, scripts not yet created.

## Script Contract

All discover scripts should:

- Be read-only (no create/update/delete actions)
- Run on PowerShell 7+
- Export deterministic CSV output for diff/baselining
- Include `Status` and `Message` per processed item
- Use shared helpers from `M365.Common.psm1` where practical

## Planned Catalog

| ID | Planned Script | Optional Input Template | Workload | Primary Output Focus | Status |
|---|---|---|---|---|---|
| D01 | `D01-Get-EntraUsers.ps1` | `D01-Get-EntraUsers.input.csv` | Entra | User inventory and core profile fields | Planned |
| D02 | `D02-Get-EntraGuestUsers.ps1` | `D02-Get-EntraGuestUsers.input.csv` | Entra | Guest user inventory and invite state | Planned |
| D03 | `D03-Get-EntraUserLicenses.ps1` | `D03-Get-EntraUserLicenses.input.csv` | Entra | User licensing assignments/plans | Planned |
| D04 | `D04-Get-OneDriveProvisioningStatus.ps1` | `D04-Get-OneDriveProvisioningStatus.input.csv` | OneDrive/SharePoint | OneDrive site existence/provisioning state | Planned |
| D05 | `D05-Get-EntraSecurityGroups.ps1` | `D05-Get-EntraSecurityGroups.input.csv` | Entra | Assigned security groups | Planned |
| D06 | `D06-Get-EntraDynamicUserSecurityGroups.ps1` | `D06-Get-EntraDynamicUserSecurityGroups.input.csv` | Entra | Dynamic groups and membership rules | Planned |
| D07 | `D07-Get-EntraSecurityGroupMembers.ps1` | `D07-Get-EntraSecurityGroupMembers.input.csv` | Entra | Group membership exports | Planned |
| D08 | `D08-Get-EntraMicrosoft365Groups.ps1` | `D08-Get-EntraMicrosoft365Groups.input.csv` | Entra | Microsoft 365 group config and ownership | Planned |
| D09 | `D09-Get-MicrosoftTeams.ps1` | `D09-Get-MicrosoftTeams.input.csv` | Teams | Teams inventory and core settings | Planned |
| D10 | `D10-Get-MicrosoftTeamMembers.ps1` | `D10-Get-MicrosoftTeamMembers.input.csv` | Teams | Team owner/member assignments | Planned |
| D11 | `D11-Get-MicrosoftTeamChannels.ps1` | `D11-Get-MicrosoftTeamChannels.input.csv` | Teams | Channel inventory by Team | Planned |
| D12 | `D12-Get-MicrosoftTeamChannelMembers.ps1` | `D12-Get-MicrosoftTeamChannelMembers.input.csv` | Teams | Private/shared channel membership | Planned |
| D13 | `D13-Get-ExchangeMailContacts.ps1` | `D13-Get-ExchangeMailContacts.input.csv` | Exchange | Mail contact inventory | Planned |
| D14 | `D14-Get-ExchangeDistributionLists.ps1` | `D14-Get-ExchangeDistributionLists.input.csv` | Exchange | Distribution list inventory | Planned |
| D15 | `D15-Get-ExchangeDistributionListMembers.ps1` | `D15-Get-ExchangeDistributionListMembers.input.csv` | Exchange | Distribution list membership | Planned |
| D16 | `D16-Get-ExchangeSharedMailboxes.ps1` | `D16-Get-ExchangeSharedMailboxes.input.csv` | Exchange | Shared mailbox inventory | Planned |
| D17 | `D17-Get-ExchangeSharedMailboxPermissions.ps1` | `D17-Get-ExchangeSharedMailboxPermissions.input.csv` | Exchange | Shared mailbox permission matrix | Planned |
| D18 | `D18-Get-ExchangeResourceMailboxes.ps1` | `D18-Get-ExchangeResourceMailboxes.input.csv` | Exchange | Room/equipment mailbox inventory | Planned |
| D19 | `D19-Get-ExchangeResourceMailboxBookingDelegates.ps1` | `D19-Get-ExchangeResourceMailboxBookingDelegates.input.csv` | Exchange | Resource booking delegate/policy state | Planned |
| D20 | `D20-Get-ExchangeMailboxDelegations.ps1` | `D20-Get-ExchangeMailboxDelegations.input.csv` | Exchange | Mailbox delegation matrix | Planned |
| D21 | `D21-Get-ExchangeMailboxFolderPermissions.ps1` | `D21-Get-ExchangeMailboxFolderPermissions.input.csv` | Exchange | Folder permission and delegate flags | Planned |

## Standard Output Columns

Recommended baseline columns:

- `RowNumber`
- `PrimaryKey`
- `Action`
- `Status`
- `Message`
- Workload/object-specific fields

## Suggested Execution Pattern

1. Entra baseline: `D01` to `D08`
2. Teams baseline: `D09` to `D12`
3. Exchange baseline: `D13` to `D21`

## Related Docs

- [Discover README](./README.md)
- [Root README](../README.md)
- [Build Detailed Catalog](../Build/README-Build-Catalog.md)

