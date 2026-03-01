# Update Detailed Catalog

Detailed catalog for update/change scripts in `Update/`.

Current implementation status: partial implementation (`U41-Set-SharePointSiteAdmins.ps1` is implemented; remaining entries are planned).

## Script Contract

All update scripts should:

- Run on PowerShell 7+
- Support `-WhatIf` and `ShouldProcess`
- Be idempotent when practical
- Validate CSV headers and required values
- Export per-record `Status` and `Message`
- Include rollback/remediation notes for high-impact operations

## Planned Catalog

| ID | Planned Script | Input Template | Workload | Primary Change Scope | Status |
|---|---|---|---|---|---|
| U01 | `U01-Update-EntraUsers.ps1` | `U01-Update-EntraUsers.input.csv` | Entra | User profile/attribute updates | Planned |
| U02 | `U02-Set-EntraUserAccountState.ps1` | `U02-Set-EntraUserAccountState.input.csv` | Entra | Enable/disable accounts | Planned |
| U03 | `U03-Set-EntraUserLicenses.ps1` | `U03-Set-EntraUserLicenses.input.csv` | Entra | Add/remove license assignments | Planned |
| U04 | `U04-Set-OneDriveSiteSettings.ps1` | `U04-Set-OneDriveSiteSettings.input.csv` | OneDrive/SharePoint | Site settings/quota updates | Planned |
| U05 | `U05-Update-EntraAssignedSecurityGroups.ps1` | `U05-Update-EntraAssignedSecurityGroups.input.csv` | Entra | Assigned security group property updates | Planned |
| U06 | `U06-Update-EntraDynamicUserSecurityGroups.ps1` | `U06-Update-EntraDynamicUserSecurityGroups.input.csv` | Entra | Dynamic rule/processing updates | Planned |
| U07 | `U07-Set-EntraSecurityGroupMembers.ps1` | `U07-Set-EntraSecurityGroupMembers.input.csv` | Entra | Add/remove group members | Planned |
| U08 | `U08-Update-EntraMicrosoft365Groups.ps1` | `U08-Update-EntraMicrosoft365Groups.input.csv` | Entra | M365 group settings/visibility updates | Planned |
| U09 | `U09-Update-MicrosoftTeams.ps1` | `U09-Update-MicrosoftTeams.input.csv` | Teams | Team settings updates | Planned |
| U10 | `U10-Set-MicrosoftTeamMembers.ps1` | `U10-Set-MicrosoftTeamMembers.input.csv` | Teams | Add/remove team owners/members | Planned |
| U11 | `U11-Update-MicrosoftTeamChannels.ps1` | `U11-Update-MicrosoftTeamChannels.input.csv` | Teams | Channel settings updates | Planned |
| U12 | `U12-Set-MicrosoftTeamChannelMembers.ps1` | `U12-Set-MicrosoftTeamChannelMembers.input.csv` | Teams | Add/remove private/shared channel members | Planned |
| U13 | `U13-Update-ExchangeMailContacts.ps1` | `U13-Update-ExchangeMailContacts.input.csv` | Exchange | Mail contact property updates | Planned |
| U14 | `U14-Update-ExchangeDistributionLists.ps1` | `U14-Update-ExchangeDistributionLists.input.csv` | Exchange | Distribution list property updates | Planned |
| U15 | `U15-Set-ExchangeDistributionListMembers.ps1` | `U15-Set-ExchangeDistributionListMembers.input.csv` | Exchange | Add/remove DL members | Planned |
| U16 | `U16-Update-ExchangeSharedMailboxes.ps1` | `U16-Update-ExchangeSharedMailboxes.input.csv` | Exchange | Shared mailbox property updates | Planned |
| U17 | `U17-Set-ExchangeSharedMailboxPermissions.ps1` | `U17-Set-ExchangeSharedMailboxPermissions.input.csv` | Exchange | Add/remove shared mailbox permissions | Planned |
| U18 | `U18-Update-ExchangeResourceMailboxes.ps1` | `U18-Update-ExchangeResourceMailboxes.input.csv` | Exchange | Resource mailbox settings updates | Planned |
| U19 | `U19-Set-ExchangeResourceMailboxBookingDelegates.ps1` | `U19-Set-ExchangeResourceMailboxBookingDelegates.input.csv` | Exchange | Booking delegate/policy changes | Planned |
| U20 | `U20-Set-ExchangeMailboxDelegations.ps1` | `U20-Set-ExchangeMailboxDelegations.input.csv` | Exchange | Add/remove mailbox delegation rights | Planned |
| U21 | `U21-Set-ExchangeMailboxFolderPermissions.ps1` | `U21-Set-ExchangeMailboxFolderPermissions.input.csv` | Exchange | Add/remove folder permissions/delegate flags | Planned |
| U41 | `U41-Set-SharePointSiteAdmins.ps1` | `U41-Set-SharePointSiteAdmins.input.csv` | SharePoint | Add/remove site collection administrators with last-admin safety guard | Implemented |

## Safety and Sequencing Guidance

Recommended execution phases:

1. Entra changes: `U01` to `U08`
2. Teams changes: `U09` to `U12`
3. Exchange changes: `U13` to `U21`

High-impact change classes that should always start with `-WhatIf`:

- Membership removals (`U07`, `U10`, `U12`, `U15`)
- Permission removals (`U17`, `U20`, `U21`)
- Account state changes (`U02`)
- License removals (`U03`)

## Standard Result Columns

Recommended baseline columns:

- `RowNumber`
- `PrimaryKey`
- `Action`
- `Status`
- `Message`

## Related Docs

- [Update README](./README.md)
- [Root README](../README.md)
- [Build Detailed Catalog](../Build/README-Build-Catalog.md)
