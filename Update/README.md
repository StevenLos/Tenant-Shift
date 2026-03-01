# Update Folder

`Update` is for scripts that modify existing tenant objects/configuration.

Current status: initial implementation started (`U41-Set-SharePointSiteAdmins.ps1` added).

## Purpose

Use this folder for controlled change operations after initial provisioning:

- Attribute updates
- Membership changes
- Policy/permission changes
- Lifecycle changes on existing objects

Do not use this folder for:
- Initial object creation (use `Build`)
- Read-only reporting (use `Discover`)

## Naming Standard

- Script: `U##-<Action>-<Target>.ps1`
- Input template: `U##-<Action>-<Target>.input.csv`
- Output pattern: `Results_U##-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`

Example:
- `U01-Update-EntraUsers.ps1`
- `U01-Update-EntraUsers.input.csv`

## Run Pattern

Run from repository root:

```powershell
pwsh ./Update/U01-Update-EntraUsers.ps1 -InputCsvPath ./Update/U01-Update-EntraUsers.input.csv -WhatIf
pwsh ./Update/U01-Update-EntraUsers.ps1 -InputCsvPath ./Update/U01-Update-EntraUsers.input.csv

pwsh ./Update/U41-Set-SharePointSiteAdmins.ps1 -InputCsvPath ./Update/U41-Set-SharePointSiteAdmins.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
```

## Required Safety Model

All update scripts should include:

- `-WhatIf` support and clear `ShouldProcess` messaging
- Idempotent behavior where practical
- Per-record validation and error capture
- Timestamped result export with `Status` and `Message`
- Clear rollback or remediation notes for high-impact changes

## Proposed Initial Update Backlog

| ID | Proposed Script | Workload | Purpose |
|---|---|---|---|
| U01 | `U01-Update-EntraUsers.ps1` | Entra | Update user profile attributes. |
| U02 | `U02-Set-EntraUserAccountState.ps1` | Entra | Enable/disable user accounts. |
| U03 | `U03-Set-EntraUserLicenses.ps1` | Entra | Add/remove user licenses. |
| U04 | `U04-Set-OneDriveSiteSettings.ps1` | OneDrive/SharePoint | Update OneDrive site settings/quota. |
| U05 | `U05-Update-EntraAssignedSecurityGroups.ps1` | Entra | Update assigned security group properties. |
| U06 | `U06-Update-EntraDynamicUserSecurityGroups.ps1` | Entra | Update dynamic membership rules/settings. |
| U07 | `U07-Set-EntraSecurityGroupMembers.ps1` | Entra | Add/remove security group members. |
| U08 | `U08-Update-EntraMicrosoft365Groups.ps1` | Entra | Update M365 group properties/visibility. |
| U09 | `U09-Update-MicrosoftTeams.ps1` | Teams | Update Team settings. |
| U10 | `U10-Set-MicrosoftTeamMembers.ps1` | Teams | Add/remove Team owners/members. |
| U11 | `U11-Update-MicrosoftTeamChannels.ps1` | Teams | Update channel settings. |
| U12 | `U12-Set-MicrosoftTeamChannelMembers.ps1` | Teams | Add/remove channel members. |
| U13 | `U13-Update-ExchangeMailContacts.ps1` | Exchange | Update mail contact attributes. |
| U14 | `U14-Update-ExchangeDistributionLists.ps1` | Exchange | Update DL properties. |
| U15 | `U15-Set-ExchangeDistributionListMembers.ps1` | Exchange | Add/remove DL members. |
| U16 | `U16-Update-ExchangeSharedMailboxes.ps1` | Exchange | Update shared mailbox settings. |
| U17 | `U17-Set-ExchangeSharedMailboxPermissions.ps1` | Exchange | Add/remove shared mailbox permissions. |
| U18 | `U18-Update-ExchangeResourceMailboxes.ps1` | Exchange | Update room/equipment mailbox settings. |
| U19 | `U19-Set-ExchangeResourceMailboxBookingDelegates.ps1` | Exchange | Update booking delegates/policy flags. |
| U20 | `U20-Set-ExchangeMailboxDelegations.ps1` | Exchange | Add/remove mailbox delegation entries. |
| U21 | `U21-Set-ExchangeMailboxFolderPermissions.ps1` | Exchange | Add/remove folder-level permissions. |
| U41 | `U41-Set-SharePointSiteAdmins.ps1` | SharePoint | Add/remove site collection administrators. |

## Update Standards

- Keep workload explicit in script names.
- Include matched `.input.csv` templates for repeatable change sets.
- Reuse `M365.Common.psm1` for shared validation, connectivity, and result handling.
- Support bulk operations with per-record outcome tracking.

## References

- [Root README](../README.md)
- [Build README](../Build/README.md)
- [Update Detailed Catalog](./README-Update-Catalog.md)
