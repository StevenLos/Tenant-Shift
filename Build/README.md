# Build Folder

`Build` contains initial provisioning scripts and matching CSV templates.

## What Is Here

- Provisioning scripts: `B01` through `B21`, plus `B40` through `B43`
- Matching input templates for all build scripts (`.input.csv`)
- Shared helper module: `M365.Common.psm1`

## Build Catalog

| ID | Script | Workload | Purpose |
|---|---|---|---|
| B01 | `B01-Create-EntraUsers.ps1` | Entra | Create cloud users. |
| B02 | `B02-Invite-EntraGuestUsers.ps1` | Entra | Invite guest users. |
| B03 | `B03-Assign-EntraUserLicenses.ps1` | Entra | Assign user licenses. |
| B04 | `B04-PreProvision-OneDrive.ps1` | OneDrive/SharePoint | Pre-provision OneDrive sites. |
| B05 | `B05-Create-EntraAssignedSecurityGroups.ps1` | Entra | Create assigned security groups. |
| B06 | `B06-Create-EntraDynamicUserSecurityGroups.ps1` | Entra | Create dynamic user security groups. |
| B07 | `B07-Add-EntraUsersToSecurityGroups.ps1` | Entra | Add users to security groups. |
| B08 | `B08-Create-EntraMicrosoft365Groups.ps1` | Entra | Create Microsoft 365 groups. |
| B09 | `B09-Create-MicrosoftTeams.ps1` | Teams | Create Teams. |
| B10 | `B10-Add-UsersToMicrosoftTeams.ps1` | Teams | Add users to Teams. |
| B11 | `B11-Add-ChannelsToMicrosoftTeams.ps1` | Teams | Add channels to Teams. |
| B12 | `B12-Add-UsersToMicrosoftTeamChannels.ps1` | Teams | Add users to private/shared channels. |
| B13 | `B13-Create-ExchangeMailContacts.ps1` | Exchange | Create mail contacts. |
| B14 | `B14-Create-ExchangeDistributionLists.ps1` | Exchange | Create distribution lists. |
| B15 | `B15-Add-MembersToExchangeDistributionLists.ps1` | Exchange | Add members to distribution lists. |
| B16 | `B16-Create-ExchangeSharedMailboxes.ps1` | Exchange | Create shared mailboxes. |
| B17 | `B17-Set-ExchangeSharedMailboxPermissions.ps1` | Exchange | Set shared mailbox permissions. |
| B18 | `B18-Create-ExchangeResourceMailboxes.ps1` | Exchange | Create room/equipment mailboxes. |
| B19 | `B19-Set-ExchangeResourceMailboxBookingDelegates.ps1` | Exchange | Set resource mailbox booking delegates. |
| B20 | `B20-Set-ExchangeMailboxDelegations.ps1` | Exchange | Set mailbox delegations (full/mail send). |
| B21 | `B21-Set-ExchangeMailboxFolderPermissions.ps1` | Exchange | Set folder-level permissions/delegates. |
| B40 | `B40-Create-SharePointSites.ps1` | SharePoint | Create SharePoint sites. |
| B41 | `B41-Set-SharePointSiteAdmins.ps1` | SharePoint | Set SharePoint site collection administrators. |
| B42 | `B42-Create-SharePointHubSites.ps1` | SharePoint | Register SharePoint hub sites. |
| B43 | `B43-Associate-SharePointSitesToHub.ps1` | SharePoint | Associate SharePoint sites to hubs. |

## Execution Order

Run scripts in numeric order (`B01` to `B21`) unless there is a specific scoped need.

Execution phases:
- `B01` to `B08`: Entra identity, licensing, and groups
- `B09` to `B12`: Teams provisioning and membership
- `B13` to `B21`: Exchange recipients, permissions, and delegation
- `B40` to `B43`: SharePoint site and hub architecture

## Run Pattern

Run from repository root:

```powershell
pwsh ./Build/B01-Create-EntraUsers.ps1 -InputCsvPath ./Build/B01-Create-EntraUsers.input.csv -WhatIf
pwsh ./Build/B01-Create-EntraUsers.ps1 -InputCsvPath ./Build/B01-Create-EntraUsers.input.csv

pwsh ./Build/B40-Create-SharePointSites.ps1 -InputCsvPath ./Build/B40-Create-SharePointSites.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
```

All scripts write a timestamped `Results_B##-...csv` unless `-OutputCsvPath` is supplied.

## Prerequisites

- PowerShell 7+
- Access to PSGallery for module version checks
- Admin permissions for target workload actions
- Required modules by workload:
  - Entra/Graph: `Microsoft.Graph.*`
  - Teams: `MicrosoftTeams`
  - Exchange: `ExchangeOnlineManagement`
  - SharePoint/OneDrive: `Microsoft.Online.SharePoint.PowerShell` and tenant admin URL

## Build Standards

- Every build script must have a matching `.input.csv` template.
- Keep workload explicit in filenames (`Entra`, `Exchange`, `OneDrive`, `SharePoint`, `MicrosoftTeams`).
- Reuse `M365.Common.psm1` for validation, connectivity, retries, and result output.
- Preserve `-WhatIf` behavior for safe dry runs.

## References

- [Root README](../README.md)
- [Build Detailed Catalog](./README-Build-Catalog.md)
