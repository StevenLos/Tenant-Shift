# Build Detailed Catalog

Detailed catalog of all provisioning scripts in `Build/`.

## Script Contract

All build scripts are expected to:

- Run on PowerShell 7+
- Validate required CSV headers
- Validate required modules and active connections
- Provide console status output
- Support `-WhatIf` for dry runs
- Export per-record results to timestamped CSV
- Retry transient Graph/Exchange failures with exponential backoff (including throttling patterns)
- Use shared helpers from `M365.Common.psm1`

## Catalog

| ID | Script | Input Template | Workload | Purpose | Depends On |
|---|---|---|---|---|---|
| B01 | `B01-Create-EntraUsers.ps1` | `B01-Create-EntraUsers.input.csv` | Entra | Create cloud users. | None |
| B02 | `B02-Invite-EntraGuestUsers.ps1` | `B02-Invite-EntraGuestUsers.input.csv` | Entra | Invite guest users. | None |
| B03 | `B03-Assign-EntraUserLicenses.ps1` | `B03-Assign-EntraUserLicenses.input.csv` | Entra | Assign licenses to users. | B01/B02 users exist |
| B04 | `B04-PreProvision-OneDrive.ps1` | `B04-PreProvision-OneDrive.input.csv` | OneDrive/SharePoint | Pre-provision OneDrive sites. | Licensed users exist |
| B05 | `B05-Create-EntraAssignedSecurityGroups.ps1` | `B05-Create-EntraAssignedSecurityGroups.input.csv` | Entra | Create assigned security groups. | None |
| B06 | `B06-Create-EntraDynamicUserSecurityGroups.ps1` | `B06-Create-EntraDynamicUserSecurityGroups.input.csv` | Entra | Create dynamic user security groups. | None |
| B07 | `B07-Add-EntraUsersToSecurityGroups.ps1` | `B07-Add-EntraUsersToSecurityGroups.input.csv` | Entra | Add users to security groups. | B01/B02 and B05/B06 |
| B08 | `B08-Create-EntraMicrosoft365Groups.ps1` | `B08-Create-EntraMicrosoft365Groups.input.csv` | Entra | Create Microsoft 365 groups. | Optional owners/members exist |
| B09 | `B09-Create-MicrosoftTeams.ps1` | `B09-Create-MicrosoftTeams.input.csv` | Teams | Create Teams. | B08 (or existing M365 group) |
| B10 | `B10-Add-UsersToMicrosoftTeams.ps1` | `B10-Add-UsersToMicrosoftTeams.input.csv` | Teams | Add users to Teams. | B09 and users exist |
| B11 | `B11-Add-ChannelsToMicrosoftTeams.ps1` | `B11-Add-ChannelsToMicrosoftTeams.input.csv` | Teams | Add channels to Teams. | B09 |
| B12 | `B12-Add-UsersToMicrosoftTeamChannels.ps1` | `B12-Add-UsersToMicrosoftTeamChannels.input.csv` | Teams | Add users to private/shared channels. | B09/B11 and users exist |
| B13 | `B13-Create-ExchangeMailContacts.ps1` | `B13-Create-ExchangeMailContacts.input.csv` | Exchange | Create mail contacts. | None |
| B14 | `B14-Create-ExchangeDistributionLists.ps1` | `B14-Create-ExchangeDistributionLists.input.csv` | Exchange | Create distribution lists. | None |
| B15 | `B15-Add-MembersToExchangeDistributionLists.ps1` | `B15-Add-MembersToExchangeDistributionLists.input.csv` | Exchange | Add members to distribution lists. | B14 plus recipients exist |
| B16 | `B16-Create-ExchangeSharedMailboxes.ps1` | `B16-Create-ExchangeSharedMailboxes.input.csv` | Exchange | Create shared mailboxes. | None |
| B17 | `B17-Set-ExchangeSharedMailboxPermissions.ps1` | `B17-Set-ExchangeSharedMailboxPermissions.input.csv` | Exchange | Set shared mailbox permissions. | B16 and users exist |
| B18 | `B18-Create-ExchangeResourceMailboxes.ps1` | `B18-Create-ExchangeResourceMailboxes.input.csv` | Exchange | Create room/equipment mailboxes. | None |
| B19 | `B19-Set-ExchangeResourceMailboxBookingDelegates.ps1` | `B19-Set-ExchangeResourceMailboxBookingDelegates.input.csv` | Exchange | Configure booking delegates/policies. | B18 and users exist |
| B20 | `B20-Set-ExchangeMailboxDelegations.ps1` | `B20-Set-ExchangeMailboxDelegations.input.csv` | Exchange | Assign mailbox delegations. | Target mailbox and users exist |
| B21 | `B21-Set-ExchangeMailboxFolderPermissions.ps1` | `B21-Set-ExchangeMailboxFolderPermissions.input.csv` | Exchange | Assign folder-level permissions/delegation flags. | Target mailbox and users exist |
| B40 | `B40-Create-SharePointSites.ps1` | `B40-Create-SharePointSites.input.csv` | SharePoint | Create SharePoint sites from CSV. | Site owners exist; SharePoint admin URL |
| B41 | `B41-Set-SharePointSiteAdmins.ps1` | `B41-Set-SharePointSiteAdmins.input.csv` | SharePoint | Add/remove SharePoint site collection admins. | B40 sites exist |
| B42 | `B42-Create-SharePointHubSites.ps1` | `B42-Create-SharePointHubSites.input.csv` | SharePoint | Register existing sites as hub sites. | B40 sites exist |
| B43 | `B43-Associate-SharePointSitesToHub.ps1` | `B43-Associate-SharePointSitesToHub.input.csv` | SharePoint | Associate sites to hub sites. | B42 hubs and B40 sites exist |

## Recommended Run Sequence

1. Identity baseline: `B01` to `B08`
2. Teams baseline: `B09` to `B12`
3. Exchange baseline: `B13` to `B21`
4. SharePoint baseline: `B40` to `B43`

## Known High-Dependency Steps

- `B04` requires valid `-SharePointAdminUrl` and licensed users.
- `B12` is meaningful only for private/shared channels (standard channels inherit Team membership).
- `B17`, `B19`, `B20`, and `B21` assume target mailboxes already exist.
- `B40`, `B41`, `B42`, and `B43` require valid `-SharePointAdminUrl` (`https://<tenant>-admin.sharepoint.com`).
- `B43` supports either `HubSiteUrl` or `HubSiteId`; `AllowReassociation` is optional and defaults to `FALSE`.

## Input and Behavior Notes

- The module version check requires access to PSGallery (`Find-Module`).
- If an input CSV has only headers and no data rows, the script stops with a validation message.
- Multi-value fields use semicolon-delimited values (example: `user1@contoso.com;user2@contoso.com`).
- Input templates include one sample data row.
- `B17-Set-ExchangeSharedMailboxPermissions.input.csv` supports `ReadOnly` in addition to `FullAccess`, `SendAs`, and `SendOnBehalf`.
- `B03-Assign-EntraUserLicenses.input.csv` supports `DisabledPlans` as service-plan names or GUIDs separated by semicolons.
- `B02-Invite-EntraGuestUsers.input.csv` supports optional custom message body and semicolon-delimited `CcRecipients`.
- `B04-PreProvision-OneDrive.ps1` requires `-SharePointAdminUrl` (example: `https://contoso-admin.sharepoint.com`).
- `B13-Create-ExchangeMailContacts.input.csv` uses `ExternalEmailAddress` and supports optional GAL hiding.
- `B06-Create-EntraDynamicUserSecurityGroups.input.csv` requires a valid `MembershipRule`; `MembershipRuleProcessingState` accepts `On` or `Paused`.
- `B08-Create-EntraMicrosoft365Groups.input.csv` supports optional semicolon-delimited owner/member UPN lists and `Visibility` values `Private` or `Public`.
- `B18-Create-ExchangeResourceMailboxes.input.csv` uses `ResourceType` values `Room` or `Equipment`.
- `B20-Set-ExchangeMailboxDelegations.input.csv` supports `FullAccess`, `SendAs`, `SendOnBehalf`, and `AutoMapping`.
- `B21-Set-ExchangeMailboxFolderPermissions.input.csv` supports folder-level permissions and calendar delegate flags (`CalendarDelegate`, `CalendarCanViewPrivateItems`).
- `B19-Set-ExchangeResourceMailboxBookingDelegates.input.csv` sets resource booking delegates and core booking-policy flags.
- `B09-Create-MicrosoftTeams.input.csv` can create a Team from an existing or new Microsoft 365 group and supports optional owner/member UPN lists with core Team settings.
- `B10-Add-UsersToMicrosoftTeams.input.csv` adds users by `TeamMailNickname`; `Role` supports `Member` or `Owner`.
- `B11-Add-ChannelsToMicrosoftTeams.input.csv` adds channels by `TeamMailNickname`; `MembershipType` supports `Standard`, `Private`, and `Shared`.
- `B12-Add-UsersToMicrosoftTeamChannels.input.csv` assigns users to existing private/shared channels with `Role` values `Member` or `Owner`.
- `B40-Create-SharePointSites.input.csv` supports optional secondary owners, locale (`Language`), `TimeZoneId`, and `StorageQuotaMB`.
- `B41-Set-SharePointSiteAdmins.input.csv` supports idempotent add/remove and defaults `EnsurePrimaryOwnerIsAdmin` to `TRUE`.
- `B42-Create-SharePointHubSites.input.csv` registers hubs idempotently and attempts metadata updates when supported by module version.
- `B43-Associate-SharePointSitesToHub.input.csv` associates sites to hubs with optional reassociation via `AllowReassociation`.

## Run Example

```powershell
pwsh ./Build/B03-Assign-EntraUserLicenses.ps1 -InputCsvPath ./Build/B03-Assign-EntraUserLicenses.input.csv -WhatIf
```

## Related Docs

- [Build README](./README.md)
- [Root README](../README.md)
