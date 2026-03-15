# Provision Detailed Catalog

Detailed catalog of all provisioning scripts in `SharedModule/Online/Provision/`.

Operational label: **Provision**.

## Script Contract

All provision scripts are expected to:

- Run on PowerShell 7+
- Validate required CSV headers
- Validate required modules and active connections
- Provide console status output and required per-run transcript logging
- Support `-WhatIf` for dry runs
- Export per-record results to timestamped CSV
- Retry transient Graph/Exchange Online failures with exponential backoff (including throttling patterns)
- Use shared helpers from `./SharedModule/Common/Online/M365.Common.psm1` (repository-root path)

## ID Ranges

- `30xx`: Entra
- `31xx`: Exchange Online
- `32xx`: SharePoint/OneDrive
- `33xx`: Teams

## Catalog

| ID | Script | Input Template | Workload | Purpose | Depends On |
|---|---|---|---|---|---|
| P3001 | `SM-P3001-Create-EntraUsers.ps1` | `SM-P3001-Create-EntraUsers.input.csv` | Entra | Create cloud users with expanded profile/contact/org/extension fields. | None |
| P3002 | `SM-P3002-Invite-EntraGuestUsers.ps1` | `SM-P3002-Invite-EntraGuestUsers.input.csv` | Entra | Invite guest users. | None |
| P3005 | `SM-P3005-Create-EntraAssignedSecurityGroups.ps1` | `SM-P3005-Create-EntraAssignedSecurityGroups.input.csv` | Entra | Create assigned security groups. | None |
| P3006 | `SM-P3006-Create-EntraDynamicUserSecurityGroups.ps1` | `SM-P3006-Create-EntraDynamicUserSecurityGroups.input.csv` | Entra | Create dynamic user security groups. | None |
| P3008 | `SM-P3008-Create-EntraMicrosoft365Groups.ps1` | `SM-P3008-Create-EntraMicrosoft365Groups.input.csv` | Entra | Create Microsoft 365 groups. | Optional owners/members exist |
| P3309 | `SM-P3309-Create-MicrosoftTeams.ps1` | `SM-P3309-Create-MicrosoftTeams.input.csv` | Teams | Create Teams. | P3008 (or existing M365 group) |
| P3113 | `SM-P3113-Create-ExchangeOnlineMailContacts.ps1` | `SM-P3113-Create-ExchangeOnlineMailContacts.input.csv` | Exchange Online | Create mail contacts. | None |
| P3114 | `SM-P3114-Create-ExchangeOnlineDistributionLists.ps1` | `SM-P3114-Create-ExchangeOnlineDistributionLists.input.csv` | Exchange Online | Create distribution lists. | None |
| P3115 | `SM-P3115-Create-ExchangeOnlineMailEnabledSecurityGroups.ps1` | `SM-P3115-Create-ExchangeOnlineMailEnabledSecurityGroups.input.csv` | Exchange Online | Create mail-enabled security groups. | None |
| P3116 | `SM-P3116-Create-ExchangeOnlineSharedMailboxes.ps1` | `SM-P3116-Create-ExchangeOnlineSharedMailboxes.input.csv` | Exchange Online | Create shared mailboxes. | None |
| P3118 | `SM-P3118-Create-ExchangeOnlineResourceMailboxes.ps1` | `SM-P3118-Create-ExchangeOnlineResourceMailboxes.input.csv` | Exchange Online | Create room/equipment mailboxes. | None |
| P3119 | `SM-P3119-Create-ExchangeOnlineDynamicDistributionGroups.ps1` | `SM-P3119-Create-ExchangeOnlineDynamicDistributionGroups.input.csv` | Exchange Online | Create dynamic distribution groups. | None |
| P3124 | `SM-P3124-Create-ExchangeOnlineAcceptedDomains.ps1` | `SM-P3124-Create-ExchangeOnlineAcceptedDomains.input.csv` | Exchange Online | Create/update accepted domains and optionally create matching Entra tenant domains. | Graph Domain write permissions |
| P3240 | `SM-P3240-Create-SharePointSites.ps1` | `SM-P3240-Create-SharePointSites.input.csv` | SharePoint | Create SharePoint sites from CSV. | Site owners exist; SharePoint admin URL |
| P3242 | `SM-P3242-Create-SharePointHubSites.ps1` | `SM-P3242-Create-SharePointHubSites.input.csv` | SharePoint | Register existing sites as hub sites. | P3240 sites exist |

## Recommended Run Sequence

1. Identity baseline: `P3001`, `P3002`, `P3005`, `P3006`, `P3008`
2. Teams baseline: `P3309`
3. Exchange Online baseline: `P3113`, `P3114`, `P3115`, `P3116`, `P3118`, `P3119`, `P3124`
4. SharePoint baseline: `P3240`, `P3242`

## Known High-Dependency Steps

- `P3309` depends on a Microsoft 365 group (`P3008`) or an existing group.
- `P3240` requires valid `-SharePointAdminUrl` (`https://<tenant>-admin.sharepoint.com`).
- `P3242` requires `P3240` sites to already exist and valid `-SharePointAdminUrl`.

## Input and Behavior Notes

- The module version check requires access to PSGallery (`Find-Module`).
- If an input CSV has only headers and no data rows, the script stops with a validation message.
- Multi-value fields use semicolon-delimited values (example: `user1@contoso.com;user2@contoso.com`).
- Input templates include one sample data row.
- `SM-P3002-Invite-EntraGuestUsers.input.csv` supports optional custom message body and semicolon-delimited `CcRecipients`.
- `SM-P3006-Create-EntraDynamicUserSecurityGroups.input.csv` requires a valid `MembershipRule`; `MembershipRuleProcessingState` accepts `On` or `Paused`.
- `SM-P3008-Create-EntraMicrosoft365Groups.input.csv` supports optional semicolon-delimited owner/member UPN lists and `Visibility` values `Private` or `Public`.
- `SM-P3309-Create-MicrosoftTeams.input.csv` can create a Team from an existing or new Microsoft 365 group and supports optional owner/member UPN lists with core Team settings.
- `SM-P3113-Create-ExchangeOnlineMailContacts.input.csv` uses `ExternalEmailAddress` and supports optional GAL hiding.
- `SM-P3114-Create-ExchangeOnlineDistributionLists.input.csv` supports sender allow/deny lists and moderation notification settings.
- `SM-P3115-Create-ExchangeOnlineMailEnabledSecurityGroups.input.csv` creates mail-enabled security groups and supports moderation/sender restrictions.
- `SM-P3116-Create-ExchangeOnlineSharedMailboxes.input.csv` supports optional `HiddenFromAddressListsEnabled`, send-on-behalf list, sent-item copy controls, forwarding, and compliance toggles.
- `SM-P3118-Create-ExchangeOnlineResourceMailboxes.input.csv` uses `ResourceType` values `Room` or `Equipment` and supports advanced booking policy fields.
- `SM-P3119-Create-ExchangeOnlineDynamicDistributionGroups.input.csv` supports `RecipientFilter` or `IncludedRecipients` with conditional attributes.
- `SM-P3124-Create-ExchangeOnlineAcceptedDomains.input.csv` supports Entra tenant-domain auto-create, accepted-domain type/default handling, and optional match-subdomain behavior where available.
- `SM-P3240-Create-SharePointSites.input.csv` supports optional secondary owners, locale (`Language`), `TimeZoneId`, and `StorageQuotaMB`.
- `SM-P3242-Create-SharePointHubSites.input.csv` registers hubs idempotently and applies optional metadata when supported by module version.

## Run Example

```powershell
pwsh ./SharedModule/Online/Provision/SM-P3001-Create-EntraUsers.ps1 -InputCsvPath ./SharedModule/Online/Provision/SM-P3001-Create-EntraUsers.input.csv -WhatIf
```

## Related Docs

- [Provision README](./README.md)
- [SharedModule README](../../README.md)
- [Modify Detailed Catalog](../Modify/README-Modify-Catalog.md)








