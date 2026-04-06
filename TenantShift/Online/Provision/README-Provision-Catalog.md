# Provision Detailed Catalog

Detailed catalog of all provisioning scripts in `TenantShift/Online/Provision/`.

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
- Use shared helpers from `./TenantShift/Common/Online/M365.Common.psm1` (repository-root path)

## Catalog

| ID | Script | Input Template | Workload | Purpose | Depends On |
|---|---|---|---|---|---|
| P-MEID-0010 | `P-MEID-0010-Create-EntraUsers.ps1` | `P-MEID-0010-Create-EntraUsers.input.csv` | MEID | Create cloud users with expanded profile/contact/org/extension fields. | None |
| P-MEID-0020 | `P-MEID-0020-Invite-EntraGuestUsers.ps1` | `P-MEID-0020-Invite-EntraGuestUsers.input.csv` | MEID | Invite guest users. | None |
| P-MEID-0070 | `P-MEID-0070-Create-EntraAssignedSecurityGroups.ps1` | `P-MEID-0070-Create-EntraAssignedSecurityGroups.input.csv` | MEID | Create assigned security groups. | None |
| P-MEID-0080 | `P-MEID-0080-Create-EntraDynamicUserSecurityGroups.ps1` | `P-MEID-0080-Create-EntraDynamicUserSecurityGroups.input.csv` | MEID | Create dynamic user security groups. | None |
| P-MEID-0100 | `P-MEID-0100-Create-EntraMicrosoft365Groups.ps1` | `P-MEID-0100-Create-EntraMicrosoft365Groups.input.csv` | MEID | Create Microsoft 365 groups. | Optional owners/members exist |
| P-EXOL-0010 | `P-EXOL-0010-Create-ExchangeOnlineAcceptedDomains.ps1` | `P-EXOL-0010-Create-ExchangeOnlineAcceptedDomains.input.csv` | EXOL | Create/update accepted domains and optionally create matching Entra tenant domains. | Graph Domain write permissions |
| P-EXOL-0030 | `P-EXOL-0030-Create-ExchangeOnlineMailContacts.ps1` | `P-EXOL-0030-Create-ExchangeOnlineMailContacts.input.csv` | EXOL | Create mail contacts. | None |
| P-EXOL-0040 | `P-EXOL-0040-Create-ExchangeOnlineDistributionLists.ps1` | `P-EXOL-0040-Create-ExchangeOnlineDistributionLists.input.csv` | EXOL | Create distribution lists. | None |
| P-EXOL-0050 | `P-EXOL-0050-Create-ExchangeOnlineMailEnabledSecurityGroups.ps1` | `P-EXOL-0050-Create-ExchangeOnlineMailEnabledSecurityGroups.input.csv` | EXOL | Create mail-enabled security groups. | None |
| P-EXOL-0060 | `P-EXOL-0060-Create-ExchangeOnlineDynamicDistributionGroups.ps1` | `P-EXOL-0060-Create-ExchangeOnlineDynamicDistributionGroups.input.csv` | EXOL | Create dynamic distribution groups. | None |
| P-EXOL-0070 | `P-EXOL-0070-Create-ExchangeOnlineSharedMailboxes.ps1` | `P-EXOL-0070-Create-ExchangeOnlineSharedMailboxes.input.csv` | EXOL | Create shared mailboxes. | None |
| P-EXOL-0080 | `P-EXOL-0080-Create-ExchangeOnlineResourceMailboxes.ps1` | `P-EXOL-0080-Create-ExchangeOnlineResourceMailboxes.input.csv` | EXOL | Create room/equipment mailboxes. | None |
| P-SPOL-0010 | `P-SPOL-0010-Create-SharePointSites.ps1` | `P-SPOL-0010-Create-SharePointSites.input.csv` | SPOL | Create SharePoint sites from CSV. | Site owners exist; SharePoint admin URL |
| P-SPOL-0020 | `P-SPOL-0020-Create-SharePointHubSites.ps1` | `P-SPOL-0020-Create-SharePointHubSites.input.csv` | SPOL | Register existing sites as hub sites. | `P-SPOL-0010` sites exist |
| P-TEAM-0010 | `P-TEAM-0010-Create-MicrosoftTeams.ps1` | `P-TEAM-0010-Create-MicrosoftTeams.input.csv` | TEAM | Create Teams. | `P-MEID-0100` (or existing M365 group) |

## Recommended Run Sequence

1. MEID identity baseline: `P-MEID-0010`, `P-MEID-0020`, `P-MEID-0070`, `P-MEID-0080`, `P-MEID-0100`
2. TEAM baseline: `P-TEAM-0010`
3. EXOL baseline: `P-EXOL-0010`, `P-EXOL-0030`, `P-EXOL-0040`, `P-EXOL-0050`, `P-EXOL-0060`, `P-EXOL-0070`, `P-EXOL-0080`
4. SPOL baseline: `P-SPOL-0010`, `P-SPOL-0020`

## Known High-Dependency Steps

- `P-TEAM-0010` depends on a Microsoft 365 group (`P-MEID-0100`) or an existing group.
- `P-SPOL-0010` requires valid `-SharePointAdminUrl` (`https://<tenant>-admin.sharepoint.com`).
- `P-SPOL-0020` requires `P-SPOL-0010` sites to already exist and valid `-SharePointAdminUrl`.

## Input and Behavior Notes

- The module version check requires access to PSGallery (`Find-Module`).
- If an input CSV has only headers and no data rows, the script stops with a validation message.
- Multi-value fields use semicolon-delimited values (example: `user1@contoso.com;user2@contoso.com`).
- Input templates include one sample data row.
- `P-MEID-0020-Invite-EntraGuestUsers.input.csv` supports optional custom message body and semicolon-delimited `CcRecipients`.
- `P-MEID-0080-Create-EntraDynamicUserSecurityGroups.input.csv` requires a valid `MembershipRule`; `MembershipRuleProcessingState` accepts `On` or `Paused`.
- `P-MEID-0100-Create-EntraMicrosoft365Groups.input.csv` supports optional semicolon-delimited owner/member UPN lists and `Visibility` values `Private` or `Public`.
- `P-TEAM-0010-Create-MicrosoftTeams.input.csv` can create a Team from an existing or new Microsoft 365 group and supports optional owner/member UPN lists with core Team settings.
- `P-EXOL-0030-Create-ExchangeOnlineMailContacts.input.csv` uses `ExternalEmailAddress` and supports optional GAL hiding.
- `P-EXOL-0040-Create-ExchangeOnlineDistributionLists.input.csv` supports sender allow/deny lists and moderation notification settings.
- `P-EXOL-0050-Create-ExchangeOnlineMailEnabledSecurityGroups.input.csv` creates mail-enabled security groups and supports moderation/sender restrictions.
- `P-EXOL-0070-Create-ExchangeOnlineSharedMailboxes.input.csv` supports optional `HiddenFromAddressListsEnabled`, send-on-behalf list, sent-item copy controls, forwarding, and compliance toggles.
- `P-EXOL-0080-Create-ExchangeOnlineResourceMailboxes.input.csv` uses `ResourceType` values `Room` or `Equipment` and supports advanced booking policy fields.
- `P-EXOL-0060-Create-ExchangeOnlineDynamicDistributionGroups.input.csv` supports `RecipientFilter` or `IncludedRecipients` with conditional attributes.
- `P-EXOL-0010-Create-ExchangeOnlineAcceptedDomains.input.csv` supports Entra tenant-domain auto-create, accepted-domain type/default handling, and optional match-subdomain behavior where available.
- `P-SPOL-0010-Create-SharePointSites.input.csv` supports optional secondary owners, locale (`Language`), `TimeZoneId`, and `StorageQuotaMB`.
- `P-SPOL-0020-Create-SharePointHubSites.input.csv` registers hubs idempotently and applies optional metadata when supported by module version.

## Run Example

```powershell
pwsh ./TenantShift/Online/Provision/P-MEID-0010-Create-EntraUsers.ps1 -InputCsvPath ./TenantShift/Online/Provision/P-MEID-0010-Create-EntraUsers.input.csv -WhatIf
```

## Related Docs

- [Provision README](./README.md)
- [Root README](../../README.md)
- [Modify Detailed Catalog](../Modify/README-Modify-Catalog.md)
