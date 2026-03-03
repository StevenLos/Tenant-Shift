# Provision Folder

`Provision` contains initial provisioning scripts and matching CSV templates.

Operational label: **Provision**.

## What Is Here

- Provisioning scripts: `P3001`, `P3002`, `P3005`, `P3006`, `P3008`, `P3309`, `P3113`, `P3114`, `P3115`, `P3116`, `P3118`, `P3119`, `P3240`, `P3242`
- Matching input templates for all provision scripts (`.input.csv`)
- Shared helper module used by Provision, Modify, and InventoryAndReport scripts: `./Common/Online/M365.Common.psm1` (repository-root path)

Modify-style operations are in `Online/Modify/` (`M3003`, `M3204`, `M3007`, `M3310`, `M3311`, `M3312`, `M3113`, `M3114`, `M3115`, `M3116`, `M3117`, `M3118`, `M3119`, `M3120`, `M3121`, `M3122`, `M3123`, `M3241`, `M3243`).

## ID Scheme

- `30xx`: Entra
- `31xx`: Exchange Online
- `32xx`: SharePoint/OneDrive
- `33xx`: Teams

## Provision Catalog

| ID | Script | Workload | Purpose |
|---|---|---|---|
| P3001 | `P3001-Create-EntraUsers.ps1` | Entra | Create cloud users with expanded profile/contact/org/extension fields. |
| P3002 | `P3002-Invite-EntraGuestUsers.ps1` | Entra | Invite guest users. |
| P3005 | `P3005-Create-EntraAssignedSecurityGroups.ps1` | Entra | Create assigned security groups. |
| P3006 | `P3006-Create-EntraDynamicUserSecurityGroups.ps1` | Entra | Create dynamic user security groups. |
| P3008 | `P3008-Create-EntraMicrosoft365Groups.ps1` | Entra | Create Microsoft 365 groups. |
| P3309 | `P3309-Create-MicrosoftTeams.ps1` | Teams | Create Teams. |
| P3113 | `P3113-Create-ExchangeOnlineMailContacts.ps1` | Exchange Online | Create mail contacts. |
| P3114 | `P3114-Create-ExchangeOnlineDistributionLists.ps1` | Exchange Online | Create distribution lists. |
| P3115 | `P3115-Create-ExchangeOnlineMailEnabledSecurityGroups.ps1` | Exchange Online | Create mail-enabled security groups. |
| P3116 | `P3116-Create-ExchangeOnlineSharedMailboxes.ps1` | Exchange Online | Create shared mailboxes. |
| P3118 | `P3118-Create-ExchangeOnlineResourceMailboxes.ps1` | Exchange Online | Create room/equipment mailboxes. |
| P3119 | `P3119-Create-ExchangeOnlineDynamicDistributionGroups.ps1` | Exchange Online | Create dynamic distribution groups. |
| P3240 | `P3240-Create-SharePointSites.ps1` | SharePoint | Create SharePoint sites. |
| P3242 | `P3242-Create-SharePointHubSites.ps1` | SharePoint | Register existing sites as hub sites. |

## Execution Order

Run scripts in numeric order unless there is a specific scoped need.

Execution phases:
- `P3001` to `P3008` (selected): Entra identity and group object creation (`P3001`, `P3002`, `P3005`, `P3006`, `P3008`)
- `P3309`: Team object creation
- `P3113` to `P3119` (selected): Exchange Online recipient/group object creation (`P3113`, `P3114`, `P3115`, `P3116`, `P3118`, `P3119`)
- `P3240`, `P3242`: SharePoint site and hub creation

## Run Pattern

Run from repository root:

```powershell
pwsh ./Online/Provision/P3001-Create-EntraUsers.ps1 -InputCsvPath ./Online/Provision/P3001-Create-EntraUsers.input.csv -WhatIf
pwsh ./Online/Provision/P3001-Create-EntraUsers.ps1 -InputCsvPath ./Online/Provision/P3001-Create-EntraUsers.input.csv

pwsh ./Online/Provision/P3240-Create-SharePointSites.ps1 -InputCsvPath ./Online/Provision/P3240-Create-SharePointSites.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
```

For copy/paste command building, use `./Online/Provision/Provision-Orchestrator.xlsx`.

All scripts write a timestamped `Results_PWWNN-...csv` unless `-OutputCsvPath` is supplied.
All scripts also write a required per-run transcript log (`Transcript_PWWNN-...log`) to the same folder.
Default location: `./Online/Provision/Provision_OutputCsvPath/`.

## Prerequisites

- PowerShell 7+
- Access to PSGallery for module version checks
- Admin permissions for target workload actions
- Required modules by workload:
  - Entra/Graph: `Microsoft.Graph.*`
  - Teams: `MicrosoftTeams`
  - Exchange Online: `ExchangeOnlineManagement`
  - SharePoint/OneDrive: `Microsoft.Online.SharePoint.PowerShell` and tenant admin URL

## Provision Standards

- Every build script must have a matching `.input.csv` template.
- Keep workload explicit in filenames (`Entra`, `ExchangeOnline`, `OneDrive`, `SharePoint`, `MicrosoftTeams`).
- Reuse `./Common/Online/M365.Common.psm1` (repository-root path) for validation, connectivity, retries, and result output.
- Preserve `-WhatIf` behavior for safe dry runs.

## References

- [Root README](../../README.md)
- [Provision Detailed Catalog](./README-Provision-Catalog.md)
- [Entra User Field Contract](../README-Entra-User-Field-Contract.md)











