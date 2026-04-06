# Provision Folder

`Provision` contains initial provisioning scripts and matching CSV templates.

Operational label: **Provision**.

## What Is Here

- Provisioning scripts: `P-MEID-0010`, `P-MEID-0020`, `P-MEID-0070`, `P-MEID-0080`, `P-MEID-0100`, `P-TEAM-0010`, `P-EXOL-0030`, `P-EXOL-0040`, `P-EXOL-0050`, `P-EXOL-0070`, `P-EXOL-0080`, `P-EXOL-0060`, `P-EXOL-0010`, `P-SPOL-0010`, `P-SPOL-0020`
- Matching input templates for all provision scripts (`.input.csv`)
- Shared helper module used by Provision, Modify, and Discover scripts: `./TenantShift/Common/Online/M365.Common.psm1` (repository-root path)

Modify-style operations are in `TenantShift/Online/Modify/` (`M-MEID-0030`, `M-ONDR-0010`, `M-MEID-0090`, `M-TEAM-0020`, `M-TEAM-0030`, `M-TEAM-0040`, `M-EXOL-0030`, `M-EXOL-0040`, `M-EXOL-0090`, `M-EXOL-0070`, `M-EXOL-0110`, `M-EXOL-0080`, `M-EXOL-0120`, `M-EXOL-0130`, `M-EXOL-0140`, `M-EXOL-0050`, `M-EXOL-0060`, `M-EXOL-0150`, `M-EXOL-0160`, `M-EXOL-0170`, `M-EXOL-0180`, `M-EXOL-0190`, `M-EXOL-0200`, `M-EXOL-0010`, `M-EXOL-0020`, `M-SPOL-0030`, `M-SPOL-0040`).

## Workload Code Allocation

- `MEID`: Entra
- `EXOL`: Exchange Online
- `SPOL`: SharePoint
- `TEAM`: Teams

## Provision Catalog

| ID | Script | Workload | Purpose |
|---|---|---|---|
| P-MEID-0010 | `P-MEID-0010-Create-EntraUsers.ps1` | Entra | Create cloud users with expanded profile/contact/org/extension fields. |
| P-MEID-0020 | `P-MEID-0020-Invite-EntraGuestUsers.ps1` | Entra | Invite guest users. |
| P-MEID-0070 | `P-MEID-0070-Create-EntraAssignedSecurityGroups.ps1` | Entra | Create assigned security groups. |
| P-MEID-0080 | `P-MEID-0080-Create-EntraDynamicUserSecurityGroups.ps1` | Entra | Create dynamic user security groups. |
| P-MEID-0100 | `P-MEID-0100-Create-EntraMicrosoft365Groups.ps1` | Entra | Create Microsoft 365 groups. |
| P-TEAM-0010 | `P-TEAM-0010-Create-MicrosoftTeams.ps1` | Teams | Create Teams. |
| P-EXOL-0030 | `P-EXOL-0030-Create-ExchangeOnlineMailContacts.ps1` | Exchange Online | Create mail contacts. |
| P-EXOL-0040 | `P-EXOL-0040-Create-ExchangeOnlineDistributionLists.ps1` | Exchange Online | Create distribution lists. |
| P-EXOL-0050 | `P-EXOL-0050-Create-ExchangeOnlineMailEnabledSecurityGroups.ps1` | Exchange Online | Create mail-enabled security groups. |
| P-EXOL-0070 | `P-EXOL-0070-Create-ExchangeOnlineSharedMailboxes.ps1` | Exchange Online | Create shared mailboxes. |
| P-EXOL-0080 | `P-EXOL-0080-Create-ExchangeOnlineResourceMailboxes.ps1` | Exchange Online | Create room/equipment mailboxes. |
| P-EXOL-0060 | `P-EXOL-0060-Create-ExchangeOnlineDynamicDistributionGroups.ps1` | Exchange Online | Create dynamic distribution groups. |
| P-EXOL-0010 | `P-EXOL-0010-Create-ExchangeOnlineAcceptedDomains.ps1` | Exchange Online | Create/update accepted domains and optionally create matching Entra tenant domains. |
| P-SPOL-0010 | `P-SPOL-0010-Create-SharePointSites.ps1` | SharePoint | Create SharePoint sites. |
| P-SPOL-0020 | `P-SPOL-0020-Create-SharePointHubSites.ps1` | SharePoint | Register existing sites as hub sites. |

## Execution Order

Run scripts in workload phase order unless there is a specific scoped need.

Execution phases:
- `P-MEID-0010`, `P-MEID-0020`, `P-MEID-0070`, `P-MEID-0080`, `P-MEID-0100`: Entra identity and group object creation
- `P-TEAM-0010`: Team object creation
- `P-EXOL-0010`, `P-EXOL-0030`, `P-EXOL-0040`, `P-EXOL-0050`, `P-EXOL-0060`, `P-EXOL-0070`, `P-EXOL-0080`: Exchange Online recipient/group/domain object creation
- `P-SPOL-0010`, `P-SPOL-0020`: SharePoint site and hub creation

## Run Pattern

Run from repository root:

```powershell
pwsh ./TenantShift/Online/Provision/P-MEID-0010-Create-EntraUsers.ps1 -InputCsvPath ./TenantShift/Online/Provision/P-MEID-0010-Create-EntraUsers.input.csv -WhatIf
pwsh ./TenantShift/Online/Provision/P-MEID-0010-Create-EntraUsers.ps1 -InputCsvPath ./TenantShift/Online/Provision/P-MEID-0010-Create-EntraUsers.input.csv

pwsh ./TenantShift/Online/Provision/P-SPOL-0010-Create-SharePointSites.ps1 -InputCsvPath ./TenantShift/Online/Provision/P-SPOL-0010-Create-SharePointSites.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com -WhatIf
```

For copy/paste command building, use `./TenantShift/Online/Provision/Provision-Orchestrator.xlsx`.

All scripts write a timestamped `Results_P-<WW>-<NNNN>-...csv` unless `-OutputCsvPath` is supplied.
All scripts also write a required per-run transcript log (`Transcript_P-<WW>-<NNNN>-...log`) to the same folder.
Default location: `./TenantShift/Online/Provision/Provision_OutputCsvPath/`.

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
- Reuse `./TenantShift/Common/Online/M365.Common.psm1` (repository-root path) for validation, connectivity, retries, and result output.
- Preserve `-WhatIf` behavior for safe dry runs.

## References

- [Root README](../../README.md)
- [Provision Detailed Catalog](./README-Provision-Catalog.md)
- [Entra User Field Contract](../README-Entra-User-Field-Contract.md)
- [Operator Runbook](./RUNBOOK-Provision.md)
