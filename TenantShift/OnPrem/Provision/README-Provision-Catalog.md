# Provision Detailed Catalog

Detailed catalog of provisioning scripts in `TenantShift/OnPrem/Provision/`.

Operational label: **Provision**.

Current implementation status: partial. ADUC provision baseline scripts (`P-ADUC-0010` through `P-ADUC-0050`) and EXOP provision scripts (`P-EXOP-0010` through `P-EXOP-0060`) are implemented; GPOL and FILE remain planned.

## Script Contract

All provision scripts are expected to:

- Run in native on-prem shells:
  - ADUC: Windows PowerShell `5.1`
  - EXOP: Exchange Management Shell (Windows PowerShell `5.1`)
- Validate required CSV headers
- Validate required modules and active management connections
- Provide console status output and required per-run transcript logging
- Support `-WhatIf` for dry runs
- Export per-record results to timestamped CSV
- Use shared helpers from `./TenantShift/Common/OnPrem/OnPrem.Common.psm1`

## Catalog

| ID | Script | Input Template | Workload | Purpose | Depends On |
|---|---|---|---|---|---|
| P-ADUC-0010 | `P-ADUC-0010-Create-ActiveDirectoryOrganizationalUnits.ps1` | `P-ADUC-0010-Create-ActiveDirectoryOrganizationalUnits.input.csv` | ADUC | Create AD OUs for placement/delegation boundaries. | Parent OU pathing |
| P-ADUC-0020 | `P-ADUC-0020-Create-ActiveDirectoryUsers.ps1` | `P-ADUC-0020-Create-ActiveDirectoryUsers.input.csv` | ADUC | Create AD users from CSV. | Target OU exists |
| P-ADUC-0030 | `P-ADUC-0030-Create-ActiveDirectoryContacts.ps1` | `P-ADUC-0030-Create-ActiveDirectoryContacts.input.csv` | ADUC | Create AD contacts from CSV. | Target OU exists |
| P-ADUC-0040 | `P-ADUC-0040-Create-ActiveDirectorySecurityGroups.ps1` | `P-ADUC-0040-Create-ActiveDirectorySecurityGroups.input.csv` | ADUC | Create AD security groups. | Target OU exists |
| P-ADUC-0050 | `P-ADUC-0050-Create-ActiveDirectoryDistributionGroups.ps1` | `P-ADUC-0050-Create-ActiveDirectoryDistributionGroups.input.csv` | ADUC | Create AD distribution groups. | Target OU exists |
| P-GPOL-0010 | `P-GPOL-0010-Import-GroupPolicyBackups.ps1` | `P-GPOL-0010-Import-GroupPolicyBackups.input.csv` | GPOL | Import/create GPOs from backup path definitions. | Backup folder accessibility and GPMC availability | **Planned** |
| P-EXOP-0010 | `P-EXOP-0010-Create-ExchangeOnPremMailContacts.ps1` | `P-EXOP-0010-Create-ExchangeOnPremMailContacts.input.csv` | EXOP | Create Exchange on-prem mail contacts. | AD object pathing and Exchange management connectivity |
| P-EXOP-0020 | `P-EXOP-0020-Create-ExchangeOnPremDistributionLists.ps1` | `P-EXOP-0020-Create-ExchangeOnPremDistributionLists.input.csv` | EXOP | Create Exchange on-prem distribution groups. | Exchange management connectivity |
| P-EXOP-0030 | `P-EXOP-0030-Create-ExchangeOnPremMailEnabledSecurityGroups.ps1` | `P-EXOP-0030-Create-ExchangeOnPremMailEnabledSecurityGroups.input.csv` | EXOP | Create Exchange on-prem mail-enabled security groups. | AD groups and Exchange connectivity |
| P-EXOP-0040 | `P-EXOP-0040-Create-ExchangeOnPremDynamicDistributionGroups.ps1` | `P-EXOP-0040-Create-ExchangeOnPremDynamicDistributionGroups.input.csv` | EXOP | Create Exchange on-prem dynamic distribution groups. | Recipient filter validation |
| P-EXOP-0050 | `P-EXOP-0050-Create-ExchangeOnPremSharedMailboxes.ps1` | `P-EXOP-0050-Create-ExchangeOnPremSharedMailboxes.input.csv` | EXOP | Create Exchange on-prem shared mailboxes. | Exchange management connectivity |
| P-EXOP-0060 | `P-EXOP-0060-Create-ExchangeOnPremResourceMailboxes.ps1` | `P-EXOP-0060-Create-ExchangeOnPremResourceMailboxes.input.csv` | EXOP | Create Exchange on-prem room/equipment mailboxes. | Exchange management connectivity |
| P-FILE-0010 | `P-FILE-0010-Create-FileServicesShares.ps1` | `P-FILE-0010-Create-FileServicesShares.input.csv` | FILE | Create file shares from CSV definitions. | File server pathing and naming standards | **Planned** |
| P-FILE-0020 | `P-FILE-0020-Set-FileServicesSharePermissions.ps1` | `P-FILE-0020-Set-FileServicesSharePermissions.input.csv` | FILE | Apply share-level ACL baseline. | Shares and AD groups exist | **Planned** |
| P-FILE-0030 | `P-FILE-0030-Set-FileServicesNtfsPermissions.ps1` | `P-FILE-0030-Set-FileServicesNtfsPermissions.input.csv` | FILE | Apply NTFS ACL baseline. | Folder paths and AD groups exist | **Planned** |
| P-FILE-0040 | `P-FILE-0040-Create-FileServicesHomeDrives.ps1` | `P-FILE-0040-Create-FileServicesHomeDrives.input.csv` | FILE | Create home drive folders and share mappings. | Users and target file server paths exist | **Planned** |

## Implementation Status Snapshot

- Implemented: `P-ADUC-0010`, `P-ADUC-0020`, `P-ADUC-0030`, `P-ADUC-0040`, `P-ADUC-0050`, `P-EXOP-0010`, `P-EXOP-0020`, `P-EXOP-0030`, `P-EXOP-0040`, `P-EXOP-0050`, `P-EXOP-0060`
- Planned: `P-GPOL-0010`, `P-FILE-0010`, `P-FILE-0020`, `P-FILE-0030`, `P-FILE-0040`

## Recommended Run Sequence

1. ADUC baseline: `P-ADUC-0010`, `P-ADUC-0020`, `P-ADUC-0030`, `P-ADUC-0040`, `P-ADUC-0050`
2. GPOL baseline: `P-GPOL-0010`
3. EXOP baseline: `P-EXOP-0010`, `P-EXOP-0020`, `P-EXOP-0030`, `P-EXOP-0040`, `P-EXOP-0050`, `P-EXOP-0060`
4. FILE baseline: `P-FILE-0010`, `P-FILE-0020`, `P-FILE-0030`, `P-FILE-0040`

## Input and Behavior Notes

- Provision scripts require `-InputCsvPath` and strict header validation.
- Multi-value fields use semicolon-delimited values.
- EXOP scripts are intended to run in an Exchange Management Shell-capable host/session model.
- FILE remains a planned scope and may be refined before implementation.

## Related Docs

- [OnPrem Provision README](./README.md)
- [OnPrem README](../README.md)
- [Root README](../../../README.md)
