# Provision Detailed Catalog

Detailed catalog of provisioning scripts in `OnPrem/Provision/`.

Operational label: **Provision**.

Current implementation status: partial. ActiveDirectory provision baseline scripts (`P0001`, `P0002`, `P0005`, `P0006`, `P0009`) and ExchangeOnPrem provision scripts (`P0213`, `P0214`, `P0215`, `P0216`, `P0218`, `P0219`) are implemented; remaining scripts are planned.

## Script Contract

All planned build scripts are expected to:

- Run in native on-prem shells:
  - ActiveDirectory (`00xx`): Windows PowerShell `5.1`
  - ExchangeOnPrem (`02xx`): Exchange Management Shell (Windows PowerShell `5.1`)
- Validate required CSV headers
- Validate required modules and active management connections
- Provide console status output and required per-run transcript logging
- Support `-WhatIf` for dry runs
- Export per-record results to timestamped CSV
- Use shared helpers from `./Common/OnPrem/OnPrem.Common.psm1`

## ID Ranges

- `00xx`: ActiveDirectory
- `01xx`: GroupPolicy
- `02xx`: ExchangeOnPrem
- `03xx`: FileServices

## Catalog

| ID | Script | Input Template | Workload | Purpose | Depends On |
|---|---|---|---|---|---|
| P0001 | `P0001-Create-ActiveDirectoryUsers.ps1` | `P0001-Create-ActiveDirectoryUsers.input.csv` | ActiveDirectory | Create AD users from CSV. | Target OU exists |
| P0002 | `P0002-Create-ActiveDirectoryContacts.ps1` | `P0002-Create-ActiveDirectoryContacts.input.csv` | ActiveDirectory | Create AD contacts from CSV. | Target OU exists |
| P0005 | `P0005-Create-ActiveDirectorySecurityGroups.ps1` | `P0005-Create-ActiveDirectorySecurityGroups.input.csv` | ActiveDirectory | Create AD security groups. | Target OU exists |
| P0006 | `P0006-Create-ActiveDirectoryDistributionGroups.ps1` | `P0006-Create-ActiveDirectoryDistributionGroups.input.csv` | ActiveDirectory | Create AD distribution groups. | Target OU exists |
| P0009 | `P0009-Create-ActiveDirectoryOrganizationalUnits.ps1` | `P0009-Create-ActiveDirectoryOrganizationalUnits.input.csv` | ActiveDirectory | Create AD OUs for placement/delegation boundaries. | Parent OU pathing |
| P0101 | `P0101-Import-GroupPolicyBackups.ps1` | `P0101-Import-GroupPolicyBackups.input.csv` | GroupPolicy | Import/create GPOs from backup path definitions. | Backup folder accessibility and GPMC availability |
| P0213 | `P0213-Create-ExchangeOnPremMailContacts.ps1` | `P0213-Create-ExchangeOnPremMailContacts.input.csv` | ExchangeOnPrem | Create Exchange on-prem mail contacts. | AD object pathing and Exchange management connectivity |
| P0214 | `P0214-Create-ExchangeOnPremDistributionLists.ps1` | `P0214-Create-ExchangeOnPremDistributionLists.input.csv` | ExchangeOnPrem | Create Exchange on-prem distribution groups. | Exchange management connectivity |
| P0215 | `P0215-Create-ExchangeOnPremMailEnabledSecurityGroups.ps1` | `P0215-Create-ExchangeOnPremMailEnabledSecurityGroups.input.csv` | ExchangeOnPrem | Create Exchange on-prem mail-enabled security groups. | AD groups and Exchange connectivity |
| P0216 | `P0216-Create-ExchangeOnPremSharedMailboxes.ps1` | `P0216-Create-ExchangeOnPremSharedMailboxes.input.csv` | ExchangeOnPrem | Create Exchange on-prem shared mailboxes. | Exchange management connectivity |
| P0218 | `P0218-Create-ExchangeOnPremResourceMailboxes.ps1` | `P0218-Create-ExchangeOnPremResourceMailboxes.input.csv` | ExchangeOnPrem | Create Exchange on-prem room/equipment mailboxes. | Exchange management connectivity |
| P0219 | `P0219-Create-ExchangeOnPremDynamicDistributionGroups.ps1` | `P0219-Create-ExchangeOnPremDynamicDistributionGroups.input.csv` | ExchangeOnPrem | Create Exchange on-prem dynamic distribution groups. | Recipient filter validation |
| P0301 | `P0301-Create-FileServicesShares.ps1` | `P0301-Create-FileServicesShares.input.csv` | FileServices | Create file shares from CSV definitions. | File server pathing and naming standards |
| P0302 | `P0302-Set-FileServicesSharePermissions.ps1` | `P0302-Set-FileServicesSharePermissions.input.csv` | FileServices | Apply share-level ACL baseline. | Shares and AD groups exist |
| P0303 | `P0303-Set-FileServicesNtfsPermissions.ps1` | `P0303-Set-FileServicesNtfsPermissions.input.csv` | FileServices | Apply NTFS ACL baseline. | Folder paths and AD groups exist |
| P0304 | `P0304-Create-FileServicesHomeDrives.ps1` | `P0304-Create-FileServicesHomeDrives.input.csv` | FileServices | Create home drive folders and share mappings. | Users and target file server paths exist |

## Implementation Status Snapshot

- Implemented: `P0001-Create-ActiveDirectoryUsers.ps1`, `P0002-Create-ActiveDirectoryContacts.ps1`, `P0005-Create-ActiveDirectorySecurityGroups.ps1`, `P0006-Create-ActiveDirectoryDistributionGroups.ps1`, `P0009-Create-ActiveDirectoryOrganizationalUnits.ps1`, `P0213-Create-ExchangeOnPremMailContacts.ps1`, `P0214-Create-ExchangeOnPremDistributionLists.ps1`, `P0215-Create-ExchangeOnPremMailEnabledSecurityGroups.ps1`, `P0216-Create-ExchangeOnPremSharedMailboxes.ps1`, `P0218-Create-ExchangeOnPremResourceMailboxes.ps1`, `P0219-Create-ExchangeOnPremDynamicDistributionGroups.ps1`
- Planned: `P0101`, `P0301`, `P0302`, `P0303`, `P0304`

## Recommended Run Sequence

1. ActiveDirectory baseline: `P0001`, `P0002`, `P0005`, `P0006`, `P0009`
2. GroupPolicy baseline: `P0101`
3. ExchangeOnPrem baseline: `P0213`, `P0214`, `P0215`, `P0216`, `P0218`, `P0219`
4. FileServices baseline: `P0301`, `P0302`, `P0303`, `P0304`

## Input and Behavior Notes

- Planned scripts will require `-InputCsvPath` and strict header validation.
- Multi-value fields are expected to use semicolon-delimited values.
- ExchangeOnPrem scripts are intended to run in an Exchange Management Shell-capable host/session model.
- FileServices (`03xx`) remains a planned scope and may be refined before implementation.

## Related Docs

- [OnPrem Provision README](./README.md)
- [OnPrem README](../README.md)
- [Root README](../../README.md)
