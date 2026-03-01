# EntraID_EXO_Basic

CSV-driven PowerShell automation for Microsoft 365 administration, organized by operation type so intent is clear before execution.

## Repository Structure

| Folder | Prefix | Purpose | Operation Type |
|---|---|---|---|
| `Build/` | `B` | Provision new objects and baseline configuration | Create/initial setup |
| `Discover/` | `D` | Inventory, reporting, and state collection | Read-only |
| `Update/` | `U` | Modify existing objects and configuration | Change existing state |

## Execution Flow

1. Run `Build` scripts for initial provisioning.
2. Run `Discover` scripts to validate and report current state.
3. Run `Update` scripts for controlled post-provisioning changes.

## Naming Standard

- Script: `<Prefix><##>-<Action>-<Target>.ps1`
- Input template: `<Prefix><##>-<Action>-<Target>.input.csv`
- Results output: `Results_<Prefix><##>-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`

Examples:
- `B01-Create-EntraUsers.ps1`
- `D01-Get-EntraUsers.ps1`
- `U01-Update-EntraUsers.ps1`

## Common Script Contract

All scripts should follow this baseline:

- PowerShell 7+
- CSV header validation
- Module presence/current-version checks
- Connection checks and auto-connect when needed
- Per-record error handling
- `-WhatIf` support for non-discovery scripts
- Timestamped results CSV output
- Retry with exponential backoff for transient failures

## Quick Start

Build example:

```powershell
pwsh ./Build/B01-Create-EntraUsers.ps1 -InputCsvPath ./Build/B01-Create-EntraUsers.input.csv -WhatIf
pwsh ./Build/B01-Create-EntraUsers.ps1 -InputCsvPath ./Build/B01-Create-EntraUsers.input.csv
```

## Adding New Scripts Checklist

1. Choose the correct folder by intent (`Build`, `Discover`, or `Update`).
2. Use the correct prefix (`B`, `D`, or `U`) and next sequence number.
3. Create both the script and matching `.input.csv` template.
4. Keep workload explicit in the file name (`Entra`, `Exchange`, `OneDrive`, `SharePoint`, `MicrosoftTeams`).
5. Reuse `M365.Common.psm1` for shared validation, auth, and output behavior.
6. Update the folder README catalog and run examples.

## Documentation Map

- [Build README](./Build/README.md)
- [Build Detailed Catalog](./Build/README-Build-Catalog.md)
- [Discover README](./Discover/README.md)
- [Discover Detailed Catalog](./Discover/README-Discover-Catalog.md)
- [Update README](./Update/README.md)
- [Update Detailed Catalog](./Update/README-Update-Catalog.md)
