# CarveOutToNewCo

CSV-driven PowerShell automation for Microsoft 365 and hybrid administration, organized first by environment and then by lifecycle operation.

## Repository Structure

| Folder | Purpose |
|---|---|
| `SharedModule/Online/` | Cloud workloads (Entra, Exchange Online, SharePoint/OneDrive, Teams) split by operation (`Provision`, `Modify`, `InventoryAndReport`) |
| `SharedModule/OnPrem/` | On-prem workloads planning area (ActiveDirectory, GroupPolicy, ExchangeOnPrem, FileServices) split by operation (`Provision`, `Modify`, `InventoryAndReport`) |
| `SharedModule/Common/` | Shared modules split by environment (`SharedModule/Common/Online`, `SharedModule/Common/OnPrem`) plus shared helpers (`SharedModule/Common/Shared`) |
| `SharedModule/Build/` | Build/packaging helper scripts (currently orchestrator workbook generation) |
| `SharedModule/` | Repository-native shared-module model index (what areas/scripts participate in shared-module architecture) |
| `Imported/` | Dedicated read-only staging ground for externally sourced scripts/code/processes |
| `Imported/IMPORTED-M365-Integration-Scripts/` | Existing imported external script set retained as read-only staging content |
| `Utilities/` | Top-level utility scripts for cross-workload helpers (for example CSV-driven password generation) |

## Execution Flow

1. Choose environment first: `SharedModule/Online` or `SharedModule/OnPrem`.
2. Choose operation folder: `Provision`, `Modify`, or `InventoryAndReport`.
3. Run the target script using its supported scope mode (`-InputCsvPath` for CSV-driven scope; `-DiscoverAll` where implemented for unbounded discovery).

For repository-native shared-module scripts, use `SharedModule/` guidance.
For externally sourced material, use `Imported/` staging and keep imported content read-only.

## Discovery Scope Modes

- Default model for this repository is CSV-bounded execution (`-InputCsvPath`).
- Implemented OnPrem ActiveDirectory and ExchangeOnPrem inventory/report scripts (`IR0001`, `IR0002`, `IR0005`, `IR0006`, `IR0007`, `IR0008`, `IR0009`, `IR0010`, `IR0011`, `IR0012`, `IR0213` through `IR0226`) also support unbounded discovery via `-DiscoverAll`, with script-specific scope controls (for example `-SearchBase`/`-Server`/`-MaxObjects` for directory objects and `-LogPath`/`-LookbackDays`/`-MaxObjects` for RPC log inventory).
- Online inventory/report scripts support the same dual-scope model (`-InputCsvPath` or `-DiscoverAll`) with script-specific scope controls where required by workload cmdlets.

## OnPrem Shell Baseline

- ActiveDirectory scripts (`OnPrem`, `00xx`) target native Windows PowerShell `5.1`.
- ExchangeOnPrem scripts (`OnPrem`, `02xx`) target native Exchange Management Shell (Windows PowerShell `5.1`).

## Naming Standard

- Shared-module script: `SM-<P|M|IR><WWNN>-<Action>-<Target>.ps1`
- Shared-module input template: `SM-<P|M|IR><WWNN>-<Action>-<Target>.input.csv`
- Shared inventory scope input: `SharedModule/Online/InventoryAndReport/Scope-<Domain>.input.csv` (preferred for reusable IR key scopes)
- Shared-module results pattern: `Results_SM-<P|M|IR><WWNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Shared-module transcript pattern: `Transcript_SM-<P|M|IR><WWNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`

Workload code allocation (`WWNN` in `<P|M|IR><WWNN>`):
- `00-29`: OnPrem
- `30-59`: Online
- `60-89`: Unallocated
- `90-99`: Shared/Hybrid/Reserved

Current OnPrem workload codes:
- `00xx`: ActiveDirectory
- `01xx`: GroupPolicy
- `02xx`: ExchangeOnPrem
- `03xx`: FileServices

Current Online workload codes:
- `30xx`: Entra
- `31xx`: Exchange Online
- `32xx`: SharePoint/OneDrive
- `33xx`: Teams

## Common Script Contract

All scripts should follow this baseline:

- PowerShell version declared via `#Requires`
- CSV header validation
- Module presence/current-version checks
- Connection checks and auto-connect when needed
- Per-record error handling
- `-WhatIf` support for non-discovery scripts
- Timestamped results CSV output
- Mandatory per-run transcript logging to the same output folder as `-OutputCsvPath`
- Retry with exponential backoff for transient failures

## Script Header Metadata

PowerShell scripts and modules include a standardized metadata block with:

- `.SCRIPTVERSION`: timestamp-based version in `yyyyMMdd-HHmmss` format (local timezone, no suffix)
- `.POWERSHELLREQUIRED`: required PowerShell version
- `.REQUIREDMODULES`: module names expected by the file
- `.MODULEVERSIONPOLICY`: module version expectation

## Quick Start

Online provision example (run from repository root):

```powershell
pwsh ./SharedModule/Online/Provision/SM-P3001-Create-EntraUsers.ps1 -InputCsvPath ./SharedModule/Online/Provision/SM-P3001-Create-EntraUsers.input.csv -WhatIf
pwsh ./SharedModule/Online/Provision/SM-P3001-Create-EntraUsers.ps1 -InputCsvPath ./SharedModule/Online/Provision/SM-P3001-Create-EntraUsers.input.csv
```

Online orchestrators are available for copy/paste command construction:

- `SharedModule/Online/Provision/Provision-Orchestrator.xlsx`
- `SharedModule/Online/Modify/Modify-Orchestrator.xlsx`
- `SharedModule/Online/InventoryAndReport/InventoryAndReport-Orchestrator.xlsx`

Regenerate workbooks with:

```powershell
pwsh ./SharedModule/Build/Build-OrchestratorWorkbooks.ps1
```

## Validation

Run repository contract validation:

```powershell
pwsh ./SharedModule/Build/Test-RepositoryContracts.ps1
```

Run Pester tests:

```powershell
Invoke-Pester -Path ./SharedModule/Tests
```

## Adding New Scripts Checklist

1. Choose environment: `SharedModule/Online` or `SharedModule/OnPrem`.
2. Choose operation folder by intent: `Provision`, `InventoryAndReport`, or `Modify`.
3. Use the next sequence number in the correct workload code bucket (`00xx`-`29xx` OnPrem, `30xx`-`59xx` Online, `90xx`-`99xx` Shared/Hybrid/Reserved).
4. Keep workload explicit in script filenames (`Entra`, `ExchangeOnline`, `ExchangeOnPrem`, `ActiveDirectory`, `FileServices`, etc.).
5. Reuse the environment-specific common module:
   - Online: `./SharedModule/Common/Online/M365.Common.psm1`
   - OnPrem: `./SharedModule/Common/OnPrem/OnPrem.Common.psm1`
6. Update README/catalog docs and orchestrator definitions where applicable.

## Documentation Map

- [Online Provision README](./SharedModule/Online/Provision/README.md)
- [Online Provision Detailed Catalog](./SharedModule/Online/Provision/README-Provision-Catalog.md)
- [Online Modify README](./SharedModule/Online/Modify/README.md)
- [Online Modify Detailed Catalog](./SharedModule/Online/Modify/README-Modify-Catalog.md)
- [Online InventoryAndReport README](./SharedModule/Online/InventoryAndReport/README.md)
- [Online InventoryAndReport Detailed Catalog](./SharedModule/Online/InventoryAndReport/README-InventoryAndReport-Catalog.md)
- [Common README](./SharedModule/Common/README.md)
- [OnPrem README](./SharedModule/OnPrem/README.md)
- [OnPrem Provision README](./SharedModule/OnPrem/Provision/README.md)
- [OnPrem Provision Detailed Catalog](./SharedModule/OnPrem/Provision/README-Provision-Catalog.md)
- [OnPrem Modify README](./SharedModule/OnPrem/Modify/README.md)
- [OnPrem Modify Detailed Catalog](./SharedModule/OnPrem/Modify/README-Modify-Catalog.md)
- [OnPrem InventoryAndReport README](./SharedModule/OnPrem/InventoryAndReport/README.md)
- [OnPrem InventoryAndReport Detailed Catalog](./SharedModule/OnPrem/InventoryAndReport/README-InventoryAndReport-Catalog.md)
- [Build README](./SharedModule/Build/README.md)
- [SharedModule README](./SharedModule/README.md)
- [Imported README](./Imported/README.md)
- [Utilities README](./Utilities/README.md)

## License

This repository is licensed under the MIT License. See the `LICENSE` file for full details.
