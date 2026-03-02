# CarveOutToNewCo

CSV-driven PowerShell automation for Microsoft 365 and hybrid administration, organized first by environment and then by lifecycle operation.

## Repository Structure

| Folder | Purpose |
|---|---|
| `Online/` | Cloud workloads (Entra, Exchange Online, SharePoint/OneDrive, Teams) split by operation (`Provision`, `Modify`, `InventoryAndReport`) |
| `OnPrem/` | On-prem workloads planning area (ActiveDirectory, GroupPolicy, ExchangeOnPrem, FileServices) split by operation (`Provision`, `Modify`, `InventoryAndReport`) |
| `Common/` | Shared modules split by environment (`Common/Online`, `Common/OnPrem`) plus shared helpers (`Common/Shared`) |
| `Build/` | Build/packaging helper scripts (currently orchestrator workbook generation) |

## Execution Flow

1. Choose environment first: `Online` or `OnPrem`.
2. Choose operation folder: `Provision`, `Modify`, or `InventoryAndReport`.
3. Run the target script with mandatory CSV input (`-InputCsvPath`).

## Naming Standard

- Script: `<Prefix><WW><NN>-<Action>-<Target>.ps1`
- Input template: `<Prefix><WW><NN>-<Action>-<Target>.input.csv` (for script-specific inputs)
- Shared inventory scope input: `Online/InventoryAndReport/Scope-<Domain>.input.csv` (preferred for reusable IR key scopes)
- Results output filename pattern: `Results_<Prefix><WW><NN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Transcript log filename pattern: `Transcript_<Prefix><WW><NN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`

Workload code allocation (`WW` in `<Prefix><WW><NN>`):
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
pwsh ./Online/Provision/P3001-Create-EntraUsers.ps1 -InputCsvPath ./Online/Provision/P3001-Create-EntraUsers.input.csv -WhatIf
pwsh ./Online/Provision/P3001-Create-EntraUsers.ps1 -InputCsvPath ./Online/Provision/P3001-Create-EntraUsers.input.csv
```

Online orchestrators are available for copy/paste command construction:

- `Online/Provision/Provision-Orchestrator.xlsx`
- `Online/Modify/Modify-Orchestrator.xlsx`
- `Online/InventoryAndReport/InventoryAndReport-Orchestrator.xlsx`

Regenerate workbooks with:

```powershell
pwsh ./Build/Build-OrchestratorWorkbooks.ps1
```

## Adding New Scripts Checklist

1. Choose environment: `Online` or `OnPrem`.
2. Choose operation folder by intent: `Provision`, `InventoryAndReport`, or `Modify`.
3. Use the next sequence number in the correct workload code bucket (`00xx`-`29xx` OnPrem, `30xx`-`59xx` Online, `90xx`-`99xx` Shared/Hybrid/Reserved).
4. Keep workload explicit in script filenames (`Entra`, `ExchangeOnline`, `ExchangeOnPrem`, `ActiveDirectory`, `FileServices`, etc.).
5. Reuse the environment-specific common module:
   - Online: `./Common/Online/M365.Common.psm1`
   - OnPrem: (to be added as workload implementations begin)
6. Update README/catalog docs and orchestrator definitions where applicable.

## Documentation Map

- [Online Provision README](./Online/Provision/README.md)
- [Online Provision Detailed Catalog](./Online/Provision/README-Provision-Catalog.md)
- [Online Modify README](./Online/Modify/README.md)
- [Online Modify Detailed Catalog](./Online/Modify/README-Modify-Catalog.md)
- [Online InventoryAndReport README](./Online/InventoryAndReport/README.md)
- [Online InventoryAndReport Detailed Catalog](./Online/InventoryAndReport/README-InventoryAndReport-Catalog.md)
- [Common README](./Common/README.md)
- [OnPrem README](./OnPrem/README.md)
- [OnPrem Provision README](./OnPrem/Provision/README.md)
- [OnPrem Provision Detailed Catalog](./OnPrem/Provision/README-Provision-Catalog.md)
- [OnPrem Modify README](./OnPrem/Modify/README.md)
- [OnPrem Modify Detailed Catalog](./OnPrem/Modify/README-Modify-Catalog.md)
- [OnPrem InventoryAndReport README](./OnPrem/InventoryAndReport/README.md)
- [OnPrem InventoryAndReport Detailed Catalog](./OnPrem/InventoryAndReport/README-InventoryAndReport-Catalog.md)
- [Build README](./Build/README.md)

## License

This repository is licensed under the MIT License. See the `LICENSE` file for full details.



