# Tenant-Shift

CSV-driven PowerShell automation for Microsoft 365 and hybrid administration, organized first by environment and then by lifecycle operation.

## Repository Structure

| Folder | Purpose |
|---|---|
| `SharedModule/Online/` | Cloud workloads (Entra, Exchange Online, SharePoint/OneDrive, Teams) split by operation (`Provision`, `Modify`, `InventoryAndReport`) |
| `SharedModule/OnPrem/` | On-prem workloads planning area (ActiveDirectory, GroupPolicy, ExchangeOnPrem, FileServices) split by operation (`Provision`, `Modify`, `InventoryAndReport`) |
| `SharedModule/Common/` | Shared utility layer (`Common/Shared/Shared.Common.ps1`) plus environment modules (`Common/Online/M365.Common.psm1`, `Common/OnPrem/OnPrem.Common.psm1`) |
| `SharedModule/Development/Build/` | Build automation, contract tests, quality gate, and inventory scripts |
| `SharedModule/Development/Tests/` | Pester test suite for platform contracts, build utilities, and prerequisite engine |
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
- OnPrem `ADUC` and `EXOP` discovery scripts also support unbounded discovery via `-DiscoverAll`, with script-specific scope controls (for example `-SearchBase`/`-Server`/`-MaxObjects` for directory objects and `-LogPath`/`-LookbackDays`/`-MaxObjects` for RPC log inventory).
- Online discovery scripts support the same dual-scope model (`-InputCsvPath` or `-DiscoverAll`) with script-specific scope controls where required by workload cmdlets.

## OnPrem Shell Baseline

- ActiveDirectory scripts (`ADUC` workload) target native Windows PowerShell `5.1`.
- ExchangeOnPrem scripts (`EXOP` workload) target native Exchange Management Shell (Windows PowerShell `5.1`).

## Naming Standard

- Shared-module script: `<P|M|D>-<WWWW>-<NNNN>-<Action>-<Target>.ps1`
- Shared-module input template: `<P|M|D>-<WWWW>-<NNNN>-<Action>-<Target>.input.csv`
- Shared inventory scope input: `SharedModule/Online/InventoryAndReport/Scope-<Object>.input.csv` (preferred for reusable discovery key scopes)
- Shared-module results pattern: `Results_<P|M|D>-<WWWW>-<NNNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Shared-module transcript pattern: `Transcript_<P|M|D>-<WWWW>-<NNNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`

Operation prefix:
- `P-`: Provision
- `M-`: Modify
- `D-`: Discovery/InventoryAndReport

Current OnPrem workload codes (`<WWWW>`):
- `ADUC`: ActiveDirectory
- `EXOP`: ExchangeOnPrem

Current Online workload codes (`<WWWW>`):
- `MEID`: Entra (Microsoft Entra ID)
- `EXOL`: Exchange Online
- `ONDR`: OneDrive
- `SPOL`: SharePoint
- `TEAM`: Teams

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
pwsh ./SharedModule/Online/Provision/P-MEID-0010-Create-EntraUsers.ps1 -InputCsvPath ./SharedModule/Online/Provision/P-MEID-0010-Create-EntraUsers.input.csv -WhatIf
pwsh ./SharedModule/Online/Provision/P-MEID-0010-Create-EntraUsers.ps1 -InputCsvPath ./SharedModule/Online/Provision/P-MEID-0010-Create-EntraUsers.input.csv
```

Online orchestrators are available for copy/paste command construction:

- `SharedModule/Online/Provision/Provision-Orchestrator.xlsx`
- `SharedModule/Online/Modify/Modify-Orchestrator.xlsx`
- `SharedModule/Online/InventoryAndReport/InventoryAndReport-Orchestrator.xlsx`

Regenerate workbooks with:

```powershell
pwsh ./SharedModule/Development/Build/Build-OrchestratorWorkbooks.ps1
```

## Validation

Run repository contract validation:

```powershell
pwsh ./SharedModule/Development/Build/Test-RepositoryContracts.ps1
```

Run the quality gate before committing:

```powershell
pwsh ./SharedModule/Development/Build/Invoke-QualityGate.ps1 -Path <your-script>
```

Run Pester tests:

```powershell
Invoke-Pester -Path ./SharedModule/Development/Tests
```

## Adding New Scripts Checklist

1. Run `SharedModule/Utilities/Initialize-Platform.ps1 -Profile Contributor` to validate your environment before starting.
2. Choose environment: `SharedModule/Online` or `SharedModule/OnPrem`.
3. Choose operation folder by intent: `Provision`, `InventoryAndReport`, or `Modify`.
4. Use the correct workload code (`MEID`, `EXOL`, `ONDR`, `SPOL`, `TEAM` for Online; `ADUC`, `EXOP` for OnPrem) and the next available sequence number.
5. Keep workload explicit in script filenames (`EntraUsers`, `ExchangeOnlineDistributionLists`, `ActiveDirectoryUsers`, etc.).
6. Reuse the environment-specific common module:
   - Online: `./SharedModule/Common/Online/M365.Common.psm1`
   - OnPrem: `./SharedModule/Common/OnPrem/OnPrem.Common.psm1`
7. Run `SharedModule/Development/Build/Invoke-QualityGate.ps1 -Path <your-script>` before committing.
8. Update README/catalog docs and orchestrator definitions where applicable.

## Documentation Map

- [Online Provision README](./SharedModule/Online/Provision/README.md)
- [Online Provision Detailed Catalog](./SharedModule/Online/Provision/README-Provision-Catalog.md)
- [Online Provision Runbook](./SharedModule/Online/Provision/RUNBOOK-Provision.md)
- [Online Modify README](./SharedModule/Online/Modify/README.md)
- [Online Modify Detailed Catalog](./SharedModule/Online/Modify/README-Modify-Catalog.md)
- [Online Modify Runbook](./SharedModule/Online/Modify/RUNBOOK-Modify.md)
- [Online InventoryAndReport README](./SharedModule/Online/InventoryAndReport/README.md)
- [Online InventoryAndReport Detailed Catalog](./SharedModule/Online/InventoryAndReport/README-InventoryAndReport-Catalog.md)
- [Online InventoryAndReport Runbook](./SharedModule/Online/InventoryAndReport/RUNBOOK-InventoryAndReport.md)
- [Common README](./SharedModule/Common/README.md)
- [OnPrem README](./SharedModule/OnPrem/README.md)
- [OnPrem Provision README](./SharedModule/OnPrem/Provision/README.md)
- [OnPrem Provision Detailed Catalog](./SharedModule/OnPrem/Provision/README-Provision-Catalog.md)
- [OnPrem Provision Runbook](./SharedModule/OnPrem/Provision/RUNBOOK-Provision.md)
- [OnPrem Modify README](./SharedModule/OnPrem/Modify/README.md)
- [OnPrem Modify Detailed Catalog](./SharedModule/OnPrem/Modify/README-Modify-Catalog.md)
- [OnPrem Modify Runbook](./SharedModule/OnPrem/Modify/RUNBOOK-Modify.md)
- [OnPrem InventoryAndReport README](./SharedModule/OnPrem/InventoryAndReport/README.md)
- [OnPrem InventoryAndReport Detailed Catalog](./SharedModule/OnPrem/InventoryAndReport/README-InventoryAndReport-Catalog.md)
- [OnPrem InventoryAndReport Runbook](./SharedModule/OnPrem/InventoryAndReport/RUNBOOK-InventoryAndReport.md)
- [Build README](./SharedModule/Development/Build/README.md)
- [SharedModule README](./SharedModule/README.md)
- [CONTRIBUTING.md](./SharedModule/CONTRIBUTING.md)
- [CHANGELOG.md](./SharedModule/CHANGELOG.md)
- [Imported README](./Imported/README.md)
- [Utilities README](./Utilities/README.md)

## License

This repository is licensed under the MIT License. See the `LICENSE` file for full details.
