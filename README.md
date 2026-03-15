# CarveOutToNewCo Production

Production package for running migration and reporting PowerShell scripts.

## Included Content

- `SharedModule/`: standard script sets with shared runtime modules
- `Standalone/`: self-contained scripts for one-off execution
- `Utilities/`: helper utilities
- `LICENSE`

## What This Package Is For

Use this package to:

- run Provision (`P`), Modify (`M`), and InventoryAndReport (`IR`) scripts
- use included `.input.csv` files as templates for your own runs
- generate results/transcripts in each operation output folder

## Prerequisites

- PowerShell 7+ for online workloads and utilities
- Windows PowerShell 5.1 / Exchange Management Shell for on-prem workloads
- Required workload modules and admin permissions for your target environment

## Quick Start

Run commands from this folder (repository root):

```powershell
# SharedModule online example
pwsh ./SharedModule/Online/Provision/SM-P3001-Create-EntraUsers.ps1 -InputCsvPath ./SharedModule/Online/Provision/SM-P3001-Create-EntraUsers.input.csv -WhatIf

# Standalone online example
pwsh ./Standalone/Online/Provision/SA-P3001-Create-EntraUsers.ps1 -InputCsvPath ./Standalone/Online/Provision/SA-P3001-Create-EntraUsers.input.csv -WhatIf

# Utility example
pwsh ./Utilities/Utility-Espresso/Utility-Espresso.ps1
```

## Output Locations

By default, scripts write results and transcript logs to the matching output folder for that script family (for example `Provision_OutputCsvPath`, `Modify_OutputCsvPath`, `InventoryAndReport_OutputCsvPath`, and `Standalone_OutputCsvPath`).

## Folder Documentation

- [SharedModule README](./SharedModule/README.md)
- [Standalone README](./Standalone/README.md)
- [Utilities README](./Utilities/README.md)
