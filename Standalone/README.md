# Standalone Folder

`Standalone` contains self-contained scripts for one-off execution that do not depend on `SharedModule/Common/*` modules.

## Purpose

Use `Standalone` when you need a script that can run independently for a specific case.

## Folder Layout

- `Standalone/Online/`: standalone online scripts grouped by `Provision`, `Modify`, and `InventoryAndReport`
- `Standalone/OnPrem/`: standalone on-prem scripts grouped by `Provision`, `Modify`, and `InventoryAndReport`
- `Standalone/Standalone_OutputCsvPath/`: default output/transcript location for standalone runs

## Naming

- Script: `SA-<P|M|IR><WWNN>-<Action>-<Target>.ps1`
- Input file: `SA-<P|M|IR><WWNN>-<Action>-<Target>.input.csv`
- Output file: `Results_SA-<P|M|IR><WWNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Transcript file: `Transcript_SA-<P|M|IR><WWNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`

## Quick Start

From repository root:

```powershell
pwsh ./Standalone/Online/Provision/SA-P3001-Create-EntraUsers.ps1 -InputCsvPath ./Standalone/Online/Provision/SA-P3001-Create-EntraUsers.input.csv -WhatIf
```

## References

- [Standalone Online README](./Online/README.md)
- [Standalone OnPrem README](./OnPrem/README.md)
- [SharedModule README](../SharedModule/README.md)
