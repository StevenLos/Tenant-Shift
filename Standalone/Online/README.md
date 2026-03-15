# Standalone Online Script Set

Self-contained online scripts for one-off execution without `SharedModule/Common/*` dependencies.

## Shell Requirement

- PowerShell 7.0+
- Required online modules vary by script (`Microsoft.Graph.*`, Exchange Online management modules, SharePoint/Teams modules as applicable)

## Script Catalog

- Entra workload (`30xx`) is implemented for `Provision`, `Modify`, and `InventoryAndReport`.
- Exchange Online workload (`31xx`) is implemented for `Provision`, `Modify`, and `InventoryAndReport`.
- Current ID coverage:
  - Entra: `P3001,P3002,P3005,P3006,P3008`, `M3001-M3009`, `IR3001-IR3008`
  - Exchange Online: `P3113-P3116,P3118,P3119,P3124`, `M3113-M3131`, `IR3113-IR3130`

Matching input templates:

- Every `SA-*.ps1` operation script has a same-name `SA-*.input.csv` template in the same folder.
- Inventory templates that are scope-based in shared-module are mirrored here as script-specific `SA-IR*.input.csv` templates.

## Output Location

By default, results/transcripts are written to:

- `Standalone/Standalone_OutputCsvPath/`
