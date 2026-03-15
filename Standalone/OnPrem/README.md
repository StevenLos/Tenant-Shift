# Standalone OnPrem Script Set

Self-contained on-prem scripts for one-off execution without `SharedModule/Common/*` dependencies.

## Shell Requirement

- Windows PowerShell 5.1
- ActiveDirectory module available (RSAT/domain tooling)

## Script Catalog

- Active Directory workload (`00xx`) is implemented for `Provision`, `Modify`, and `InventoryAndReport`.
- Exchange on-prem workload (`02xx`) is implemented for `Provision`, `Modify`, and `InventoryAndReport`.
- Current ID coverage:
  - AD: `P0001,P0002,P0005,P0006,P0009`, `M0001,M0002,M0005-M0010`, `IR0001,IR0002,IR0005-IR0012`
  - Exchange on-prem: `P0213-P0216,P0218,P0219`, `M0213-M0226`, `IR0213-IR0226`

Matching input templates:

- Every `SA-*.ps1` operation script has a same-name `SA-*.input.csv` template in the same folder.
- Inventory templates that are scope-based in shared-module are mirrored here as script-specific `SA-IR*.input.csv` templates.

## Output Location

By default, results/transcripts are written to:

- `Standalone/Standalone_OutputCsvPath/`

## Example Commands

Run from repository root.

```powershell
# Users - provision
powershell -ExecutionPolicy Bypass -File ./Standalone/OnPrem/Provision/SA-P0001-Create-ActiveDirectoryUsers.ps1 \
  -InputCsvPath ./Standalone/OnPrem/Provision/SA-P0001-Create-ActiveDirectoryUsers.input.csv -WhatIf

# Users - modify
powershell -ExecutionPolicy Bypass -File ./Standalone/OnPrem/Modify/SA-M0001-Update-ActiveDirectoryUsers.ps1 \
  -InputCsvPath ./Standalone/OnPrem/Modify/SA-M0001-Update-ActiveDirectoryUsers.input.csv -WhatIf

# Users - inventory from CSV scope
powershell -ExecutionPolicy Bypass -File ./Standalone/OnPrem/InventoryAndReport/SA-IR0001-Get-ActiveDirectoryUsers.ps1 \
  -InputCsvPath ./Standalone/OnPrem/InventoryAndReport/SA-IR0001-Get-ActiveDirectoryUsers.input.csv

# Users - inventory full directory scope
powershell -ExecutionPolicy Bypass -File ./Standalone/OnPrem/InventoryAndReport/SA-IR0001-Get-ActiveDirectoryUsers.ps1 \
  -DiscoverAll -SearchBase "OU=Users,OU=Corp,DC=contoso,DC=com" -MaxObjects 0

# Security groups - provision
powershell -ExecutionPolicy Bypass -File ./Standalone/OnPrem/Provision/SA-P0005-Create-ActiveDirectorySecurityGroups.ps1 \
  -InputCsvPath ./Standalone/OnPrem/Provision/SA-P0005-Create-ActiveDirectorySecurityGroups.input.csv -WhatIf

# Security groups - modify
powershell -ExecutionPolicy Bypass -File ./Standalone/OnPrem/Modify/SA-M0005-Update-ActiveDirectorySecurityGroups.ps1 \
  -InputCsvPath ./Standalone/OnPrem/Modify/SA-M0005-Update-ActiveDirectorySecurityGroups.input.csv -WhatIf

# Security groups - inventory from CSV scope
powershell -ExecutionPolicy Bypass -File ./Standalone/OnPrem/InventoryAndReport/SA-IR0005-Get-ActiveDirectorySecurityGroups.ps1 \
  -InputCsvPath ./Standalone/OnPrem/InventoryAndReport/SA-IR0005-Get-ActiveDirectorySecurityGroups.input.csv

# Security groups - inventory full directory scope
powershell -ExecutionPolicy Bypass -File ./Standalone/OnPrem/InventoryAndReport/SA-IR0005-Get-ActiveDirectorySecurityGroups.ps1 \
  -DiscoverAll -SearchBase "OU=Groups,OU=Corp,DC=contoso,DC=com" -MaxObjects 0
```
