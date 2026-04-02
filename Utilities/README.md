# Utilities

`Utilities/` is for top-level helper scripts that are not tied to a single workload folder.

## Purpose

- Host one-off or cross-workload utility scripts.
- Keep utility logic separate from `SharedModule/` operation scripts.
- Support CSV-driven utility workflows when repeatable batch input is needed.

## Current Scripts

- `Utility-Generate-Passwords.ps1`
  - Generates passwords from CSV column value pools.
  - Writes generated passwords to a results CSV.
- `Utility-Espresso.ps1`
  - Sends a periodic key sequence to keep the workstation session active.
  - Reports cycle count and cycle timestamp on every loop iteration.
  - Supports finite duration (`-DurationMinutes`) or run-until-interrupted mode.
- `Utility-Convert-CsvFolderToSingleXlsx.ps1`
  - Scans a folder for CSV files and creates a single XLSX workbook.
  - Writes one worksheet per CSV file.
  - Derives worksheet names from the CSV relative path and preserves CSV text values.
- `Utility-Convert-CsvFolderToIndividualXlsx.ps1`
  - Scans a folder for CSV files and creates one XLSX file per CSV file.
  - Preserves relative subfolders when an alternate output folder is used.
  - Preserves CSV text values.
- `Utility-Check-Prerequisites.ps1`
  - Scans repository PowerShell files for declared module requirements using the shared prerequisite engine.
  - Supports include/exclude path filters, optional profile validation, JSON/table output, and fail-on conditions for automation.
- `Utility-Unblock-Files.ps1`
  - Recursively removes `Zone.Identifier` from files under a target path.
  - Helps prepare downloaded scripts for execution under `RemoteSigned`.

## Password Generator Input Contract

- You can use an optional first column named `PasswordCount` as a control column.
- If `PasswordCount` exists, place one integer value (`1` to `100000`) in that column to control how many passwords are generated.
- All remaining columns are treated as password component pools.
- The script randomly picks one value from each component column and concatenates in header order.
- If `PasswordCount` is not present, the script uses `-PasswordCount` (default `20`).
- If both are provided, explicit `-PasswordCount` parameter takes precedence.
- Example columns:
  - `PasswordCount`,`Category1`,`Category2`,`Category3`,`Category4`,`Symbol1`,`Symbol2`,`Number1`,`Number2`,`Number3`,`Number4`

## Usage Pattern

Password generator (run from repository root):

```powershell
pwsh ./Utilities/Utility-Generate-Passwords.ps1 -InputCsvPath ./Utilities/Utility-Generate-Passwords.input.csv -PasswordCount 20
```

Password generator with CSV-controlled count:

```powershell
pwsh ./Utilities/Utility-Generate-Passwords.ps1 -InputCsvPath ./Utilities/Utility-Generate-Passwords.input.csv
```

Espresso utility (run until interrupted):

```powershell
pwsh ./Utilities/Utility-Espresso.ps1
```

Espresso utility with finite duration:

```powershell
pwsh ./Utilities/Utility-Espresso.ps1 -DurationMinutes 120 -IntervalSeconds 59
```

CSV folder to single XLSX workbook:

```powershell
pwsh ./Utilities/Utility-Convert-CsvFolderToSingleXlsx.ps1 -InputFolderPath ./SharedModule/Online/InventoryAndReport/InventoryAndReport_OutputCsvPath
```

CSV folder to single XLSX workbook with recursion and overwrite:

```powershell
pwsh ./Utilities/Utility-Convert-CsvFolderToSingleXlsx.ps1 -InputFolderPath ./SharedModule/Online/InventoryAndReport -Recurse -OutputXlsxPath ./SharedModule/Online/InventoryAndReport/InventoryWorkbook.xlsx -Overwrite
```

CSV folder to individual XLSX files:

```powershell
pwsh ./Utilities/Utility-Convert-CsvFolderToIndividualXlsx.ps1 -InputFolderPath ./SharedModule/Online/InventoryAndReport/InventoryAndReport_OutputCsvPath
```

CSV folder to individual XLSX files with recursion and separate output folder:

```powershell
pwsh ./Utilities/Utility-Convert-CsvFolderToIndividualXlsx.ps1 -InputFolderPath ./SharedModule/Online/InventoryAndReport -Recurse -OutputFolderPath ./SharedModule/Online/InventoryAndReport/ConvertedXlsx -Overwrite
```

Repository prerequisite scan:

```powershell
pwsh ./Utilities/Utility-Check-Prerequisites.ps1 -RepositoryRoot . -SkipGalleryCheck
```

Repository prerequisite scan with Online operator validation and CI-style fail conditions:

```powershell
pwsh ./Utilities/Utility-Check-Prerequisites.ps1 -RepositoryRoot . -Profile OnlineOperator -FailOn Missing,ProfileFailure
```

Recursively unblock downloaded files under a folder:

```powershell
pwsh ./Utilities/Utility-Unblock-Files.ps1 -TargetPath ./SharedModule
```

## Security Notes

- Generated passwords are written as plain text to the output CSV by design.
- Treat output files as sensitive secrets and store/transfer them securely.
- Delete or rotate output files when no longer needed.
- `Utility-Convert-CsvFolderToSingleXlsx` requires Windows and a locally installed Microsoft Excel desktop application because it uses Excel COM automation.
- `Utility-Convert-CsvFolderToIndividualXlsx` requires Windows and a locally installed Microsoft Excel desktop application because it uses Excel COM automation.
