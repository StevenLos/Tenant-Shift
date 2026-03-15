# Utilities

`Utilities/` is for top-level helper scripts that are not tied to a single workload folder.

## Purpose

- Host one-off or cross-workload utility scripts.
- Keep utility logic separate from `SharedModule/` and `Standalone/` operation scripts.
- Support CSV-driven utility workflows when repeatable batch input is needed.

## Current Scripts

- `Utility-Generate-Passwords/Utility-Generate-Passwords.ps1`
  - Generates passwords from CSV column value pools.
  - Writes generated passwords to a results CSV.
- `Utility-Espresso/Utility-Espresso.ps1`
  - Sends a periodic key sequence to keep the workstation session active.
  - Reports cycle count and cycle timestamp on every loop iteration.
  - Supports finite duration (`-DurationMinutes`) or run-until-interrupted mode.
- `Utility-Check-Prerequisites/Utility-Check-Prerequisites.ps1`
  - Verifies PowerShell 5.1 and PowerShell 7+ runtime presence.
  - Discovers required modules from repository scripts and checks whether each module is present.
  - Checks PSGallery latest versions (unless `-SkipGalleryCheck`) to flag modules that are out of date.
  - Prints grouped tables for present, missing, up-to-date, and out-of-date modules.
  - Includes install/update command lines for missing or out-of-date modules.

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
pwsh ./Utilities/Utility-Generate-Passwords/Utility-Generate-Passwords.ps1 -InputCsvPath ./Utilities/Utility-Generate-Passwords/Utility-Generate-Passwords.input.csv -PasswordCount 20
```

Password generator with CSV-controlled count:

```powershell
pwsh ./Utilities/Utility-Generate-Passwords/Utility-Generate-Passwords.ps1 -InputCsvPath ./Utilities/Utility-Generate-Passwords/Utility-Generate-Passwords.input.csv
```

Espresso utility (run until interrupted):

```powershell
pwsh ./Utilities/Utility-Espresso/Utility-Espresso.ps1
```

Espresso utility with finite duration:

```powershell
pwsh ./Utilities/Utility-Espresso/Utility-Espresso.ps1 -DurationMinutes 120 -IntervalSeconds 59
```

Prerequisites audit utility:

```powershell
pwsh ./Utilities/Utility-Check-Prerequisites/Utility-Check-Prerequisites.ps1
```

Prerequisites audit utility (skip PSGallery checks/offline-safe):

```powershell
pwsh ./Utilities/Utility-Check-Prerequisites/Utility-Check-Prerequisites.ps1 -SkipGalleryCheck
```

## Security Notes

- Generated passwords are written as plain text to the output CSV by design.
- Treat output files as sensitive secrets and store/transfer them securely.
- Delete or rotate output files when no longer needed.
