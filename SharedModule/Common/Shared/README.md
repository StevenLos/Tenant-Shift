# Common Shared

`Common/Shared` contains the shared utility layer used by both Online and OnPrem environment modules.

## Contents

| File | Description |
|------|-------------|
| `Shared.Common.ps1` | Eight utility functions dot-sourced by both environment modules |
| `Shared.Common.psd1` | Module manifest for Shared.Common |

## Shared.Common.ps1 — Utility Functions

| Function | Purpose |
|----------|---------|
| `Write-Status` | Console status output with severity-aware formatting |
| `Start-RunTranscript` | Begin a timestamped transcript log file |
| `Stop-RunTranscript` | Close the active transcript log |
| `ConvertTo-Bool` | Normalize string/int values to `[bool]` |
| `ConvertTo-Array` | Ensure a value is always returned as an array |
| `Import-ValidatedCsv` | Import a CSV and validate required headers |
| `New-ResultObject` | Create a standard per-record result object |
| `Export-ResultsCsv` | Write the result collection to a timestamped CSV |
| `Get-TrimmedValue` | Null-safe string trim — returns `''` for null, trims whitespace otherwise |
| `Convert-MultiValueToString` | Normalize a string or collection to a sorted, deduplicated semicolon-delimited string |
| `Convert-ToOrderedReportObject` | Reorder a PSCustomObject's properties per a specified column order, appending extras |

## Loading Pattern

Both `M365.Common.psm1` and `OnPrem.Common.psm1` dot-source `Shared.Common.ps1` at module load time using a `Get-Command Write-Status` sentinel guard to prevent double-loading:

```powershell
if (-not (Get-Command Write-Status -ErrorAction SilentlyContinue)) {
    . "$PSScriptRoot\..\Shared\Shared.Common.ps1"
}
```

## Circular Dependency Rule

`Shared.Common.ps1` must never call functions from `M365.Common.psm1` or `OnPrem.Common.psm1`. The dependency is strictly one-way: environment modules depend on Shared, never the reverse.

See `CONTRIBUTING.md` for the full developer guide.
