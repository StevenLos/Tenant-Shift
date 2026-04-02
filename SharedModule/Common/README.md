# Common Folder

`Common` stores shared modules and helper assets used by all SharedModule scripts.

## Structure

| Folder | Contents |
|--------|----------|
| `Common/Shared/` | `Shared.Common.ps1` and `Shared.Common.psd1` — eight utility functions shared across both environments |
| `Common/Online/` | `M365.Common.psm1` — environment module for cloud workload scripts; dot-sources `Shared.Common.ps1` |
| `Common/OnPrem/` | `OnPrem.Common.psm1` — environment module for on-prem workload scripts; dot-sources `Shared.Common.ps1` |

## Shared Layer — Shared.Common.ps1

`Shared.Common.ps1` contains the eleven utility functions common to both environments:

- `Write-Status`
- `Start-RunTranscript`
- `Stop-RunTranscript`
- `ConvertTo-Bool`
- `ConvertTo-Array`
- `Import-ValidatedCsv`
- `New-ResultObject`
- `Export-ResultsCsv`
- `Get-TrimmedValue`
- `Convert-MultiValueToString`
- `Convert-ToOrderedReportObject`

Both environment modules dot-source this file and use a `Get-Command Write-Status` guard to prevent double-loading.

**Circular dependency rule:** `Shared.Common.ps1` must never call functions from `M365.Common.psm1` or `OnPrem.Common.psm1`. The dependency is one-way: environment modules depend on Shared, never the reverse.

See `CONTRIBUTING.md` for the full circular dependency rule and the developer guide.

