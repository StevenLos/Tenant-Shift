# BEHAVIOR-CHANGES.md

Behavioral difference audit for the 8 shared functions being extracted from
`M365.Common.psm1` and `OnPrem.Common.psm1` into `Shared.Common.ps1`.

Produced as part of section-01 pre-refactoring setup.
Audit date: 2026-03-28.

---

## Summary

All 8 candidate shared functions are **byte-for-byte identical** between
`Common/Online/M365.Common.psm1` and `Common/OnPrem/OnPrem.Common.psm1`.
No behavioral normalization decisions are required for the promoted functions.

Three functions that exist in both modules are **intentionally diverged** and
are excluded from the shared module. Those are documented below for completeness.

---

## Audit: Shared Functions (8 total)

### 1. Write-Status

**Classification: Promote — no differences**

Both implementations are identical. Accepts `-Message` (string, mandatory) and
`-Level` (ValidateSet 'INFO','WARN','ERROR','SUCCESS', default 'INFO'). Writes
a `[yyyy-MM-dd HH:mm:ss] [LEVEL] Message` line to the host with a color code
per level. No environment-specific calls.

**Circular dependency check:** Calls only `Write-Host` (built-in). Clean.

---

### 2. Start-RunTranscript

**Classification: Promote — no differences**

Both implementations are identical. Accepts `-OutputCsvPath` (mandatory string)
and `-ScriptPath` (AllowNull string). Derives the transcript directory from the
output CSV path; falls back to the script path directory if the CSV path has no
parent. Creates the directory if it does not exist. Names the transcript file
`Transcript_{ScriptName}_{yyyyMMdd-HHmmss}.log`. Returns the transcript path.
Throws on all failures.

**Circular dependency check:** Calls `Write-Status` (shared), `Split-Path`,
`Test-Path`, `New-Item`, `Start-Transcript`, `Get-Date` (all built-ins). Clean.

---

### 3. Stop-RunTranscript

**Classification: Promote — no differences**

Both implementations are identical. Calls `Stop-Transcript -ErrorAction Stop`.
Silently suppresses the 'not currently transcribing' error; re-throws all
others. No parameters.

**Circular dependency check:** Calls only `Stop-Transcript` (built-in). Clean.

---

### 4. ConvertTo-Bool

**Classification: Promote — no differences**

Both implementations are identical. Accepts `-Value` (AllowNull/AllowEmptyString
object) and `-Default` (bool, default `$false`). Returns `$Default` for null
or whitespace input. Accepts (case-insensitive): `1/true/t/yes/y` → `$true`;
`0/false/f/no/n` → `$false`. Throws for any other value.

**Circular dependency check:** No external function calls. Clean.

---

### 5. ConvertTo-Array

**Classification: Promote — no differences**

Both implementations are identical. Accepts `-Value` (AllowNull/AllowEmptyString
string) and `-Delimiter` (string, default `;`). Returns an empty `[string[]]`
for null or whitespace input. Splits on the delimiter (regex-escaped), trims
each part, and drops whitespace-only parts. Returns `[string[]]`.

**Circular dependency check:** No external function calls (uses
`[System.Collections.Generic.List[string]]` internally). Clean.

---

### 6. Import-ValidatedCsv

**Classification: Promote — no differences**

Both implementations are identical. Accepts `-InputCsvPath` (mandatory string)
and `-RequiredHeaders` (mandatory string[]). Verifies file existence, reads the
header row only for validation (strips BOM and quotes), checks for duplicate
headers, checks for missing required headers, then loads all rows via
`Import-Csv`. Returns `@($rows)`. Throws descriptive messages on all failures.

**Circular dependency check:** Calls `Test-Path`, `Get-Content`, `Import-Csv`
(all built-ins). Clean.

---

### 7. New-ResultObject

**Classification: Promote — no differences**

Both implementations are identical. Parameters: `-RowNumber` (int, mandatory),
`-PrimaryKey` (string, mandatory), `-Action` (string, mandatory), `-Status`
(string, mandatory), `-Message` (string, mandatory). Returns a `[PSCustomObject]`
with exactly six fields in this order: `TimestampUtc`, `RowNumber`, `PrimaryKey`,
`Action`, `Status`, `Message`. `TimestampUtc` uses
`(Get-Date).ToUniversalTime().ToString('o')` — ISO 8601 round-trip format.

**Schema baseline:** `TimestampUtc | RowNumber | PrimaryKey | Action | Status | Message`

**Circular dependency check:** Calls only `Get-Date` (built-in). Clean.

---

### 8. Export-ResultsCsv

**Classification: Promote — no differences**

Both implementations are identical. Accepts `-Results` (object[], mandatory) and
`-OutputCsvPath` (string, mandatory). Creates the output directory if needed.
Pipes results to `Export-Csv -LiteralPath $OutputCsvPath -NoTypeInformation -Encoding UTF8`.
Writes a SUCCESS status message on completion.

**BOM note:** `Export-Csv -Encoding UTF8` behavior differs by PowerShell version:
- Windows PowerShell 5.1: writes UTF-8 **with** BOM
- PowerShell 7.x: writes UTF-8 **without** BOM (UTF8NoBOM is the default)

This difference is pre-existing and intentional — both behaviors are acceptable
for the platform's consumers. The shared module preserves this behavior
unchanged; it does not alter the `-Encoding` parameter.

**Circular dependency check:** Calls `Split-Path`, `Test-Path`, `New-Item`,
`Export-Csv` (built-ins) and `Write-Status` (shared). Clean.

---

## Intentionally Diverged Functions (NOT promoted — stay in environment modules)

### Assert-ModuleCurrent

**Classification: Preserve — divergence is intentional**

**Online (M365.Common.psm1):** Always requires PSGallery access. Throws if the
installed module is outdated. Throws if PSGallery lookup fails.

**OnPrem (OnPrem.Common.psm1):** Offline-safe. Has `-FailOnOutdated` switch
(default: warns only, does not throw). Has `-FailOnGalleryLookupError` switch
(default: warns only, continues with installed version if PSGallery is
unreachable). This design reflects that OnPrem environments may not have
internet access.

**Decision:** Both versions stay in their respective environment modules.
A shared version cannot satisfy both contracts without losing the offline-safety
guarantee of the OnPrem version.

---

### Test-IsTransientException

**Classification: Preserve — divergence is intentional**

**Online (M365.Common.psm1):** Uses `Get-HttpStatusCodeFromException` (an
Online-only helper) to detect HTTP 429 and 5xx status codes. Message patterns
include: `temporar|timeout|timed out|service unavailable|rate limit|try again|gateway|429|500|502|503|504`.

**OnPrem (OnPrem.Common.psm1):** No HTTP status code inspection (AD responses
are not HTTP). Message patterns include: `temporar|timeout|timed out|service
unavailable|rate limit|try again|gateway|429|500|502|503|504|server is not
operational`. The `server is not operational` pattern is AD-specific.

**Decision:** Both versions stay in their respective environment modules.
The transient failure patterns are fundamentally different between HTTP-based
Graph/EXO services and LDAP-based AD connectivity.

---

### Get-RetryDelaySeconds

**Classification: Preserve — divergence is intentional**

**Online (M365.Common.psm1):** Inspects the `Retry-After` HTTP response header
first. If the header is present and numeric, returns that value (capped at
`$MaxDelaySeconds`). Falls back to exponential backoff with jitter if no header.

**OnPrem (OnPrem.Common.psm1):** Exponential backoff with jitter only. No HTTP
header inspection (AD/Exchange on-prem responses do not carry `Retry-After`).

**Decision:** Both versions stay in their respective environment modules.

---

## Import Mechanism Audit Results

**Finding:** All 154 SharedModule scripts load their environment module via
`Import-Module`, **not** dot-sourcing.

Actual pattern (verified across all scripts):
```powershell
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking
```

The plan's Track 1 approach is unaffected by this finding. The "dot-source
chain" in the plan refers to how the .psm1 environment module files will
internally load `Shared.Common.ps1` (via `. "$PSScriptRoot\..\Shared\Shared.Common.ps1"`).
Individual scripts do not need to change their loading mechanism — they continue
to call `Import-Module` on the environment modules as they do today.

The section-01 plan document assumed scripts used dot-sourcing; this is a
documentation error in the plan. The implementation approach is correct and
requires no change.

**$PSScriptRoot usage:** All 154 import paths use `$PSScriptRoot`-relative
references via the `Split-Path -Parent` pattern. No absolute paths found.

---

## Script Signing Audit Results

**Finding:** No scripts in SharedModule use signed execution. No `# SIG #`
block found in any `.ps1` or `.psm1` file. The execution policy does not
require signed scripts.

No re-signing is required for any deliverable in this project.

---

## Graph SDK v2 Compliance Results

**Finding:** Fully compliant. `Ensure-GraphConnection` in `M365.Common.psm1`
uses `Connect-MgGraph -Scopes $RequiredScopes -NoWelcome`. No `-AccessToken`
parameter usage found anywhere in the codebase. No migration action required.

---

## Circular Dependency Audit Results

**Finding:** Clean. All 8 shared functions call only:
- Built-in cmdlets (`Write-Host`, `Split-Path`, `Test-Path`, `New-Item`,
  `Start-Transcript`, `Stop-Transcript`, `Get-Content`, `Import-Csv`,
  `Export-Csv`, `Get-Date`)
- Other functions within the 8 shared set (`Write-Status` is called by
  `Start-RunTranscript` and `Export-ResultsCsv`)

No shared function calls any environment-specific function
(`Assert-ModuleCurrent`, `Ensure-GraphConnection`, `Ensure-ActiveDirectoryConnection`,
etc.). The 8 functions are safe to promote without modification.
