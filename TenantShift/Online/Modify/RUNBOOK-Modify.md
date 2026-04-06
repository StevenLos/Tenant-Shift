# Operator Runbook — Online Modify Scripts

This runbook covers running `M-WW-NNNN` scripts that update existing Microsoft 365 objects. Modify scripts change attributes of objects that already exist. They support `-WhatIf` for dry-run validation and write a result CSV documenting every row processed.

> **Always run with `-WhatIf` first.** Modify scripts make live changes to existing objects. A dry run shows what would change without applying it.

---

## 1. Prerequisites

### Required PowerShell Version

PowerShell **7.0 or later** is required.

```powershell
$PSVersionTable.PSVersion
```

### Step 1: Run Initialize-TenantShift.ps1

**Run this before anything else.** Use the exact Online operator profile. It validates PowerShell 7+, `Microsoft.Graph.Authentication`, `ExchangeOnlineManagement`, `PnP.PowerShell`, and the `PNP_CLIENT_ID` environment variable.

```powershell
.\Initialize-TenantShift\Initialize-TenantShift.ps1 -Profile OnlineOperator
```

Resolve all failures before proceeding.

### Required Modules

| Module | Minimum version | Purpose |
|--------|----------------|---------|
| `Microsoft.Graph.*` | Latest | Entra / identity updates |
| `ExchangeOnlineManagement` | 3.x | Exchange Online changes |
| `PnP.PowerShell` | Latest | SharePoint Online / OneDrive changes |
| `MicrosoftTeams` | Latest | Teams changes |

### Execution Policy

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Permissions

Modify scripts make changes to existing objects. Ensure your account has write permissions.

| Workload | Required role |
|----------|--------------|
| Entra users / groups (GR) | User Administrator or Groups Administrator |
| Exchange Online (EX) | Exchange Administrator or Exchange Recipient Administrator |
| SharePoint / OneDrive (SP) | SharePoint Administrator |
| Teams (TM) | Teams Administrator |

### Deprecated Module Warning

> **Warning:** The `AzureAD` and `MSOnline` modules are deprecated by Microsoft and must **not** be installed alongside these scripts. This platform uses `Microsoft.Graph` exclusively. If either deprecated module is present in your environment, uninstall it before running scripts.
>
> To check: `Get-Module AzureAD, MSOnline -ListAvailable`
> To remove: `Uninstall-Module AzureAD, MSOnline -AllVersions`

### PnP Entra App Registration

SharePoint Online and OneDrive modify scripts use `PnP.PowerShell`. The previously available built-in app registration was removed in September 2024. You must register your own Entra application.

1. Register a new application in the Entra portal (name: `PnP-Platform` or similar).
2. Add the API permissions required for your workload (typically `Sites.FullControl.All`, `User.ReadWrite.All`).
3. Grant admin consent.
4. Note the **Application (client) ID**.
5. Set the `PNP_CLIENT_ID` environment variable before running platform scripts. If you connect manually outside the platform, you can also pass `-ClientId` to `Connect-PnPOnline`.
6. `Initialize-TenantShift.ps1 -Profile OnlineOperator` validates that `PNP_CLIENT_ID` is configured.

---

## 2. Preparing the Input CSV

### Finding the Template

Each script ships with a `*.input.csv` template in the same directory:

```
.\TenantShift\Online\Modify\M-MEID-0010-Update-EntraUsers.ps1
.\TenantShift\Online\Modify\M-MEID-0010-Update-EntraUsers.input.csv    ← template
```

Copy the template, populate with your target records, and save.

### Encoding

Use **UTF-8 without BOM**. Re-save Excel exports as UTF-8 without BOM.

### Required vs. Optional Columns

The script's `.NOTES` block contains the CSV field reference table:

```powershell
Get-Help .\M-MEID-0010-Update-EntraUsers.ps1 -Full
```

`Import-ValidatedCsv` throws a clear error listing missing required columns before any rows are processed.

### Pre-Flight CSV Validation

```powershell
(Import-Csv .\myinput.csv | Select-Object -First 1).PSObject.Properties.Name
```

---

## 3. Running a Script

### Always Run -WhatIf First

```powershell
.\M-MEID-0010-Update-EntraUsers.ps1 -InputCsvPath .\myinput.csv -WhatIf
```

Every row shows `Status: WhatIf` with the `Message` column describing the change that would be applied. Review before proceeding to a live run.

### Live Run (CSV-Bounded)

```powershell
.\M-MEID-0010-Update-EntraUsers.ps1 -InputCsvPath .\myinput.csv
```

The script updates each target object and records the outcome for every row.

### Setting a Custom Output Path

```powershell
.\M-MEID-0010-Update-EntraUsers.ps1 -InputCsvPath .\myinput.csv -OutputCsvPath C:\Reports\UpdateUsers_2026-03.csv
```

---

## 4. Reading the Output

### Output File Location

Results are in the path specified by `-OutputCsvPath`. Default location:

```
.\TenantShift\Online\Modify\Modify_OutputCsvPath\
```

### Status Field Values

| Value | Meaning |
|-------|---------|
| `Success` | Object updated successfully |
| `Skipped` | Object already in the desired state, or a pre-condition was not met; no change made |
| `WhatIf` | Script ran with `-WhatIf`; no changes were applied |
| `Failed` | Update attempted and failed; see `Message` column for error detail |

### Message Field

Provides the reason for `Skipped` rows (e.g., "Attribute already set to target value") and the error message for `Failed` rows.

### Transcript Files

Transcripts are written alongside the results CSV. Retain for a minimum of **90 days**.

---

## 5. Common Errors and Remediation

| Error | Cause | Remediation |
|-------|-------|-------------|
| `Insufficient privileges` | Account lacks the write role for this workload | Assign the required role (see Prerequisites) |
| `HTTP 429 Too Many Requests` | API throttling | Platform retries automatically via `Invoke-WithRetry`. If persistent, reduce batch size |
| `Object not found` | Input row references an object that does not exist | Row is set to `Failed`. Verify the primary key in the input CSV. Use a Provision script to create missing objects first |
| `Connect-PnPOnline: ClientId not found` | PnP Entra app not registered or `PNP_CLIENT_ID` not configured | Complete the Entra app registration and set `PNP_CLIENT_ID` (see Prerequisites) |
| `Import-ValidatedCsv: required column missing` | CSV header mismatch | Check `Get-Help <script> -Full` for the CSV field table |
| `AzureAD` / `MSOnline` conflict | Deprecated modules installed | Uninstall: `Uninstall-Module AzureAD, MSOnline -AllVersions` |
| `ShouldProcess not supported` | Script called with `-WhatIf` but does not declare `SupportsShouldProcess` | This indicates a platform bug — report it. Do not run the live version until resolved |
