# Operator Runbook — Online Provision Scripts

This runbook covers running `P-WW-NNNN` scripts that create new Microsoft 365 objects. Provision scripts create objects that do not yet exist. They support `-WhatIf` for dry-run validation and write a result CSV documenting every row processed.

> **Always run with `-WhatIf` first.** Provision scripts create live objects. A dry run with `-WhatIf` shows exactly what would be created without making any changes.

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

Resolve all failures before proceeding. Do not run Provision scripts in an environment that has not passed `Initialize-TenantShift.ps1`.

### Required Modules

| Module | Minimum version | Purpose |
|--------|----------------|---------|
| `Microsoft.Graph.*` | Latest | Entra / identity operations |
| `ExchangeOnlineManagement` | 3.x | Exchange Online provisioning |
| `PnP.PowerShell` | Latest | SharePoint Online / OneDrive provisioning |
| `MicrosoftTeams` | Latest | Teams provisioning |

### Execution Policy

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Permissions

Provision scripts make changes. Ensure your account has write permissions before running.

| Workload | Required role |
|----------|--------------|
| Entra users / groups (GR) | User Administrator or Groups Administrator |
| Exchange Online (EX) | Exchange Administrator or Exchange Recipient Administrator |
| SharePoint / OneDrive (SP) | SharePoint Administrator |
| Teams (TM) | Teams Administrator |

### Deprecated Module Warning

> **Warning:** The `AzureAD` and `MSOnline` modules are deprecated by Microsoft and must **not** be installed alongside these scripts. This platform uses `Microsoft.Graph` exclusively for identity operations. If either deprecated module is present in your environment, uninstall it before running scripts to prevent cmdlet name conflicts and authentication interference.
>
> To check: `Get-Module AzureAD, MSOnline -ListAvailable`
> To remove: `Uninstall-Module AzureAD, MSOnline -AllVersions`

### PnP Entra App Registration

SharePoint Online and OneDrive provision scripts use `PnP.PowerShell`. The previously available built-in app registration shipped with PnP was removed in September 2024. You must register your own Entra application.

1. Register a new application in the Entra portal (name: `PnP-Platform` or similar).
2. Add the required API permissions for your workload (typically `Sites.FullControl.All`, `User.ReadWrite.All`).
3. Grant admin consent.
4. Note the **Application (client) ID**.
5. Set the `PNP_CLIENT_ID` environment variable before running platform scripts. If you connect manually outside the platform, you can also pass `-ClientId` to `Connect-PnPOnline`.
6. `Initialize-TenantShift.ps1 -Profile OnlineOperator` validates that `PNP_CLIENT_ID` is configured.

### OneDrive 199-User Batch Limit

> **Note:** `Request-SPOPersonalSite` accepts a maximum of **199 users per call**. When provisioning OneDrive for more than 199 users, the platform scripts automatically chunk the input CSV into batches of 199. Each batch is submitted as a separate request. Do not attempt to pass more than 199 users in a single call outside the platform scripts.

### Provisioning Latency

> **Note:** OneDrive and SharePoint site provisioning can take up to **24 hours** to complete after a request is submitted. The output `Status` field uses `Submitted` to indicate that the request was accepted by the service, not that provisioning is complete. A `Submitted` result is **not** a confirmation of readiness. If you need to verify completion, re-run the corresponding D script after the latency window has passed to confirm `Status` is `Success`.

---

## 2. Preparing the Input CSV

### Finding the Template

Each script ships with a `*.input.csv` template in the same directory, containing only the header row:

```
.\TenantShift\Online\Provision\P-MEID-0010-Create-EntraUsers.ps1
.\TenantShift\Online\Provision\P-MEID-0010-Create-EntraUsers.input.csv    ← template
```

Copy the template, populate it with your data, and save it with a descriptive name.

### Encoding

Use **UTF-8 without BOM**. Excel exports UTF-8 with BOM by default — save as "UTF-8" (without BOM) when exporting from a spreadsheet editor.

### Required vs. Optional Columns

Check the script's `.NOTES` block for the CSV field reference table:

```powershell
Get-Help .\P-MEID-0010-Create-EntraUsers.ps1 -Full
```

The script calls `Import-ValidatedCsv` internally. If a required column is missing, the script throws a clear error before processing any rows.

### Pre-Flight CSV Validation

Before running, verify your CSV headers match expectations:

```powershell
(Import-Csv .\myinput.csv | Select-Object -First 1).PSObject.Properties.Name
```

---

## 3. Running a Script

### Always Run -WhatIf First

```powershell
.\P-MEID-0010-Create-EntraUsers.ps1 -InputCsvPath .\myinput.csv -WhatIf
```

Review the output CSV. Every row should show `Status: WhatIf` with a `Message` describing what would be created. Fix any issues in the input CSV before proceeding to a live run.

### Live Run (CSV-Bounded)

After validating with `-WhatIf`:

```powershell
.\P-MEID-0010-Create-EntraUsers.ps1 -InputCsvPath .\myinput.csv
```

The script processes every row, creates the object if conditions are met, and writes a result for each row.

### Setting a Custom Output Path

By default results go to a timestamped file under `Provision_OutputCsvPath\` in the script directory:

```powershell
.\P-MEID-0010-Create-EntraUsers.ps1 -InputCsvPath .\myinput.csv -OutputCsvPath C:\Reports\CreateUsers_2026-03.csv
```

---

## 4. Reading the Output

### Output File Location

Results are written to the path specified by `-OutputCsvPath`. If you used the default:

```
.\TenantShift\Online\Provision\Provision_OutputCsvPath\
```

### Status Field Values

| Value | Meaning |
|-------|---------|
| `Success` | Object created successfully |
| `Submitted` | Provisioning request accepted by the service but completion not yet confirmed — applies to async operations such as OneDrive site provisioning |
| `Skipped` | Object already exists or a pre-condition was not met; no change made |
| `WhatIf` | Script ran with `-WhatIf`; no objects were created |
| `Failed` | Creation was attempted and failed; see `Message` column for the error detail |

### Following Up on Submitted Rows

`Submitted` means the request was accepted, not that provisioning is complete. For OneDrive provisioning, wait up to 24 hours, then run the corresponding Discover script (e.g., `D-ONDR-0010-Get-OneDriveProvisioningStatus.ps1`) against the same input CSV to verify completion.

### Transcript Files

Transcripts are written to the same output directory. Retain for a minimum of **90 days**.

---

## 5. Common Errors and Remediation

| Error | Cause | Remediation |
|-------|-------|-------------|
| `Insufficient privileges` | Account lacks write role | Assign the required role (see Prerequisites) |
| `HTTP 429 Too Many Requests` | API throttling | Platform retries automatically. If persistent, reduce input CSV size and retry |
| `Object already exists` | Input row targets an object that already exists | Row is `Skipped` by default. Verify the input data; use a Modify script if you need to update an existing object |
| `Request-SPOPersonalSite: too many users` | Batch size exceeds 199 | Platform scripts chunk automatically. If calling directly, split into batches of ≤ 199 |
| `Connect-PnPOnline: ClientId not found` | PnP Entra app not registered or `PNP_CLIENT_ID` not configured | Complete the Entra app registration and set `PNP_CLIENT_ID` (see Prerequisites) |
| `Import-ValidatedCsv: required column missing` | CSV header mismatch | Check `Get-Help <script> -Full` for the CSV field table |
| `AzureAD` / `MSOnline` conflict | Deprecated modules installed | Uninstall both: `Uninstall-Module AzureAD, MSOnline -AllVersions` |
