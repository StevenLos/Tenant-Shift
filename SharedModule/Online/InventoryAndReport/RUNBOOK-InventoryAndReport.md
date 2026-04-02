# Operator Runbook — Online InventoryAndReport Scripts

This runbook covers running `D-WW-NNNN` scripts that inventory Microsoft 365 resources and export results to CSV. These scripts are **read-only** — they never modify data. Use them to produce reports, validate provisioning state, or gather data before running Provision or Modify scripts.

---

## 1. Prerequisites

### Required PowerShell Version

PowerShell **7.0 or later** is required. Online scripts do not run on Windows PowerShell 5.1.

Verify your version:

```powershell
$PSVersionTable.PSVersion
```

### Step 1: Run Initialize-Platform.ps1

**Run this before anything else.** Use the exact Online operator profile. It validates PowerShell 7+, `Microsoft.Graph.Authentication`, `ExchangeOnlineManagement`, `PnP.PowerShell`, and the `PNP_CLIENT_ID` environment variable.

```powershell
.\SharedModule\Utilities\Initialize-Platform.ps1 -Profile OnlineOperator
```

Resolve any failures reported before proceeding.

### Required Modules

| Module | Minimum version | Purpose |
|--------|----------------|---------|
| `Microsoft.Graph.*` | Latest | Entra / identity queries |
| `ExchangeOnlineManagement` | 3.x | Exchange Online queries |
| `PnP.PowerShell` | Latest | SharePoint Online / OneDrive queries |
| `MicrosoftTeams` | Latest | Teams queries |

Resolve missing or outdated modules manually using the remediation command that `Initialize-Platform.ps1 -Profile OnlineOperator` prints. The script does not install or update modules automatically.

### Execution Policy

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Permissions

Permissions vary by workload. Minimum read-only roles:

| Workload | Required role |
|----------|--------------|
| Entra / Graph (GR) | Global Reader or Reports Reader |
| Exchange Online (EX) | Exchange Recipient Administrator or View-Only Recipients |
| SharePoint / OneDrive (SP) | SharePoint Administrator (read-only access requires admin role) |
| Teams (TM) | Teams Administrator or Teams Communications Support Engineer |

Connect to the required services before running the script. The scripts call `Ensure-GraphConnection`, `Ensure-ExchangeConnection`, or `Ensure-SharePointConnection` automatically — these will prompt for credentials if no active session exists.

### Deprecated Module Warning

> **Warning:** The `AzureAD` and `MSOnline` modules are deprecated by Microsoft and must **not** be installed alongside these scripts. This platform uses `Microsoft.Graph` exclusively for identity operations. If either deprecated module is present in your environment, uninstall it before running scripts to prevent cmdlet name conflicts and authentication interference.
>
> To check: `Get-Module AzureAD, MSOnline -ListAvailable`
> To remove: `Uninstall-Module AzureAD, MSOnline -AllVersions`

### PnP Entra App Registration

SharePoint Online and OneDrive scripts use `PnP.PowerShell`. The previously available built-in app registration shipped with PnP was removed in September 2024. You must register your own Entra application.

1. In the Azure / Entra portal, register a new application (name: `PnP-Platform` or similar).
2. Add API permissions required for your workload (typically `Sites.Read.All`, `User.Read.All`).
3. Grant admin consent.
4. Note the **Application (client) ID**.
5. Set the `PNP_CLIENT_ID` environment variable before running platform scripts. If you connect manually outside the platform, you can also pass `-ClientId` to `Connect-PnPOnline`.
6. `Initialize-Platform.ps1 -Profile OnlineOperator` validates that `PNP_CLIENT_ID` is configured.

> **Warning:** The `AzureAD` and `MSOnline` modules conflict with PnP authentication. Ensure they are not installed (see above).

### EXO Result Truncation Warning

> **Warning:** Exchange Online commands use a default result limit that silently truncates output. Any IR script that retrieves all mailboxes, recipients, or distribution group members must pass `-ResultSize Unlimited` to the underlying EXO command. Platform IR scripts do this automatically when run in `-DiscoverAll` mode. If you are calling EXO commands directly or writing a custom script, you **must** include `-ResultSize Unlimited` or your results will be incomplete without any error or warning.

---

## 2. Preparing the Input CSV

IR scripts support two modes: CSV-bounded (process only the objects listed in the CSV) and DiscoverAll (enumerate all objects in scope). For CSV-bounded runs, you need an input CSV.

### Finding the Template

Each script ships with a `*.input.csv` template file in the same directory. The template contains only the header row with the required column names.

```
.\SharedModule\Online\InventoryAndReport\D-MEID-0010-Get-EntraUsers.ps1
.\SharedModule\Online\InventoryAndReport\Scope-Users.input.csv    ← template
```

### Encoding

Use **UTF-8 without BOM**. Most spreadsheet editors export UTF-8 with BOM — save as "UTF-8" not "UTF-8 with BOM" when preparing input files.

### Required vs. Optional Columns

Each script's `.NOTES` block documents which columns are required and which are optional:

```powershell
Get-Help .\D-MEID-0010-Get-EntraUsers.ps1 -Full
```

The script calls `Import-ValidatedCsv` internally, which throws a clear error listing any missing required columns before processing begins.

### Validating the CSV Before Running

To check that your CSV has the right headers without running the script:

```powershell
$headers = (Import-Csv .\myinput.csv | Select-Object -First 1).PSObject.Properties.Name
$headers
```

---

## 3. Running a Script

### CSV-Bounded Mode

Process only the objects listed in the input CSV:

```powershell
.\D-MEID-0010-Get-EntraUsers.ps1 -InputCsvPath .\myinput.csv
```

Use this when you need to check a specific list of objects (e.g., a migration batch, a ticket list).

### DiscoverAll Mode

Enumerate all objects in scope without a CSV:

```powershell
.\D-MEID-0010-Get-EntraUsers.ps1 -DiscoverAll
```

Use this for full-environment audits, baseline captures, or when you do not have a pre-defined target list.

> **Note:** DiscoverAll runs against the full tenant. For large environments this can take significant time and consume API quota. Avoid running DiscoverAll repeatedly in quick succession — use the CSV output from a previous run to target specific objects instead.

### EXO Result Truncation Warning (Repeated)

> **Warning:** When running Exchange Online IR scripts in `-DiscoverAll` mode, the platform passes `-ResultSize Unlimited` automatically. If you call EXO cmdlets directly in a custom script, you **must** include `-ResultSize Unlimited` or output will be silently truncated.

### Setting a Custom Output Path

By default, results are written to a timestamped file under `InventoryAndReport_OutputCsvPath\` in the script directory. To write results elsewhere:

```powershell
.\D-MEID-0010-Get-EntraUsers.ps1 -DiscoverAll -OutputCsvPath C:\Reports\EntraUsers_2026-03.csv
```

---

## 4. Reading the Output

### Output File Location

Results are written to the path specified by `-OutputCsvPath`. If you used the default, look in:

```
.\SharedModule\Online\InventoryAndReport\InventoryAndReport_OutputCsvPath\
```

A transcript of the run is written alongside the results file.

### Status Field Values

Every output row contains a `Status` column:

| Value | Meaning |
|-------|---------|
| `Success` | Object retrieved and all fields populated |
| `Failed` | Object lookup failed; see `Message` column for the error detail |
| `Skipped` | Object was in the input CSV but could not be found or was excluded by a pre-condition |

### Message Field

The `Message` column contains human-readable detail for `Failed` and `Skipped` rows. For `Success` rows it may contain supplementary information.

### Transcript Files

Transcripts are written to the same output directory as the results CSV. Retain transcripts for a minimum of **90 days** for audit purposes. Do not commit transcript files to source control.

---

## 5. Common Errors and Remediation

| Error | Cause | Remediation |
|-------|-------|-------------|
| `Connect-MgGraph: insufficient privileges` | Account lacks required Graph role | Assign the required role (see Prerequisites) or re-connect with an account that has it |
| `HTTP 429 Too Many Requests` | Graph or EXO throttling | The platform retries automatically via `Invoke-WithRetry`. If it persists, wait and retry the run |
| `The term 'Get-EXOMailbox' is not recognized` | `ExchangeOnlineManagement` not installed or wrong version | Install or update `ExchangeOnlineManagement`, then re-run `Initialize-Platform.ps1 -Profile OnlineOperator` |
| `Connect-PnPOnline: AADSTS700016 / ClientId not found` | PnP Entra app registration missing or `PNP_CLIENT_ID` not configured | Complete the PnP Entra app registration and set `PNP_CLIENT_ID` (see Prerequisites) |
| `Import-ValidatedCsv: required column missing` | Input CSV header does not match the script's expected columns | Check `Get-Help <script> -Full` for the `.NOTES` CSV field table; regenerate from the `.input.csv` template |
| `UTF-8 BOM encoding error` | Input CSV saved with BOM from Excel | Re-save as UTF-8 without BOM, or open in Notepad and save as UTF-8 |
| `AzureAD module conflict` | Both AzureAD and Microsoft.Graph modules installed | Uninstall: `Uninstall-Module AzureAD -AllVersions` |
| `MSOnline module conflict` | MSOnline module still installed | Uninstall: `Uninstall-Module MSOnline -AllVersions` |
