# Operator Runbook — OnPrem Provision Scripts

This runbook covers running `P-WW-NNNN` scripts that create new on-premises Active Directory and Exchange objects. Provision scripts create objects that do not yet exist. They support `-WhatIf` for dry-run validation and write a result CSV documenting every row processed.

> **Always run with `-WhatIf` first.** Provision scripts create live AD objects. A dry run shows exactly what would be created without making any changes.

---

## 1. Prerequisites

### Required PowerShell Version

Windows PowerShell **5.1** is required. The `ActiveDirectory` module ships with Windows RSAT and is not available in PowerShell 7.

```powershell
$PSVersionTable.PSVersion
```

### Step 1: Run Initialize-Platform.ps1

**Run this before anything else.** Use the exact OnPrem operator profile. It validates that you are on a Windows host, running Windows PowerShell 5.1 Desktop, and have the `ActiveDirectory` module available. Exchange Management Shell requirements remain workload-specific and must be validated manually.

```powershell
.\SharedModule\Utilities\Initialize-Platform.ps1 -Profile OnPremOperator
```

Resolve all failures before proceeding.

### Required Modules

| Module | Source | Notes |
|--------|--------|-------|
| `ActiveDirectory` | Windows RSAT | Required for all AD scripts |
| Exchange Management Shell | On-premises Exchange server | Required for Exchange OnPrem scripts (P-EXOP-xxxx) |

To install RSAT on Windows 10/11:

```powershell
Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'
```

### Execution Policy

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Permissions

OnPrem Provision scripts create new AD objects. Required permissions:

| Workload | Minimum permission |
|----------|--------------------|
| Active Directory (AD) | Create permission on the target OU (delegated or Domain Administrator) |
| Exchange On-Premises (XC) | Exchange Recipient Administrator or Recipient Management role group |

Ensure your account has create permission on the target OU before running. Use the `-SearchBase` parameter (if available on the script) to confirm OU targeting during a `-WhatIf` run.

### AD Replication Latency

> **Note:** Objects created by Provision scripts are written to a single domain controller and then replicate to other DCs. New objects may not be visible immediately on all domain controllers. Allow sufficient time for replication (typically 15 seconds within a site; up to 3 hours across sites by default) before running a follow-up IR script to verify creation. Use the `-Server` parameter to target the same DC used for creation if immediate verification is needed.

---

## 2. Preparing the Input CSV

### Finding the Template

Each script ships with a `*.input.csv` template in the same directory:

```
.\SharedModule\OnPrem\Provision\P-ADUC-0020-Create-ActiveDirectoryUsers.ps1
.\SharedModule\OnPrem\Provision\P-ADUC-0020-Create-ActiveDirectoryUsers.input.csv    ← template
```

Copy the template, populate with your data, and save it.

### Encoding

Use **UTF-8 without BOM**, or plain UTF-8. Windows PowerShell 5.1 accepts both.

### Required vs. Optional Columns

Check the script's `.NOTES` block for the CSV field reference table:

```powershell
Get-Help .\P-ADUC-0020-Create-ActiveDirectoryUsers.ps1 -Full
```

`Import-ValidatedCsv` validates headers before processing any rows and throws a clear error if required columns are missing.

### Pre-Flight CSV Validation

```powershell
(Import-Csv .\myinput.csv | Select-Object -First 1).PSObject.Properties.Name
```

### Data Quality Checks Before Running

Before running a Provision script live:

1. Verify that required fields (UPN, SAMAccountName, OU path) are populated for every row.
2. Verify that target OUs exist in AD. An invalid OU path causes the row to fail.
3. Run a corresponding IR script against the input CSV to confirm objects do not already exist.

---

## 3. Running a Script

### Always Run -WhatIf First

```powershell
.\P-ADUC-0020-Create-ActiveDirectoryUsers.ps1 -InputCsvPath .\myinput.csv -WhatIf
```

Every row shows `Status: WhatIf` with the `Message` column describing what would be created. Investigate any row that shows an unexpected message before proceeding to a live run.

### Live Run (CSV-Bounded)

After validating with `-WhatIf`:

```powershell
.\P-ADUC-0020-Create-ActiveDirectoryUsers.ps1 -InputCsvPath .\myinput.csv
```

### Setting a Custom Output Path

```powershell
.\P-ADUC-0020-Create-ActiveDirectoryUsers.ps1 -InputCsvPath .\myinput.csv -OutputCsvPath C:\Reports\CreateADUsers_2026-03.csv
```

---

## 4. Reading the Output

### Output File Location

Results are written to the path specified by `-OutputCsvPath`. Default:

```
.\SharedModule\OnPrem\Provision\Provision_OutputCsvPath\
```

### Status Field Values

| Value | Meaning |
|-------|---------|
| `Success` | Object created successfully |
| `Skipped` | Object already exists, or a pre-condition was not met; no change made |
| `WhatIf` | Script ran with `-WhatIf`; no objects were created |
| `Failed` | Creation attempted and failed; see `Message` column for error detail |

### Verifying Creation

After a successful run, verify the created objects using the corresponding IR script:

```powershell
.\D-ADUC-0020-Get-ActiveDirectoryUsers.ps1 -InputCsvPath .\myinput.csv
```

Due to AD replication latency, run the verification against the same domain controller used by the Provision script (use `-Server` if needed).

### Transcript Files

Transcripts are written alongside the results CSV. Retain for a minimum of **90 days**.

---

## 5. Common Errors and Remediation

| Error | Cause | Remediation |
|-------|-------|-------------|
| `The term 'New-ADUser' is not recognized` | RSAT not installed | Install RSAT (see Prerequisites) |
| `Access is denied` | Account lacks create permission on the target OU | Request OU delegation or run as Domain Administrator |
| `The object already exists` | Input row targets an existing object | Row is set to `Skipped`. Verify input data; use a Modify script to update existing objects |
| `The specified OU does not exist` | OU path in CSV is invalid | Verify the distinguished name path; create the OU first if needed |
| `Import-ValidatedCsv: required column missing` | CSV header mismatch | Check `Get-Help <script> -Full` for the CSV field table |
| `Object not visible after creation` | AD replication not yet complete | Wait for replication or use `-Server` to target the DC where the object was created |
| `Password does not meet complexity requirements` | Initial password in CSV fails AD policy | Check the domain password policy and update the password in the input CSV |
