# Operator Runbook — OnPrem Modify Scripts

This runbook covers running `M-WW-NNNN` scripts that update existing on-premises Active Directory and Exchange objects. Modify scripts change attributes of objects that already exist. They support `-WhatIf` for dry-run validation and write a result CSV documenting every row processed.

> **Always run with `-WhatIf` first.** Modify scripts make live changes to existing AD objects. A dry run shows what would change without applying it.

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
| Exchange Management Shell | On-premises Exchange server | Required for Exchange OnPrem scripts (M-EXOP-xxxx) |

To install RSAT on Windows 10/11:

```powershell
Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'
```

### Execution Policy

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Permissions

OnPrem Modify scripts update existing AD objects. Required permissions:

| Workload | Minimum permission |
|----------|--------------------|
| Active Directory (AD) | Write permission on the target attribute(s) for the target OU (delegated or Domain Administrator) |
| Exchange On-Premises (XC) | Exchange Recipient Administrator or Recipient Management role group |

The specific attributes being changed determine the delegation needed. If you are unsure whether your account has the required permission, run `-WhatIf` first — if the permission check fails before the write, the row will show `Failed` with an access error.

### AD Replication Latency

> **Note:** Changes made by Modify scripts are written to a single domain controller and then replicate to other DCs. Updated attributes may not be visible immediately on all domain controllers. Allow sufficient replication time (typically 15 seconds within a site; up to 3 hours across sites by default) before running a follow-up IR script to verify the change. Use the `-Server` parameter to target the same DC used for the update if immediate verification is needed.

---

## 2. Preparing the Input CSV

### Finding the Template

Each script ships with a `*.input.csv` template in the same directory:

```
.\SharedModule\OnPrem\Modify\M-ADUC-0020-Update-ActiveDirectoryUsers.ps1
.\SharedModule\OnPrem\Modify\M-ADUC-0020-Update-ActiveDirectoryUsers.input.csv    ← template
```

Copy the template, populate with target records and desired values, and save.

### Encoding

Use **UTF-8 without BOM**, or plain UTF-8.

### Required vs. Optional Columns

Check the script's `.NOTES` block:

```powershell
Get-Help .\M-ADUC-0020-Update-ActiveDirectoryUsers.ps1 -Full
```

`Import-ValidatedCsv` validates all required headers before processing begins.

### Pre-Flight CSV Validation

```powershell
(Import-Csv .\myinput.csv | Select-Object -First 1).PSObject.Properties.Name
```

### Data Quality Checks Before Running

Before running a Modify script live:

1. Verify that the primary key column (UPN, SAMAccountName, etc.) matches existing objects. Run the corresponding IR script against the same input to confirm objects exist.
2. Verify that the new values in the CSV are valid for the target attribute (e.g., valid OU DN for a move operation).
3. Run with `-WhatIf` and review the output.

---

## 3. Running a Script

### Always Run -WhatIf First

```powershell
.\M-ADUC-0020-Update-ActiveDirectoryUsers.ps1 -InputCsvPath .\myinput.csv -WhatIf
```

Each row shows `Status: WhatIf` and a `Message` describing the change that would be applied. Investigate unexpected results before running live.

### Live Run (CSV-Bounded)

After `-WhatIf` validation:

```powershell
.\M-ADUC-0020-Update-ActiveDirectoryUsers.ps1 -InputCsvPath .\myinput.csv
```

### Setting a Custom Output Path

```powershell
.\M-ADUC-0020-Update-ActiveDirectoryUsers.ps1 -InputCsvPath .\myinput.csv -OutputCsvPath C:\Reports\UpdateADUsers_2026-03.csv
```

---

## 4. Reading the Output

### Output File Location

Results are written to the path specified by `-OutputCsvPath`. Default:

```
.\SharedModule\OnPrem\Modify\Modify_OutputCsvPath\
```

### Status Field Values

| Value | Meaning |
|-------|---------|
| `Success` | Object updated successfully |
| `Skipped` | Object already in the desired state, or a pre-condition was not met; no change made |
| `WhatIf` | Script ran with `-WhatIf`; no changes were applied |
| `Failed` | Update attempted and failed; see `Message` column for error detail |

### Verifying Changes

After a successful run, verify the updated attributes using the corresponding IR script:

```powershell
.\D-ADUC-0020-Get-ActiveDirectoryUsers.ps1 -InputCsvPath .\myinput.csv
```

Due to AD replication latency, allow time for replication before verifying, or use `-Server` to target the DC where changes were written.

### Transcript Files

Transcripts are written alongside the results CSV. Retain for a minimum of **90 days**.

---

## 5. Common Errors and Remediation

| Error | Cause | Remediation |
|-------|-------|-------------|
| `The term 'Set-ADUser' is not recognized` | RSAT not installed | Install RSAT (see Prerequisites) |
| `Access is denied` | Account lacks write permission on the target attribute | Request delegation or run as Domain Administrator |
| `Cannot find object with identity` | Primary key in CSV does not match any AD object | Verify the identity value; check for UPN vs. SAMAccountName mismatch |
| `The value violates attribute syntax constraints` | New value in CSV is invalid for the attribute type | Correct the value in the input CSV and re-run with `-WhatIf` |
| `Import-ValidatedCsv: required column missing` | CSV header mismatch | Check `Get-Help <script> -Full` for the CSV field table |
| `Object not visible after update` | AD replication not yet complete | Wait for replication or use `-Server` to query the DC where the change was written |
| `Protected from accidental deletion` | Object is protected and cannot be moved or renamed | Uncheck "Protect object from accidental deletion" in ADUC, or use appropriate AD delegation |
