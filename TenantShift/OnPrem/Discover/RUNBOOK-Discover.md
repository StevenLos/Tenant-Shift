# Operator Runbook — OnPrem Discover Scripts

This runbook covers running `D-WW-NNNN` scripts that inventory on-premises Active Directory and Exchange resources. These scripts are **read-only** — they never modify data. Use them to produce reports or gather baseline data before running Provision or Modify scripts.

---

## 1. Prerequisites

### Required PowerShell Version

Windows PowerShell **5.1** is required. These scripts use the `ActiveDirectory` module, which only ships with Windows RSAT and is not available in PowerShell 7.

> **Important:** Run these scripts from Windows PowerShell 5.1, **not** from PowerShell 7 (`pwsh`). The `ActiveDirectory` module does not load correctly in PowerShell 7 without compatibility shims that are not supported by this platform.

Verify your version:

```powershell
$PSVersionTable.PSVersion
```

### Step 1: Run Initialize-TenantShift.ps1

**Run this before anything else.** Use the exact OnPrem operator profile. It validates that you are on a Windows host, running Windows PowerShell 5.1 Desktop, and have the `ActiveDirectory` module available. Exchange Management Shell requirements remain workload-specific and must be validated manually.

```powershell
.\Initialize-TenantShift\Initialize-TenantShift.ps1 -Profile OnPremOperator
```

Resolve all failures before proceeding.

### Required Modules

| Module | Source | Notes |
|--------|--------|-------|
| `ActiveDirectory` | Windows RSAT | Required for all AD scripts |
| Exchange Management Shell | On-premises Exchange server | Required for Exchange OnPrem scripts (D-EXOP-xxxx) |

To install RSAT on Windows 10/11:

```powershell
Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'
```

### Execution Policy

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Permissions

OnPrem D scripts query Active Directory and on-premises Exchange. Required permissions:

| Workload | Minimum permission |
|----------|--------------------|
| Active Directory (AD) | Domain Users (read access is sufficient for most D scripts); some scripts require delegation to read confidential attributes |
| Exchange On-Premises (XC) | View-Only Organization Management or View-Only Recipients role group |

If your account does not have read access to the target OU, use `-SearchBase` to scope queries to OUs where your account has delegation.

### AD Replication Latency

> **Note:** Active Directory uses multi-master replication. Changes made by Provision or Modify scripts may not be visible immediately on all domain controllers. If you are running an D script to verify a recent Provision or Modify operation, allow sufficient time for replication (typically 15 seconds within a site; up to 3 hours across sites for the default schedule). Use the `-Server` parameter to target a specific domain controller if you need to read from the DC where the change was written.

---

## 2. Preparing the Input CSV

### Finding the Template

Each script ships with a `*.input.csv` template in the same directory:

```
.\TenantShift\OnPrem\Discover\D-ADUC-0020-Get-ActiveDirectoryUsers.ps1
.\TenantShift\OnPrem\Discover\Scope-ActiveDirectoryUsers.input.csv    ← template
```

### Encoding

Use **UTF-8 without BOM**, or plain UTF-8. Windows PowerShell 5.1 accepts both.

### Required vs. Optional Columns

Check the script's `.NOTES` block:

```powershell
Get-Help .\D-ADUC-0020-Get-ActiveDirectoryUsers.ps1 -Full
```

The script validates CSV headers via `Import-ValidatedCsv` before processing begins.

---

## 3. Running a Script

### CSV-Bounded Mode

Process only the objects in the input CSV:

```powershell
.\D-ADUC-0020-Get-ActiveDirectoryUsers.ps1 -InputCsvPath .\myinput.csv
```

### DiscoverAll Mode

Enumerate all objects in scope without a CSV:

```powershell
.\D-ADUC-0020-Get-ActiveDirectoryUsers.ps1 -DiscoverAll
```

To scope the discovery to a specific OU:

```powershell
.\D-ADUC-0020-Get-ActiveDirectoryUsers.ps1 -DiscoverAll -SearchBase 'OU=Users,DC=contoso,DC=com'
```

To target a specific domain controller:

```powershell
.\D-ADUC-0020-Get-ActiveDirectoryUsers.ps1 -DiscoverAll -Server dc01.contoso.com
```

### Limiting Result Count (Large Environments)

Use `-MaxObjects` to cap the number of objects returned in DiscoverAll mode. This is useful for testing or when sampling a large domain:

```powershell
.\D-ADUC-0020-Get-ActiveDirectoryUsers.ps1 -DiscoverAll -MaxObjects 500
```

`0` (the default) means no limit.

### Setting a Custom Output Path

```powershell
.\D-ADUC-0020-Get-ActiveDirectoryUsers.ps1 -DiscoverAll -OutputCsvPath C:\Reports\ADUsers_2026-03.csv
```

---

## 4. Reading the Output

### Output File Location

Results are written to the path specified by `-OutputCsvPath`. Default:

```
.\TenantShift\OnPrem\Discover\Discover_OutputCsvPath\
```

### Status Field Values

| Value | Meaning |
|-------|---------|
| `Success` | Object retrieved and all fields populated |
| `Failed` | Object lookup failed; see `Message` column for error detail |
| `Skipped` | Object was in the input CSV but could not be found or was excluded by a pre-condition |

### Transcript Files

Transcripts are written alongside the results CSV. Retain for a minimum of **90 days** for audit purposes. Do not commit transcript files to source control.

---

## 5. Common Errors and Remediation

| Error | Cause | Remediation |
|-------|-------|-------------|
| `The term 'Get-ADUser' is not recognized` | RSAT / ActiveDirectory module not installed | Install RSAT (see Prerequisites) |
| `Unable to contact the server` | Domain controller unreachable | Check network connectivity; use `-Server` to target a reachable DC explicitly |
| `Access denied` | Account lacks read permission on the target OU | Request delegation or use `-SearchBase` to scope to an OU where you have access |
| `Get-ADUser: directory object not found` | Object referenced in CSV does not exist in AD | Verify the identity value in the input CSV; check AD replication if object was recently created |
| `Import-ValidatedCsv: required column missing` | CSV header mismatch | Check `Get-Help <script> -Full` for the CSV field table |
| `Replication delay — object not visible` | AD replication not yet complete | Wait for replication, or use `-Server` to query the DC where the change was made |
