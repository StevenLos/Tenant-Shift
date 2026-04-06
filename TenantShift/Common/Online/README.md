# Common Online

`Common/Online` contains the environment module for Online (M365) workload scripts.

## Contents

| File | Description |
|------|-------------|
| `M365.Common.psm1` | Environment module for Online scripts — provides connectivity, validation, and result helpers; dot-sources `Shared.Common.ps1` |

## Usage

Online scripts import this module at the top of each script:

```powershell
Import-Module "$PSScriptRoot\..\..\Common\Online\M365.Common.psm1" -Force
```

The module dot-sources `Common/Shared/Shared.Common.ps1` automatically using a `Get-Command Write-Status` guard to prevent double-loading.

## References

- [Common/Shared README](../Shared/README.md)
- [Common README](../README.md)
