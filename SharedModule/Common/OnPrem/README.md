# Common OnPrem

`Common/OnPrem` contains the environment module for on-prem workload scripts.

## Contents

| File | Description |
|------|-------------|
| `OnPrem.Common.psm1` | Environment module for OnPrem scripts — provides connectivity, validation, and result helpers; dot-sources `Shared.Common.ps1` |

## Usage

OnPrem scripts import this module at the top of each script:

```powershell
Import-Module "$PSScriptRoot\..\..\Common\OnPrem\OnPrem.Common.psm1" -Force
```

The module dot-sources `Common/Shared/Shared.Common.ps1` automatically using a `Get-Command Write-Status` guard to prevent double-loading.

## References

- [Common/Shared README](../Shared/README.md)
- [Common README](../README.md)
