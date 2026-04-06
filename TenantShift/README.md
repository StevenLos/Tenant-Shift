# TenantShift Folder

`TenantShift` defines the standard repository-native script model.

Use this model when scripts are expected to run as part of the repository's tenant-shift architecture:

- online scripts live under `TenantShift/Online/` and on-prem scripts live under `TenantShift/OnPrem/`
- shared helper/runtime logic comes from `TenantShift/Common/*`
- contracts/tests/docs/build artifacts are updated with the script
- tenant-shift script filenames use the `<P|M|D>-<WW>-<NNNN>-<Action>-<Target>.ps1` pattern

`TenantShift` is an index and governance location. It intentionally does **not** duplicate source scripts from `TenantShift/Online/`, `TenantShift/OnPrem/`, or `TenantShift/Common/`.
Externally sourced artifacts should be staged in `TenantShift/Development/Imported/` and treated as read-only.

## Purpose

Use the tenant-shift model when a script should be:

- reusable beyond a single incident
- aligned to the `P/M/D` lifecycle catalog
- maintained under tenant shift/contract/test expectations

## Contents

| File / Folder | Description |
|---|---|
| `README.md` | This file — tenant-shift model policy |
| `Development/CONTRIBUTING.md` | Developer guide: naming conventions, script skeleton, quality gates, circular dependency rule |
| `CHANGELOG.md` | Release history for TenantShift platform changes |
| `TenantShift-Repository-Portions.md` | Required repository areas for tenant-shift scripting |
| `TenantShift-Repository-Portions.csv` | Machine-readable inclusion manifest |
| `Common/` | Shared utility layer (`Shared.Common.ps1`) and environment modules (`M365.Common.psm1`, `OnPrem.Common.psm1`) |
| `Online/` | Cloud workload scripts (Entra, Exchange Online, SharePoint, OneDrive, Teams) |
| `OnPrem/` | On-prem workload scripts (ActiveDirectory, ExchangeOnPrem, FileServices) |
| `Development/Build/` | Build automation, contract tests, quality gate, and inventory scripts |
| `Development/Tests/` | Pester test suite for platform contracts and build utilities |
| `Utilities/` | TenantShift cross-workload helper scripts (CSV converters, password generator, prerequisite scanner, file unblock) |

## Developer Quickstart

1. Run `../Initialize-TenantShift/Initialize-TenantShift.ps1 -Profile Contributor` to validate your environment before contributing.
2. Use the workload runbooks for operator-host validation with `-Profile OnlineOperator` or `-Profile OnPremOperator`.
3. Read [`CONTRIBUTING.md`](./Development/CONTRIBUTING.md) for naming conventions, the script skeleton, and the quality gate workflow.
4. Before each commit, run `Development/Build/Invoke-QualityGate.ps1 -Path <your-script>`.

## References

- [Root README](../README.md)
- [Imported README](./Development/Imported/README.md)
- [CONTRIBUTING.md](./Development/CONTRIBUTING.md)
- [CHANGELOG.md](./CHANGELOG.md)
- [Online README](./Online/README.md)
- [OnPrem README](./OnPrem/README.md)
- [Common README](./Common/README.md)
- [Build README](./Development/Build/README.md)
