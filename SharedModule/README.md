# SharedModule Folder

`SharedModule` defines the standard repository-native script model.

Use this model when scripts are expected to run as part of the repository's shared-module architecture:

- online scripts live under `SharedModule/Online/` and on-prem scripts live under `SharedModule/OnPrem/`
- shared helper/runtime logic comes from `SharedModule/Common/*`
- contracts/tests/docs/build artifacts are updated with the script
- shared-module script filenames use the `<P|M|D>-<WW>-<NNNN>-<Action>-<Target>.ps1` pattern

`SharedModule` is an index and governance location. It intentionally does **not** duplicate source scripts from `SharedModule/Online/`, `SharedModule/OnPrem/`, or `SharedModule/Common/`.
Externally sourced artifacts should be staged in `Imported/` and treated as read-only.

## Purpose

Use the shared-module model when a script should be:

- reusable beyond a single incident
- aligned to the `P/M/D` lifecycle catalog
- maintained under shared module/contract/test expectations

## Contents

| File / Folder | Description |
|---|---|
| `README.md` | This file — shared-module model policy |
| `CONTRIBUTING.md` | Developer guide: naming conventions, script skeleton, quality gates, circular dependency rule |
| `CHANGELOG.md` | Release history for SharedModule platform changes |
| `SharedModule-Repository-Portions.md` | Required repository areas for shared-module scripting |
| `SharedModule-Repository-Portions.csv` | Machine-readable inclusion manifest |
| `Common/` | Shared utility layer (`Shared.Common.ps1`) and environment modules (`M365.Common.psm1`, `OnPrem.Common.psm1`) |
| `Online/` | Cloud workload scripts (Entra, Exchange Online, SharePoint, OneDrive, Teams) |
| `OnPrem/` | On-prem workload scripts (ActiveDirectory, ExchangeOnPrem, FileServices) |
| `Development/Build/` | Build automation, contract tests, quality gate, and inventory scripts |
| `Development/Tests/` | Pester test suite for platform contracts and build utilities |
| `Utilities/` | Operator and contributor utilities — run `Initialize-Platform.ps1 -Profile Contributor` before contributing |

## Developer Quickstart

1. Run `Utilities/Initialize-Platform.ps1 -Profile Contributor` to validate your environment before contributing.
2. Use the workload runbooks for operator-host validation with `-Profile OnlineOperator` or `-Profile OnPremOperator`.
3. Read `CONTRIBUTING.md` for naming conventions, the script skeleton, and the quality gate workflow.
4. Before each commit, run `Development/Build/Invoke-QualityGate.ps1 -Path <your-script>`.

## References

- [Root README](../README.md)
- [Imported README](../Imported/README.md)
- [CONTRIBUTING.md](./CONTRIBUTING.md)
- [CHANGELOG.md](./CHANGELOG.md)
- [Online README](./Online/README.md)
- [OnPrem README](./OnPrem/README.md)
- [Common README](./Common/README.md)
- [Build README](./Development/Build/README.md)
