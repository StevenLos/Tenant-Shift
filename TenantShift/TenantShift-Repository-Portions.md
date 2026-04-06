# TenantShift Repository Portions

This document identifies which existing repository areas are part of the tenant-shift script model.

## Core Portions

| Portion | Path | Why It Is Included |
|---|---|---|
| Workload Script Planes | `TenantShift/Online/*` and `TenantShift/OnPrem/*` | Primary location for tenant-shift scripts (`Provision`, `Modify`, `Discover`) |
| Shared Modules | `TenantShift/Common/Online`, `TenantShift/Common/OnPrem`, `TenantShift/Common/Shared` | Shared helper logic, connection patterns, telemetry/result behavior |
| Build Automation | `TenantShift/Development/Build/*` | Orchestrator generation, roadmap/backlog, repository contract checks |
| Test Coverage | `TenantShift/Development/Tests/*` | Contract and behavior validation for tenant-shift scripts |
| Root Documentation | `README.md` and workload README/catalog files | Discoverability and execution guidance |

## Shared-Module Insertion Checklist

When adding or promoting a tenant-shift script, the following repository portions should be updated as applicable:

1. Script placement under correct workload and operation (`TenantShift/Online` or `TenantShift/OnPrem` + `Provision|Modify|Discover`).
2. Shared module reuse via `TenantShift/Common/Online/M365.Common.psm1` or `TenantShift/Common/OnPrem/OnPrem.Common.psm1`.
3. Script-specific input template (`*.input.csv`) or shared scope template in inventory.
4. Workload README/catalog update in the target operation folder.
5. Build/contract updates when new contracts or metadata rules are introduced.
6. Test additions/updates in `TenantShift/Development/Tests/` for contract and behavioral parity.

## Non-Goals

- Do not copy scripts into `TenantShift/`.
- Do not bypass common modules for tenant-shift scripts.
