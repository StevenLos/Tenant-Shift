# Online Folder

`Online` is the implementation area for cloud workloads targeting Microsoft 365 services.

Current status: Entra (`30xx`) and Exchange Online (`31xx`) provision/modify/inventory baselines are fully implemented. SharePoint/OneDrive (`32xx`) and Teams (`33xx`) script sets are implemented across all three operation types. See the operation folder READMEs for full script catalogs.

## Workload Codes

- `30xx`: Entra (users, groups, guests, licensing, privileged roles)
- `31xx`: Exchange Online (mailboxes, distribution lists, contacts, accepted domains)
- `32xx`: SharePoint / OneDrive (sites, hub sites, OneDrive provisioning, sharing, lock state)
- `33xx`: Teams (teams, channels, membership)

## Runtime Assumptions

- All Online scripts require PowerShell 7.0+.
- Required modules by workload:
  - Entra/Graph: `Microsoft.Graph.Authentication` 2.0+
  - Exchange Online: `ExchangeOnlineManagement` 3.0+
  - SharePoint/OneDrive: `PnP.PowerShell` 2.0+ (requires `PNP_CLIENT_ID` environment variable)
  - Teams: `MicrosoftTeams`

Run `Utilities/Initialize-Platform.ps1 -Profile OnlineOperator` to verify your operator environment before running Online scripts. Use `-Online` only when you also want the `Contributor` profile validated on the same host.

## Operation Folders

- `SharedModule/Online/Provision/` — initial object creation
- `SharedModule/Online/Modify/` — updates to existing objects
- `SharedModule/Online/InventoryAndReport/` — read-only inventory and reporting

## Script Matrix Overview

| Operation | Entra (`30xx`) | Exchange Online (`31xx`) | SharePoint/OneDrive (`32xx`) | Teams (`33xx`) |
|---|---|---|---|---|
| Provision (`P`) | Implemented | Implemented | Implemented | Implemented |
| Modify (`M`) | Implemented | Implemented | Implemented | Implemented |
| InventoryAndReport (`IR`) | Implemented | Implemented | Implemented | Implemented |

## Online Documentation

- [Online Provision README](./Provision/README.md)
- [Online Provision Detailed Catalog](./Provision/README-Provision-Catalog.md)
- [Online Modify README](./Modify/README.md)
- [Online Modify Detailed Catalog](./Modify/README-Modify-Catalog.md)
- [Online InventoryAndReport README](./InventoryAndReport/README.md)
- [Online InventoryAndReport Detailed Catalog](./InventoryAndReport/README-InventoryAndReport-Catalog.md)
- [Entra User Field Contract](./README-Entra-User-Field-Contract.md)
- [Root README](../README.md)
