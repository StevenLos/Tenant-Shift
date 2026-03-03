# OnPrem Folder

`OnPrem` is the planning and implementation area for on-premises workloads.

Current status: planning matrix is defined for ActiveDirectory (`00xx`), GroupPolicy (`01xx`), ExchangeOnPrem (`02xx`), and FileServices (`03xx`). ActiveDirectory script set (`P0001`, `P0002`, `P0005`, `P0006`, `P0009`, `M0001`, `M0002`, `M0005`, `M0006`, `M0007`, `M0008`, `M0009`, `IR0001`, `IR0002`, `IR0005`, `IR0006`, `IR0007`, `IR0008`, `IR0009`) is implemented. ExchangeOnPrem script set (`P0213`, `P0214`, `P0215`, `P0216`, `P0218`, `P0219`, `M0213`, `M0214`, `M0215`, `M0216`, `M0217`, `M0218`, `M0219`, `M0220`, `M0221`, `M0222`, `M0223`, `M0224`, `IR0213`, `IR0214`, `IR0215`, `IR0216`, `IR0217`, `IR0218`, `IR0219`, `IR0220`, `IR0221`, `IR0222`, `IR0223`) is implemented; remaining GroupPolicy and FileServices scripts are planned.

## Planned Workloads

- `00xx`: ActiveDirectory
- `01xx`: GroupPolicy
- `02xx`: ExchangeOnPrem
- `03xx`: FileServices

## Runtime Assumptions

- ActiveDirectory scripts (`00xx`) run natively in Windows PowerShell `5.1`.
- ExchangeOnPrem scripts (`02xx`) run natively in Exchange Management Shell (Windows PowerShell `5.1`).

## Operation Folders

- `OnPrem/Provision/`
- `OnPrem/Modify/`
- `OnPrem/InventoryAndReport/`

## Matrix Overview

| Operation | ActiveDirectory (`00xx`) | GroupPolicy (`01xx`) | ExchangeOnPrem (`02xx`) | FileServices (`03xx`) |
|---|---|---|---|---|
| Provision (`P`) | Implemented (`P0001`, `P0002`, `P0005`, `P0006`, `P0009`) | Planned | Implemented (`P0213`, `P0214`, `P0215`, `P0216`, `P0218`, `P0219`) | Planned (scope to be refined) |
| Modify (`M`) | Implemented (`M0001`, `M0002`, `M0005`, `M0006`, `M0007`, `M0008`, `M0009`) | Planned | Implemented (`M0213`, `M0214`, `M0215`, `M0216`, `M0217`, `M0218`, `M0219`, `M0220`, `M0221`, `M0222`, `M0223`, `M0224`) | Planned (scope to be refined) |
| InventoryAndReport (`IR`) | Implemented (`IR0001`, `IR0002`, `IR0005`, `IR0006`, `IR0007`, `IR0008`, `IR0009`) | Planned | Implemented (`IR0213`, `IR0214`, `IR0215`, `IR0216`, `IR0217`, `IR0218`, `IR0219`, `IR0220`, `IR0221`, `IR0222`, `IR0223`) | Planned (scope to be refined) |

## OnPrem Documentation

- [OnPrem Provision README](./Provision/README.md)
- [OnPrem Provision Detailed Catalog](./Provision/README-Provision-Catalog.md)
- [OnPrem Modify README](./Modify/README.md)
- [OnPrem Modify Detailed Catalog](./Modify/README-Modify-Catalog.md)
- [OnPrem InventoryAndReport README](./InventoryAndReport/README.md)
- [OnPrem InventoryAndReport Detailed Catalog](./InventoryAndReport/README-InventoryAndReport-Catalog.md)
- [Root README](../README.md)
