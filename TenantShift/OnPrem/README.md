# OnPrem Folder

`OnPrem` is the planning and implementation area for on-premises workloads.

Current status: planning matrix is defined for ActiveDirectory (`00xx`), GroupPolicy (`01xx`), ExchangeOnPrem (`02xx`), and FileServices (`03xx`). ActiveDirectory script set (`P0001`, `P0002`, `P0005`, `P0006`, `P0009`, `M0001`, `M0002`, `M0005`, `M0006`, `M0007`, `M0008`, `M0009`, `M0010`, `D0001`, `D0002`, `D0005`, `D0006`, `D0007`, `D0008`, `D0009`, `D0010`, `D0011`, `D0012`) is implemented. ExchangeOnPrem script set (`P0213`, `P0214`, `P0215`, `P0216`, `P0218`, `P0219`, `M0213`, `M0214`, `M0215`, `M0216`, `M0217`, `M0218`, `M0219`, `M0220`, `M0221`, `M0222`, `M0223`, `M0224`, `M0225`, `M0226`, `D0213`, `D0214`, `D0215`, `D0216`, `D0217`, `D0218`, `D0219`, `D0220`, `D0221`, `D0222`, `D0223`, `D0224`, `D0225`, `D0226`) is implemented; remaining GroupPolicy and FileServices scripts are planned.

## Planned Workloads

- `00xx`: ActiveDirectory
- `01xx`: GroupPolicy
- `02xx`: ExchangeOnPrem
- `03xx`: FileServices

## Runtime Assumptions

- ActiveDirectory scripts (`00xx`) run natively in Windows PowerShell `5.1`.
- ExchangeOnPrem scripts (`02xx`) run natively in Exchange Management Shell (Windows PowerShell `5.1`).

## Operation Folders

- `TenantShift/OnPrem/Provision/`
- `TenantShift/OnPrem/Modify/`
- `TenantShift/OnPrem/Discover/`

## Matrix Overview

| Operation | ActiveDirectory (`00xx`) | GroupPolicy (`01xx`) | ExchangeOnPrem (`02xx`) | FileServices (`03xx`) |
|---|---|---|---|---|
| Provision (`P`) | Implemented (`P0001`, `P0002`, `P0005`, `P0006`, `P0009`) | Planned | Implemented (`P0213`, `P0214`, `P0215`, `P0216`, `P0218`, `P0219`) | Planned (scope to be refined) |
| Modify (`M`) | Implemented (`M0001`, `M0002`, `M0005`, `M0006`, `M0007`, `M0008`, `M0009`, `M0010`) | Planned | Implemented (`M0213`, `M0214`, `M0215`, `M0216`, `M0217`, `M0218`, `M0219`, `M0220`, `M0221`, `M0222`, `M0223`, `M0224`, `M0225`, `M0226`) | Planned (scope to be refined) |
| Discover (`D`) | Implemented (`D0001`, `D0002`, `D0005`, `D0006`, `D0007`, `D0008`, `D0009`, `D0010`, `D0011`, `D0012`) | Planned | Implemented (`D0213`, `D0214`, `D0215`, `D0216`, `D0217`, `D0218`, `D0219`, `D0220`, `D0221`, `D0222`, `D0223`, `D0224`, `D0225`, `D0226`) | Planned (scope to be refined) |

## OnPrem Documentation

- [OnPrem Provision README](./Provision/README.md)
- [OnPrem Provision Detailed Catalog](./Provision/README-Provision-Catalog.md)
- [OnPrem Modify README](./Modify/README.md)
- [OnPrem Modify Detailed Catalog](./Modify/README-Modify-Catalog.md)
- [OnPrem Discover README](./Discover/README.md)
- [OnPrem Discover Detailed Catalog](./Discover/README-Discover-Catalog.md)
- [Root README](../../README.md)
