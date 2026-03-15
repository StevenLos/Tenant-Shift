# SharedModule Folder

`SharedModule` contains the production script sets that reuse shared runtime modules.

## Folder Layout

- `SharedModule/Online/`: online workload scripts grouped by `Provision`, `Modify`, and `InventoryAndReport`
- `SharedModule/OnPrem/`: on-prem workload scripts grouped by `Provision`, `Modify`, and `InventoryAndReport`
- `SharedModule/Common/`: reusable modules used by `SharedModule` scripts

## Usage Notes

- Shared-module script names follow `SM-<P|M|IR><WWNN>-<Action>-<Target>.ps1`.
- Each operation script has a matching `.input.csv` file in the same folder.
- Run scripts from the repository root so relative paths resolve correctly.

## References

- [Online Provision README](./Online/Provision/README.md)
- [Online Modify README](./Online/Modify/README.md)
- [Online InventoryAndReport README](./Online/InventoryAndReport/README.md)
- [OnPrem README](./OnPrem/README.md)
- [Common README](./Common/README.md)
- [Entra User Field Contract](./Online/README-Entra-User-Field-Contract.md)
- [Standalone README](../Standalone/README.md)
