# Build Folder

`Build` contains repository build and helper automation scripts.

## Current Contents

- `Build-OrchestratorWorkbooks.ps1`: regenerates Excel orchestrator workbooks for online `Provision`, `Modify`, and `InventoryAndReport`.
- `Test-RepositoryContracts.ps1`: validates repository-level script and CSV contracts (metadata, PowerShell declarations, required-header/template alignment).
- `README-Execution-Roadmap.md`: implementation roadmap for approved proposal scope (excluding GroupPolicy/FileServices).

Run from repository root:

```powershell
pwsh ./Build/Build-OrchestratorWorkbooks.ps1
```

Run repository contract validation from repository root:

```powershell
pwsh ./Build/Test-RepositoryContracts.ps1
```
