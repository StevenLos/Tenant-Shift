# Execution Roadmap (Scope Locked)

This roadmap tracks execution of approved proposals with these exclusions:

- Excluded for now: GroupPolicy (`01xx`), FileServices (`03xx`).

## Wave 1: Quality + Parity Foundation (Completed)

- Added online membership remove semantics:
  - `M3007-Set-EntraSecurityGroupMembers.ps1` (`MemberAction`: `Add`/`Remove`)
  - `M3115-Set-ExchangeOnlineDistributionListMembers.ps1` (`MemberAction`: `Add`/`Remove`)
- Updated sample templates to include `MemberAction` and remove examples:
  - `Online/Modify/M3007-Set-EntraSecurityGroupMembers.input.csv`
  - `Online/Modify/M3115-Set-ExchangeOnlineDistributionListMembers.input.csv`
- Removed duplicate SharePoint connection helper implementation from:
  - `Online/Modify/M3204-PreProvision-OneDrive.ps1`
  - Script now reuses `Common/Online/M365.Common.psm1` helper.
- Added contract/quality automation:
  - `Build/Test-RepositoryContracts.ps1`
  - `Tests/RepositoryContracts.Tests.ps1`
  - `Tests/MembershipTemplateContracts.Tests.ps1`
  - `.github/workflows/powershell-quality.yml`

## Wave 2: Online Planned Script Set (Completed)

- Modify:
  - `M3207-Set-OneDriveSharingSettings.ps1`
  - `M3208-Revoke-OneDriveExternalSharingLinks.ps1`
  - `M3209-Set-OneDriveSiteLockState.ps1`
  - `M3309-Update-MicrosoftTeams.ps1`
- Inventory and Report:
  - `IR3207-Get-OneDriveSharingSettings.ps1`
  - `IR3208-Get-OneDriveExternalSharingLinks.ps1`
  - `IR3209-Get-OneDriveSiteLockState.ps1`
  - `IR3312-Get-MicrosoftTeamChannelMembers.ps1`
- Added sample templates for new wave 2 scripts:
  - `Online/Modify/M3207-Set-OneDriveSharingSettings.input.csv`
  - `Online/Modify/M3208-Revoke-OneDriveExternalSharingLinks.input.csv`
  - `Online/Modify/M3209-Set-OneDriveSiteLockState.input.csv`
  - `Online/Modify/M3309-Update-MicrosoftTeams.input.csv`
  - `Online/InventoryAndReport/IR3312-Get-MicrosoftTeamChannelMembers.input.csv`

## Wave 3: OnPrem Exchange Remaining Set (Completed)

- Provision:
  - `P0216-Create-ExchangeOnPremSharedMailboxes.ps1`
  - `P0218-Create-ExchangeOnPremResourceMailboxes.ps1`
- Modify:
  - `M0215-Set-ExchangeOnPremDistributionListMembers.ps1`
  - `M0216-Update-ExchangeOnPremSharedMailboxes.ps1`
  - `M0217-Set-ExchangeOnPremSharedMailboxPermissions.ps1`
  - `M0218-Update-ExchangeOnPremResourceMailboxes.ps1`
  - `M0219-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1`
  - `M0220-Set-ExchangeOnPremMailboxDelegations.ps1`
  - `M0221-Set-ExchangeOnPremMailboxFolderPermissions.ps1`
- Inventory and Report:
  - `IR0215-Get-ExchangeOnPremDistributionListMembers.ps1`
  - `IR0216-Get-ExchangeOnPremSharedMailboxes.ps1`
  - `IR0217-Get-ExchangeOnPremSharedMailboxPermissions.ps1`
  - `IR0218-Get-ExchangeOnPremResourceMailboxes.ps1`
  - `IR0219-Get-ExchangeOnPremResourceMailboxBookingDelegates.ps1`
  - `IR0220-Get-ExchangeOnPremMailboxDelegations.ps1`
  - `IR0221-Get-ExchangeOnPremMailboxFolderPermissions.ps1`

## Wave 4: Online Discovery Scope Parity (Completed)

- Add dual scope mode to online IR scripts where practical:
  - CSV-bounded: `-InputCsvPath`
  - Unbounded: `-DiscoverAll` with optional scope controls.
- Align online documentation with the approved dual-scope discovery model.

## Wave 5: Sample Data and Contract Expansion (Pending)

- Expand online sample CSV coverage so object templates consistently reflect a richer company model.
- Continue contract alignment where user/group object families still have narrow sample coverage.
