# Discover Folder

`Discover` is for read-only inventory and reporting scripts.

Operational label: **Discover**.

Current status: inventory/report coverage includes Entra, Exchange Online, OneDrive, SharePoint, and Teams with dual discovery scope support (`-InputCsvPath` or `-DiscoverAll`), including Entra privileged-role inventory plus flattened Microsoft 365 group owner/member reporting, Exchange accepted-domain verification records plus mailbox analytics/permission consolidations/retention-tag tests and user/shared-mailbox SMTP-address reporting, OneDrive sharing/lock/external-principal reporting, and Teams channel-member reporting. A user-centric entitlement reconstruction series (0500 sequence) is implemented across MEID, EXOL, SPOL, and TEAM — these scripts take a user list as input and produce explainable access footprints with path attribution.

## Purpose

Use this folder for:
- State baselines before/after provision or modify runs
- Compliance and audit exports
- Environment inventory snapshots by workload

Do not use this folder for:
- Creation, modification, or deletion operations

## Naming Standard

- Script: `D-<WW>-<NNNN>-<Action>-<Target>.ps1`
- Input CSV: shared `Scope-*.input.csv` (preferred) or script-specific `D-<WW>-<NNNN>-<Action>-<Target>.input.csv` when needed
- Output pattern: `Results_D-<WW>-<NNNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Transcript log pattern: `Transcript_D-<WW>-<NNNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`
- Default output directory (unless `-OutputCsvPath` is supplied): `./TenantShift/Online/Discover/Discover_OutputCsvPath/`

Workload code allocation (`WW` in `D-<WW>-<NNNN>`):
- `MEID`: Entra
- `EXOL`: Exchange Online
- `ONDR`: OneDrive
- `SPOL`: SharePoint
- `TEAM`: Teams

Example:
- `D-MEID-0010-Get-EntraUsers.ps1`
- `Scope-Users.input.csv`

## Run Pattern

Run from repository root:

```powershell
pwsh ./TenantShift/Online/Discover/D-MEID-0010-Get-EntraUsers.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv
pwsh ./TenantShift/Online/Discover/D-ONDR-0010-Get-OneDriveProvisioningStatus.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-ONDR-0020-Get-OneDriveStorageAndQuota.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-ONDR-0030-Get-OneDriveSiteCollectionAdmins.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-ONDR-0040-Get-OneDriveSharingSettings.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-ONDR-0050-Get-OneDriveExternalSharingLinks.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-ONDR-0060-Get-OneDriveSiteLockState.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-EXOL-0030-Get-ExchangeOnlineMailContacts.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-MailContacts.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0020-Get-EntraGuestUsers.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-GuestUsers.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0060-Get-EntraPrivilegedRoles.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-EntraPrivilegedRoles.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0080-Get-EntraDynamicUserSecurityGroups.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-EntraDynamicUserSecurityGroups.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0120-Get-EntraMicrosoft365GroupOwners.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-M365Groups.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0110-Get-EntraMicrosoft365GroupMembers.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-M365Groups.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0080-Get-ExchangeOnlineResourceMailboxes.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-ResourceMailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0130-Get-ExchangeOnlineMailboxDelegations.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Mailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0010-Get-ExchangeOnlineDomainVerificationRecords.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-AcceptedDomains.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0210-Get-ExchangeOnlineRecipientTypeCounts.ps1 -InputCsvPath ./TenantShift/Online/Discover/D-EXOL-0210-Get-ExchangeOnlineRecipientTypeCounts.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0220-Get-ExchangeOnlineMailboxHighLevelStats.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Mailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0230-Get-ExchangeOnlineMailboxSizes.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Mailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0240-Get-ExchangeOnlineMailboxStatsPerMailbox.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Mailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0250-Get-ExchangeOnlineMailboxPermissionsConsolidated.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Mailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0280-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.ps1 -InputCsvPath ./TenantShift/Online/Discover/D-EXOL-0280-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0260-Get-ExchangeOnlineUserMailboxSmtpAddresses.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Mailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0270-Get-ExchangeOnlineSharedMailboxSmtpAddresses.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-SharedMailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0100-Get-ExchangeOnlineMailEnabledSecurityGroupMembers.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-MailEnabledSecurityGroups.input.csv
pwsh ./TenantShift/Online/Discover/D-TEAM-0010-Get-MicrosoftTeams.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Teams.input.csv
pwsh ./TenantShift/Online/Discover/D-SPOL-0010-Get-SharePointSites.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-SharePointSites.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-TEAM-0040-Get-MicrosoftTeamChannelMembers.ps1 -InputCsvPath ./TenantShift/Online/Discover/D-TEAM-0040-Get-MicrosoftTeamChannelMembers.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0500-Get-EntraAssignedGroupsByUser.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0510-Get-EntraDynamicGroupsByUser.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0520-Get-EntraM365GroupsByUser.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0500-Get-ExchangeOnlineMailboxAccessByUser.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv
pwsh ./TenantShift/Online/Discover/D-SPOL-0500-Get-SharePointSitesByUser.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-TEAM-0500-Get-MicrosoftTeamsByUser.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv
pwsh ./TenantShift/Online/Discover/D-TEAM-0510-Get-MicrosoftTeamChannelsByUser.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv
```

Discover-all pattern examples:

```powershell
pwsh ./TenantShift/Online/Discover/D-MEID-0010-Get-EntraUsers.ps1 -DiscoverAll
pwsh ./TenantShift/Online/Discover/D-MEID-0060-Get-EntraPrivilegedRoles.ps1 -DiscoverAll
pwsh ./TenantShift/Online/Discover/D-MEID-0120-Get-EntraMicrosoft365GroupOwners.ps1 -DiscoverAll
pwsh ./TenantShift/Online/Discover/D-MEID-0110-Get-EntraMicrosoft365GroupMembers.ps1 -DiscoverAll
pwsh ./TenantShift/Online/Discover/D-EXOL-0070-Get-ExchangeOnlineSharedMailboxes.ps1 -DiscoverAll
pwsh ./TenantShift/Online/Discover/D-EXOL-0010-Get-ExchangeOnlineDomainVerificationRecords.ps1 -DiscoverAll
pwsh ./TenantShift/Online/Discover/D-EXOL-0210-Get-ExchangeOnlineRecipientTypeCounts.ps1 -DiscoverAll
pwsh ./TenantShift/Online/Discover/D-EXOL-0230-Get-ExchangeOnlineMailboxSizes.ps1 -DiscoverAll
pwsh ./TenantShift/Online/Discover/D-EXOL-0280-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.ps1 -DiscoverAll
pwsh ./TenantShift/Online/Discover/D-EXOL-0260-Get-ExchangeOnlineUserMailboxSmtpAddresses.ps1 -DiscoverAll
pwsh ./TenantShift/Online/Discover/D-EXOL-0270-Get-ExchangeOnlineSharedMailboxSmtpAddresses.ps1 -DiscoverAll
pwsh ./TenantShift/Online/Discover/D-EXOL-0100-Get-ExchangeOnlineMailEnabledSecurityGroupMembers.ps1 -DiscoverAll
pwsh ./TenantShift/Online/Discover/D-SPOL-0010-Get-SharePointSites.ps1 -DiscoverAll -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-TEAM-0040-Get-MicrosoftTeamChannelMembers.ps1 -DiscoverAll
```

## Discovery Scope Modes

- CSV-bounded mode: `-InputCsvPath` with validated required headers.
- Unbounded mode: `-DiscoverAll` (internally maps required scope headers to `*`).
- Both modes export a `ScopeMode` column in the results (`Csv` or `DiscoverAll`).

For copy/paste command building, use `./TenantShift/Online/Discover/Discover-Orchestrator.xlsx`.

## Scoped Input Files

Use shared scope CSVs when multiple scripts use the same object key:

| Scope File | Key Column(s) | Reused By |
|---|---|---|
| `Scope-Users.input.csv` | `UserPrincipalName` | `D-MEID-0010`, `D-MEID-0030`, `D-ONDR-0010`, `D-ONDR-0020`, `D-ONDR-0030`, `D-ONDR-0040`, `D-ONDR-0050`, `D-ONDR-0060`, `D-MEID-0500`, `D-MEID-0510`, `D-MEID-0520`, `D-EXOL-0500`, `D-SPOL-0500`, `D-TEAM-0500`, `D-TEAM-0510` |
| `Scope-GuestUsers.input.csv` | `UserPrincipalName` | `D-MEID-0020` |
| `Scope-EntraPrivilegedRoles.input.csv` | `RoleDisplayName` | `D-MEID-0060` |
| `Scope-EntraSecurityGroups.input.csv` | `GroupDisplayName` | `D-MEID-0070`, `D-MEID-0090` |
| `Scope-EntraDynamicUserSecurityGroups.input.csv` | `GroupDisplayName` | `D-MEID-0080` |
| `Scope-M365Groups.input.csv` | `GroupMailNickname` | `D-MEID-0100`, `D-MEID-0110`, `D-MEID-0120` |
| `Scope-Teams.input.csv` | `TeamMailNickname` | `D-TEAM-0010`, `D-TEAM-0020`, `D-TEAM-0030` |
| `D-TEAM-0040-Get-MicrosoftTeamChannelMembers.input.csv` | `TeamMailNickname`, `ChannelDisplayName` | `D-TEAM-0040` |
| `Scope-MailContacts.input.csv` | `MailContactIdentity` | `D-EXOL-0030` |
| `Scope-DistributionLists.input.csv` | `DistributionGroupIdentity` | `D-EXOL-0040`, `D-EXOL-0090` |
| `Scope-MailEnabledSecurityGroups.input.csv` | `SecurityGroupIdentity` | `D-EXOL-0050`, `D-EXOL-0100` |
| `Scope-DynamicDistributionGroups.input.csv` | `DynamicDistributionGroupIdentity` | `D-EXOL-0060` |
| `Scope-AcceptedDomains.input.csv` | `DomainName` | `D-EXOL-0010` |
| `D-EXOL-0210-Get-ExchangeOnlineRecipientTypeCounts.input.csv` | `RecipientIdentity` | `D-EXOL-0210` |
| `Scope-SharedMailboxes.input.csv` | `SharedMailboxIdentity` | `D-EXOL-0070`, `D-EXOL-0110`, `D-EXOL-0270` |
| `Scope-ResourceMailboxes.input.csv` | `ResourceMailboxIdentity` | `D-EXOL-0080`, `D-EXOL-0120` |
| `Scope-Mailboxes.input.csv` | `MailboxIdentity` | `D-EXOL-0130`, `D-EXOL-0140`, `D-EXOL-0220`, `D-EXOL-0230`, `D-EXOL-0240`, `D-EXOL-0250`, `D-EXOL-0260` |
| `D-EXOL-0280-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.input.csv` | `MailboxIdentity`, `ExpectedTagNames` | `D-EXOL-0280` |
| `Scope-SharePointSites.input.csv` | `SiteUrl` | `D-SPOL-0010` |

Use `*` in scope files to inventory all objects for that key type (for example `Scope-GuestUsers.input.csv` defaults to `*` because guest UPN formats vary by tenant).

## Sample Build Set Verification

The scope CSVs are preloaded with sample objects from `TenantShift/Online/Provision/*.input.csv`, so you can:

1. Run the Provision scripts with the provided sample set.
2. Run the Discover scripts with the matching `Scope-*.input.csv` files to verify created objects and relationships quickly.

Example verification run set:

```powershell
pwsh ./TenantShift/Online/Discover/D-MEID-0010-Get-EntraUsers.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0020-Get-EntraGuestUsers.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-GuestUsers.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0030-Get-EntraUserLicenses.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0060-Get-EntraPrivilegedRoles.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-EntraPrivilegedRoles.input.csv
pwsh ./TenantShift/Online/Discover/D-ONDR-0010-Get-OneDriveProvisioningStatus.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-ONDR-0020-Get-OneDriveStorageAndQuota.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-ONDR-0030-Get-OneDriveSiteCollectionAdmins.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-ONDR-0040-Get-OneDriveSharingSettings.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-ONDR-0050-Get-OneDriveExternalSharingLinks.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-ONDR-0060-Get-OneDriveSiteLockState.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Users.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-MEID-0070-Get-EntraSecurityGroups.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-EntraSecurityGroups.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0080-Get-EntraDynamicUserSecurityGroups.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-EntraDynamicUserSecurityGroups.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0090-Get-EntraSecurityGroupMembers.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-EntraSecurityGroups.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0100-Get-EntraMicrosoft365Groups.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-M365Groups.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0120-Get-EntraMicrosoft365GroupOwners.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-M365Groups.input.csv
pwsh ./TenantShift/Online/Discover/D-MEID-0110-Get-EntraMicrosoft365GroupMembers.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-M365Groups.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0030-Get-ExchangeOnlineMailContacts.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-MailContacts.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0040-Get-ExchangeOnlineDistributionLists.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-DistributionLists.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0090-Get-ExchangeOnlineDistributionListMembers.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-DistributionLists.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0070-Get-ExchangeOnlineSharedMailboxes.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-SharedMailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0110-Get-ExchangeOnlineSharedMailboxPermissions.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-SharedMailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0270-Get-ExchangeOnlineSharedMailboxSmtpAddresses.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-SharedMailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0080-Get-ExchangeOnlineResourceMailboxes.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-ResourceMailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0120-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-ResourceMailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0130-Get-ExchangeOnlineMailboxDelegations.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Mailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0140-Get-ExchangeOnlineMailboxFolderPermissions.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Mailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0050-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-MailEnabledSecurityGroups.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0100-Get-ExchangeOnlineMailEnabledSecurityGroupMembers.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-MailEnabledSecurityGroups.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0060-Get-ExchangeOnlineDynamicDistributionGroups.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-DynamicDistributionGroups.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0010-Get-ExchangeOnlineDomainVerificationRecords.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-AcceptedDomains.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0210-Get-ExchangeOnlineRecipientTypeCounts.ps1 -InputCsvPath ./TenantShift/Online/Discover/D-EXOL-0210-Get-ExchangeOnlineRecipientTypeCounts.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0220-Get-ExchangeOnlineMailboxHighLevelStats.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Mailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0230-Get-ExchangeOnlineMailboxSizes.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Mailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0240-Get-ExchangeOnlineMailboxStatsPerMailbox.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Mailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0250-Get-ExchangeOnlineMailboxPermissionsConsolidated.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Mailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0280-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.ps1 -InputCsvPath ./TenantShift/Online/Discover/D-EXOL-0280-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.input.csv
pwsh ./TenantShift/Online/Discover/D-EXOL-0260-Get-ExchangeOnlineUserMailboxSmtpAddresses.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Mailboxes.input.csv
pwsh ./TenantShift/Online/Discover/D-SPOL-0010-Get-SharePointSites.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-SharePointSites.input.csv -SharePointAdminUrl https://contoso-admin.sharepoint.com
pwsh ./TenantShift/Online/Discover/D-TEAM-0010-Get-MicrosoftTeams.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Teams.input.csv
pwsh ./TenantShift/Online/Discover/D-TEAM-0020-Get-MicrosoftTeamMembers.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Teams.input.csv
pwsh ./TenantShift/Online/Discover/D-TEAM-0030-Get-MicrosoftTeamChannels.ps1 -InputCsvPath ./TenantShift/Online/Discover/Scope-Teams.input.csv
pwsh ./TenantShift/Online/Discover/D-TEAM-0040-Get-MicrosoftTeamChannelMembers.ps1 -InputCsvPath ./TenantShift/Online/Discover/D-TEAM-0040-Get-MicrosoftTeamChannelMembers.input.csv
```

## Discovery Scope Modes

- CSV-bounded mode: `-InputCsvPath` with validated required headers.
- Unbounded mode: `-DiscoverAll` (internally maps required scope headers to `*`).
- Both modes export a `ScopeMode` column in the results (`Csv` or `DiscoverAll`).

For copy/paste command building, use `./TenantShift/Online/Discover/Discover-Orchestrator.xlsx`.

## Discover Output Standard

Discovery scripts should export consistent, easy-to-diff output:

- Primary object key columns (for example: `UserPrincipalName`, `GroupId`)
- Workload/object metadata columns
- `Status` and `Message` columns for per-record operation logging
- Timestamped output file names

## Discover Catalog

| ID | Script | Workload | Purpose | Status |
|---|---|---|---|---|
| D-MEID-0010 | `D-MEID-0010-Get-EntraUsers.ps1` | Entra | Export tenant users with expanded profile/contact/org/extension fields. | Implemented |
| D-MEID-0020 | `D-MEID-0020-Get-EntraGuestUsers.ps1` | Entra | Export guest users. | Implemented |
| D-MEID-0030 | `D-MEID-0030-Get-EntraUserLicenses.ps1` | Entra | Export assigned licenses. | Implemented |
| D-MEID-0060 | `D-MEID-0060-Get-EntraPrivilegedRoles.ps1` | Entra | Export activated directory roles and their assigned members. | Implemented |
| D-ONDR-0010 | `D-ONDR-0010-Get-OneDriveProvisioningStatus.ps1` | OneDrive | Report OneDrive URL and provisioning status by user. | Implemented |
| D-ONDR-0020 | `D-ONDR-0020-Get-OneDriveStorageAndQuota.ps1` | OneDrive | Report OneDrive storage usage and quota settings by user. | Implemented |
| D-ONDR-0030 | `D-ONDR-0030-Get-OneDriveSiteCollectionAdmins.ps1` | OneDrive | Report OneDrive site collection admins by user. | Implemented |
| D-ONDR-0040 | `D-ONDR-0040-Get-OneDriveSharingSettings.ps1` | OneDrive | Report OneDrive sharing policy posture by user/site. | Implemented |
| D-ONDR-0050 | `D-ONDR-0050-Get-OneDriveExternalSharingLinks.ps1` | OneDrive | Report OneDrive external sharing principals/access by user/site. | Implemented |
| D-ONDR-0060 | `D-ONDR-0060-Get-OneDriveSiteLockState.ps1` | OneDrive | Report OneDrive lock state and access freeze posture. | Implemented |
| D-MEID-0070 | `D-MEID-0070-Get-EntraSecurityGroups.ps1` | Entra | Export assigned security groups. | Implemented |
| D-MEID-0080 | `D-MEID-0080-Get-EntraDynamicUserSecurityGroups.ps1` | Entra | Export dynamic user groups and rules. | Implemented |
| D-MEID-0090 | `D-MEID-0090-Get-EntraSecurityGroupMembers.ps1` | Entra | Export security group membership. | Implemented |
| D-MEID-0100 | `D-MEID-0100-Get-EntraMicrosoft365Groups.ps1` | Entra | Export Microsoft 365 groups. | Implemented |
| D-MEID-0120 | `D-MEID-0120-Get-EntraMicrosoft365GroupOwners.ps1` | Entra | Export Microsoft 365 group owners as one row per owner. | Implemented |
| D-MEID-0110 | `D-MEID-0110-Get-EntraMicrosoft365GroupMembers.ps1` | Entra | Export Microsoft 365 group members as one row per member. | Implemented |
| D-TEAM-0010 | `D-TEAM-0010-Get-MicrosoftTeams.ps1` | Teams | Export Teams and core settings. | Implemented |
| D-TEAM-0020 | `D-TEAM-0020-Get-MicrosoftTeamMembers.ps1` | Teams | Export Team membership. | Implemented |
| D-TEAM-0030 | `D-TEAM-0030-Get-MicrosoftTeamChannels.ps1` | Teams | Export channels by Team. | Implemented |
| D-TEAM-0040 | `D-TEAM-0040-Get-MicrosoftTeamChannelMembers.ps1` | Teams | Export private/shared channel membership. | Implemented |
| D-EXOL-0030 | `D-EXOL-0030-Get-ExchangeOnlineMailContacts.ps1` | Exchange Online | Export mail contacts. | Implemented |
| D-EXOL-0040 | `D-EXOL-0040-Get-ExchangeOnlineDistributionLists.ps1` | Exchange Online | Export distribution lists. | Implemented |
| D-EXOL-0090 | `D-EXOL-0090-Get-ExchangeOnlineDistributionListMembers.ps1` | Exchange Online | Export DL membership. | Implemented |
| D-EXOL-0070 | `D-EXOL-0070-Get-ExchangeOnlineSharedMailboxes.ps1` | Exchange Online | Export shared mailboxes including semicolon-delimited proxy/email addresses. | Implemented |
| D-EXOL-0110 | `D-EXOL-0110-Get-ExchangeOnlineSharedMailboxPermissions.ps1` | Exchange Online | Export mailbox permissions. | Implemented |
| D-EXOL-0080 | `D-EXOL-0080-Get-ExchangeOnlineResourceMailboxes.ps1` | Exchange Online | Export room/equipment mailboxes. | Implemented |
| D-EXOL-0120 | `D-EXOL-0120-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1` | Exchange Online | Export booking delegate settings. | Implemented |
| D-EXOL-0130 | `D-EXOL-0130-Get-ExchangeOnlineMailboxDelegations.ps1` | Exchange Online | Export mailbox delegations. | Implemented |
| D-EXOL-0140 | `D-EXOL-0140-Get-ExchangeOnlineMailboxFolderPermissions.ps1` | Exchange Online | Export folder-level permissions. | Implemented |
| D-EXOL-0050 | `D-EXOL-0050-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1` | Exchange Online | Export mail-enabled security groups with semicolon-delimited member summary. | Implemented |
| D-EXOL-0100 | `D-EXOL-0100-Get-ExchangeOnlineMailEnabledSecurityGroupMembers.ps1` | Exchange Online | Export one-row-per-member mail-enabled security group membership. | Implemented |
| D-EXOL-0060 | `D-EXOL-0060-Get-ExchangeOnlineDynamicDistributionGroups.ps1` | Exchange Online | Export dynamic distribution groups. | Implemented |
| D-EXOL-0010 | `D-EXOL-0010-Get-ExchangeOnlineDomainVerificationRecords.ps1` | Exchange Online | Export accepted-domain verification record requirements and tenant-domain verification state. | Implemented |
| D-EXOL-0210 | `D-EXOL-0210-Get-ExchangeOnlineRecipientTypeCounts.ps1` | Exchange Online | Export recipient counts grouped by RecipientTypeDetails for scoped/all recipients. | Implemented |
| D-EXOL-0220 | `D-EXOL-0220-Get-ExchangeOnlineMailboxHighLevelStats.ps1` | Exchange Online | Export high-level mailbox size/item distribution statistics for scoped/all mailboxes. | Implemented |
| D-EXOL-0230 | `D-EXOL-0230-Get-ExchangeOnlineMailboxSizes.ps1` | Exchange Online | Export per-mailbox main/archive size and quota summary. | Implemented |
| D-EXOL-0240 | `D-EXOL-0240-Get-ExchangeOnlineMailboxStatsPerMailbox.ps1` | Exchange Online | Export detailed per-mailbox usage/activity statistics. | Implemented |
| D-EXOL-0250 | `D-EXOL-0250-Get-ExchangeOnlineMailboxPermissionsConsolidated.ps1` | Exchange Online | Export one-row-per-mailbox consolidated delegated permission summary. | Implemented |
| D-EXOL-0280 | `D-EXOL-0280-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.ps1` | Exchange Online | Detect unexpected/missing mailbox retention policy tags against expected baselines. | Implemented |
| D-EXOL-0260 | `D-EXOL-0260-Get-ExchangeOnlineUserMailboxSmtpAddresses.ps1` | Exchange Online | Export one-row-per-SMTP-address user mailbox inventory for primary and alias addresses. | Implemented |
| D-EXOL-0270 | `D-EXOL-0270-Get-ExchangeOnlineSharedMailboxSmtpAddresses.ps1` | Exchange Online | Export one-row-per-SMTP-address shared mailbox inventory for primary and alias addresses. | Implemented |
| D-SPOL-0010 | `D-SPOL-0010-Get-SharePointSites.ps1` | SharePoint | Export SharePoint site inventory and core metadata. | Implemented |
| D-MEID-0500 | `D-MEID-0500-Get-EntraAssignedGroupsByUser.ps1` | Entra | Assigned security group memberships per user — direct and transitive with full path attribution. | Implemented |
| D-MEID-0510 | `D-MEID-0510-Get-EntraDynamicGroupsByUser.ps1` | Entra | Dynamic security group evaluated memberships per user. | Implemented |
| D-MEID-0520 | `D-MEID-0520-Get-EntraM365GroupsByUser.ps1` | Entra | Microsoft 365 group member and owner relationships per user. | Implemented |
| D-EXOL-0500 | `D-EXOL-0500-Get-ExchangeOnlineMailboxAccessByUser.ps1` | Exchange Online | Effective mailbox access per user across shared, resource, and equipment mailboxes with direct/group attribution. | Implemented |
| D-SPOL-0500 | `D-SPOL-0500-Get-SharePointSitesByUser.ps1` | SharePoint | SharePoint site access per user across all access paths with assignment chain. | Implemented |
| D-TEAM-0500 | `D-TEAM-0500-Get-MicrosoftTeamsByUser.ps1` | Teams | Teams membership and ownership per user with role and access type. | Implemented |
| D-TEAM-0510 | `D-TEAM-0510-Get-MicrosoftTeamChannelsByUser.ps1` | Teams | Teams channel access per user by channel type with cross-tenant detection. | Implemented |

## Discover Standards

- Keep scripts read-only.
- Support either `-InputCsvPath` or `-DiscoverAll` for every `D` script.
- Prefer shared scope CSV files where key columns overlap across scripts.
- Use script-specific `D-<WW>-<NNNN>-...input.csv` templates only when a script requires unique scope shape not covered by shared scope files.
- Keep workload explicit in script names.
- Reuse `./TenantShift/Common/Online/M365.Common.psm1` (repository-root path) where common validation and result formatting helps.
- Prefer deterministic column ordering for easier diffing between snapshots.

## References

- [Root README](../../README.md)
- [Provision README](../Provision/README.md)
- [Discover Detailed Catalog](./README-Discover-Catalog.md)
- [Entra User Field Contract](../README-Entra-User-Field-Contract.md)
- [Operator Runbook](./RUNBOOK-Discover.md)
