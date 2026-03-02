# 4-Digit ID Migration Draft

Draft only. No file/script renames are applied by this document.

## Scheme

- Format: `<Prefix><WW><NN>`
- Prefix: `P`, `M`, `IR`
- `WW`: workload code
- `NN`: sequence (`01-99`)

## Workload Code Allocation

- `00-29`: OnPrem
- `30-59`: Online
- `60-89`: Unallocated
- `90-99`: Shared/Hybrid/Reserved

## Current Mapping Used For This Draft

- `00`: OnPrem ActiveDirectory (from `5xx`)
- `01`: OnPrem GroupPolicy (planned workload codes now defined)
- `02`: OnPrem ExchangeOnPrem (from `6xx`)
- `03`: OnPrem FileServices (from `7xx`)
- `30`: Online Entra (from `1xx`)
- `31`: Online ExchangeOnline (from `2xx`)
- `32`: Online SharePoint/OneDrive (from `3xx`)
- `33`: Online Teams (from `4xx`)

## Old To New ID Table

| Old ID | New ID | Script | Status |
|---|---|---|---|
| IR101 | IR3001 | IR101-Get-EntraUsers.ps1 | Implemented |
| IR102 | IR3002 | IR102-Get-EntraGuestUsers.ps1 | Implemented |
| IR103 | IR3003 | IR103-Get-EntraUserLicenses.ps1 | Implemented |
| IR105 | IR3005 | IR105-Get-EntraSecurityGroups.ps1 | Implemented |
| IR106 | IR3006 | IR106-Get-EntraDynamicUserSecurityGroups.ps1 | Implemented |
| IR107 | IR3007 | IR107-Get-EntraSecurityGroupMembers.ps1 | Implemented |
| IR108 | IR3008 | IR108-Get-EntraMicrosoft365Groups.ps1 | Implemented |
| IR213 | IR3113 | IR213-Get-ExchangeOnlineMailContacts.ps1 | Implemented |
| IR214 | IR3114 | IR214-Get-ExchangeOnlineDistributionLists.ps1 | Implemented |
| IR215 | IR3115 | IR215-Get-ExchangeOnlineDistributionListMembers.ps1 | Implemented |
| IR216 | IR3116 | IR216-Get-ExchangeOnlineSharedMailboxes.ps1 | Implemented |
| IR217 | IR3117 | IR217-Get-ExchangeOnlineSharedMailboxPermissions.ps1 | Implemented |
| IR218 | IR3118 | IR218-Get-ExchangeOnlineResourceMailboxes.ps1 | Implemented |
| IR219 | IR3119 | IR219-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1 | Implemented |
| IR220 | IR3120 | IR220-Get-ExchangeOnlineMailboxDelegations.ps1 | Implemented |
| IR221 | IR3121 | IR221-Get-ExchangeOnlineMailboxFolderPermissions.ps1 | Implemented |
| IR222 | IR3122 | IR222-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1 | Implemented |
| IR223 | IR3123 | IR223-Get-ExchangeOnlineDynamicDistributionGroups.ps1 | Implemented |
| IR304 | IR3204 | IR304-Get-OneDriveProvisioningStatus.ps1 | Implemented |
| IR305 | IR3205 | IR305-Get-OneDriveStorageAndQuota.ps1 | Implemented |
| IR306 | IR3206 | IR306-Get-OneDriveSiteCollectionAdmins.ps1 | Implemented |
| IR307 | IR3207 | IR307-Get-OneDriveSharingSettings.ps1 | Planned |
| IR308 | IR3208 | IR308-Get-OneDriveExternalSharingLinks.ps1 | Planned |
| IR309 | IR3209 | IR309-Get-OneDriveSiteLockState.ps1 | Planned |
| IR340 | IR3240 | IR340-Get-SharePointSites.ps1 | Implemented |
| IR409 | IR3309 | IR409-Get-MicrosoftTeams.ps1 | Implemented |
| IR410 | IR3310 | IR410-Get-MicrosoftTeamMembers.ps1 | Implemented |
| IR411 | IR3311 | IR411-Get-MicrosoftTeamChannels.ps1 | Implemented |
| IR412 | IR3312 | IR412-Get-MicrosoftTeamChannelMembers.ps1 | Planned |
| IR501 | IR0001 | IR501-Get-ActiveDirectoryUsers.ps1 | Planned |
| IR502 | IR0002 | IR502-Get-ActiveDirectoryContacts.ps1 | Planned |
| IR505 | IR0005 | IR505-Get-ActiveDirectorySecurityGroups.ps1 | Planned |
| IR507 | IR0007 | IR507-Get-ActiveDirectorySecurityGroupMembers.ps1 | Planned |
| IR509 | IR0009 | IR509-Get-ActiveDirectoryOrganizationalUnits.ps1 | Planned |
| IR613 | IR0213 | IR613-Get-ExchangeOnPremMailContacts.ps1 | Planned |
| IR614 | IR0214 | IR614-Get-ExchangeOnPremDistributionLists.ps1 | Planned |
| IR615 | IR0215 | IR615-Get-ExchangeOnPremDistributionListMembers.ps1 | Planned |
| IR616 | IR0216 | IR616-Get-ExchangeOnPremSharedMailboxes.ps1 | Planned |
| IR617 | IR0217 | IR617-Get-ExchangeOnPremSharedMailboxPermissions.ps1 | Planned |
| IR618 | IR0218 | IR618-Get-ExchangeOnPremResourceMailboxes.ps1 | Planned |
| IR619 | IR0219 | IR619-Get-ExchangeOnPremResourceMailboxBookingDelegates.ps1 | Planned |
| IR620 | IR0220 | IR620-Get-ExchangeOnPremMailboxDelegations.ps1 | Planned |
| IR621 | IR0221 | IR621-Get-ExchangeOnPremMailboxFolderPermissions.ps1 | Planned |
| IR622 | IR0222 | IR622-Get-ExchangeOnPremMailEnabledSecurityGroups.ps1 | Planned |
| IR623 | IR0223 | IR623-Get-ExchangeOnPremDynamicDistributionGroups.ps1 | Planned |
| IR701 | IR0301 | IR701-Get-FileServicesShares.ps1 | Planned |
| IR702 | IR0302 | IR702-Get-FileServicesSharePermissions.ps1 | Planned |
| IR703 | IR0303 | IR703-Get-FileServicesNtfsPermissions.ps1 | Planned |
| IR704 | IR0304 | IR704-Get-FileServicesHomeDriveUsage.ps1 | Planned |
| M101 | M3001 | M101-Update-EntraUsers.ps1 | Planned |
| M102 | M3002 | M102-Set-EntraUserAccountState.ps1 | Planned |
| M103 | M3003 | M103-Set-EntraUserLicenses.ps1 | Implemented |
| M105 | M3005 | M105-Update-EntraAssignedSecurityGroups.ps1 | Planned |
| M106 | M3006 | M106-Update-EntraDynamicUserSecurityGroups.ps1 | Planned |
| M107 | M3007 | M107-Set-EntraSecurityGroupMembers.ps1 | Implemented |
| M108 | M3008 | M108-Update-EntraMicrosoft365Groups.ps1 | Planned |
| M213 | M3113 | M213-Update-ExchangeOnlineMailContacts.ps1 | Implemented |
| M214 | M3114 | M214-Update-ExchangeOnlineDistributionLists.ps1 | Implemented |
| M215 | M3115 | M215-Set-ExchangeOnlineDistributionListMembers.ps1 | Implemented |
| M216 | M3116 | M216-Update-ExchangeOnlineSharedMailboxes.ps1 | Implemented |
| M217 | M3117 | M217-Set-ExchangeOnlineSharedMailboxPermissions.ps1 | Implemented |
| M218 | M3118 | M218-Update-ExchangeOnlineResourceMailboxes.ps1 | Implemented |
| M219 | M3119 | M219-Set-ExchangeOnlineResourceMailboxBookingDelegates.ps1 | Implemented |
| M220 | M3120 | M220-Set-ExchangeOnlineMailboxDelegations.ps1 | Implemented |
| M221 | M3121 | M221-Set-ExchangeOnlineMailboxFolderPermissions.ps1 | Implemented |
| M222 | M3122 | M222-Update-ExchangeOnlineMailEnabledSecurityGroups.ps1 | Implemented |
| M223 | M3123 | M223-Update-ExchangeOnlineDynamicDistributionGroups.ps1 | Implemented |
| M304 | M3204 | M304-PreProvision-OneDrive.ps1 | Implemented |
| M305 | M3205 | M305-Set-OneDriveStorageQuota.ps1 | Implemented |
| M306 | M3206 | M306-Set-OneDriveSiteCollectionAdmins.ps1 | Implemented |
| M307 | M3207 | M307-Set-OneDriveSharingSettings.ps1 | Planned |
| M308 | M3208 | M308-Revoke-OneDriveExternalSharingLinks.ps1 | Planned |
| M309 | M3209 | M309-Set-OneDriveSiteLockState.ps1 | Planned |
| M341 | M3241 | M341-Set-SharePointSiteAdmins.ps1 | Implemented |
| M343 | M3243 | M343-Associate-SharePointSitesToHub.ps1 | Implemented |
| M409 | M3309 | M409-Update-MicrosoftTeams.ps1 | Planned |
| M410 | M3310 | M410-Set-MicrosoftTeamMembers.ps1 | Implemented |
| M411 | M3311 | M411-Update-MicrosoftTeamChannels.ps1 | Implemented |
| M412 | M3312 | M412-Set-MicrosoftTeamChannelMembers.ps1 | Implemented |
| M501 | M0001 | M501-Update-ActiveDirectoryUsers.ps1 | Planned |
| M502 | M0002 | M502-Update-ActiveDirectoryContacts.ps1 | Planned |
| M505 | M0005 | M505-Update-ActiveDirectorySecurityGroups.ps1 | Planned |
| M507 | M0007 | M507-Set-ActiveDirectorySecurityGroupMembers.ps1 | Planned |
| M509 | M0009 | M509-Move-ActiveDirectoryObjects.ps1 | Planned |
| M613 | M0213 | M613-Update-ExchangeOnPremMailContacts.ps1 | Planned |
| M614 | M0214 | M614-Update-ExchangeOnPremDistributionLists.ps1 | Planned |
| M615 | M0215 | M615-Set-ExchangeOnPremDistributionListMembers.ps1 | Planned |
| M616 | M0216 | M616-Update-ExchangeOnPremSharedMailboxes.ps1 | Planned |
| M617 | M0217 | M617-Set-ExchangeOnPremSharedMailboxPermissions.ps1 | Planned |
| M618 | M0218 | M618-Update-ExchangeOnPremResourceMailboxes.ps1 | Planned |
| M619 | M0219 | M619-Set-ExchangeOnPremResourceMailboxBookingDelegates.ps1 | Planned |
| M620 | M0220 | M620-Set-ExchangeOnPremMailboxDelegations.ps1 | Planned |
| M621 | M0221 | M621-Set-ExchangeOnPremMailboxFolderPermissions.ps1 | Planned |
| M622 | M0222 | M622-Update-ExchangeOnPremMailEnabledSecurityGroups.ps1 | Planned |
| M623 | M0223 | M623-Update-ExchangeOnPremDynamicDistributionGroups.ps1 | Planned |
| M701 | M0301 | M701-Update-FileServicesShares.ps1 | Planned |
| M702 | M0302 | M702-Set-FileServicesSharePermissions.ps1 | Planned |
| M703 | M0303 | M703-Set-FileServicesNtfsPermissions.ps1 | Planned |
| M704 | M0304 | M704-Update-FileServicesHomeDrives.ps1 | Planned |
| M705 | M0305 | M705-Set-FileServicesOwnerAndFullControlBySid.ps1 | Planned |
| M706 | M0306 | M706-Grant-FileServicesFullControlBySid.ps1 | Planned |
| P101 | P3001 | P101-Create-EntraUsers.ps1 | Implemented |
| P102 | P3002 | P102-Invite-EntraGuestUsers.ps1 | Implemented |
| P105 | P3005 | P105-Create-EntraAssignedSecurityGroups.ps1 | Implemented |
| P106 | P3006 | P106-Create-EntraDynamicUserSecurityGroups.ps1 | Implemented |
| P108 | P3008 | P108-Create-EntraMicrosoft365Groups.ps1 | Implemented |
| P213 | P3113 | P213-Create-ExchangeOnlineMailContacts.ps1 | Implemented |
| P214 | P3114 | P214-Create-ExchangeOnlineDistributionLists.ps1 | Implemented |
| P215 | P3115 | P215-Create-ExchangeOnlineMailEnabledSecurityGroups.ps1 | Implemented |
| P216 | P3116 | P216-Create-ExchangeOnlineSharedMailboxes.ps1 | Implemented |
| P218 | P3118 | P218-Create-ExchangeOnlineResourceMailboxes.ps1 | Implemented |
| P219 | P3119 | P219-Create-ExchangeOnlineDynamicDistributionGroups.ps1 | Implemented |
| P340 | P3240 | P340-Create-SharePointSites.ps1 | Implemented |
| P342 | P3242 | P342-Create-SharePointHubSites.ps1 | Implemented |
| P409 | P3309 | P409-Create-MicrosoftTeams.ps1 | Implemented |
| P501 | P0001 | P501-Create-ActiveDirectoryUsers.ps1 | Planned |
| P502 | P0002 | P502-Create-ActiveDirectoryContacts.ps1 | Planned |
| P505 | P0005 | P505-Create-ActiveDirectorySecurityGroups.ps1 | Planned |
| P509 | P0009 | P509-Create-ActiveDirectoryOrganizationalUnits.ps1 | Planned |
| P613 | P0213 | P613-Create-ExchangeOnPremMailContacts.ps1 | Planned |
| P614 | P0214 | P614-Create-ExchangeOnPremDistributionLists.ps1 | Planned |
| P615 | P0215 | P615-Create-ExchangeOnPremMailEnabledSecurityGroups.ps1 | Planned |
| P616 | P0216 | P616-Create-ExchangeOnPremSharedMailboxes.ps1 | Planned |
| P618 | P0218 | P618-Create-ExchangeOnPremResourceMailboxes.ps1 | Planned |
| P619 | P0219 | P619-Create-ExchangeOnPremDynamicDistributionGroups.ps1 | Planned |
| P701 | P0301 | P701-Create-FileServicesShares.ps1 | Planned |
| P702 | P0302 | P702-Set-FileServicesSharePermissions.ps1 | Planned |
| P703 | P0303 | P703-Set-FileServicesNtfsPermissions.ps1 | Planned |
| P704 | P0304 | P704-Create-FileServicesHomeDrives.ps1 | Planned |
