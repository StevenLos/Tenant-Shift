# Entra User Field Contract (Expanded Model)

This document defines the shared Entra user field families used by:

- `SharedModule/Online/Provision/P-MEID-0010-Create-EntraUsers.ps1`
- `SharedModule/Online/Modify/M-MEID-0010-Update-EntraUsers.ps1`
- `SharedModule/Online/Modify/M-MEID-0040-Set-EntraUserAccountState.ps1`
- `SharedModule/Online/InventoryAndReport/D-MEID-0010-Get-EntraUsers.ps1`

## Design Rules

- Entra-native properties are the source of truth.
- AD parity is handled by naming and mapping where semantics match.
- AD-only attributes that do not have Entra equivalents are treated as `NotApplicable` for Entra scripts.
- `ClearAttributes` (in `M3001`) is the reset mechanism for nullable Entra user properties.

## Shared Field Families

### Core Identity

| CSV Field | Create (P3001) | Update (M3001) | Inventory (IR3001) | Clear in M3001 |
|---|---|---|---|---|
| `UserPrincipalName` | Required | Required (lookup key) | Yes | No |
| `DisplayName` | Required | Yes | Yes | No |
| `GivenName` | Yes | Yes | Yes | Yes |
| `Surname` | Yes | Yes | Yes | Yes |
| `MailNickname` | Yes | Yes | Yes | No |
| `UserType` | Yes (`Member`/`Guest`) | Yes (`Member`/`Guest`) | Yes | No |

### Account and Access

| CSV Field | Create (P3001) | Update (M3001) | Inventory (IR3001) | Clear in M3001 |
|---|---|---|---|---|
| `Password` | Required | Yes (reset) | No | No |
| `ForceChangePasswordNextSignIn` | Yes | Yes (with `Password`) | No | No |
| `ForceChangePasswordNextSignInWithMfa` | Yes | Yes (with `Password`) | No | No |
| `AccountEnabled` | Yes | Yes | Yes | No |
| `UsageLocation` | Yes | Yes | Yes | Yes |
| `PreferredLanguage` | Yes | Yes | Yes | Yes |
| `PasswordPolicies` | Yes | Yes | Yes | Yes |

### Organization and HR

| CSV Field | Create (P3001) | Update (M3001) | Inventory (IR3001) | Clear in M3001 |
|---|---|---|---|---|
| `Department` | Yes | Yes | Yes | Yes |
| `JobTitle` | Yes | Yes | Yes | Yes |
| `CompanyName` | Yes | Yes | Yes | Yes |
| `OfficeLocation` | Yes | Yes | Yes | Yes |
| `EmployeeId` | Yes | Yes | Yes | Yes |
| `EmployeeType` | Yes | Yes | Yes | Yes |
| `EmployeeHireDate` | Yes | Yes | Yes | Yes |

### Contact and Address

| CSV Field | Create (P3001) | Update (M3001) | Inventory (IR3001) | Clear in M3001 |
|---|---|---|---|---|
| `MobilePhone` | Yes | Yes | Yes | Yes |
| `BusinessPhones` | Yes | Yes | Yes | Yes |
| `FaxNumber` | Yes | Yes | Yes | Yes |
| `OtherMails` | Yes | Yes | Yes | Yes |
| `StreetAddress` | Yes | Yes | Yes | Yes |
| `City` | Yes | Yes | Yes | Yes |
| `State` | Yes | Yes | Yes | Yes |
| `PostalCode` | Yes | Yes | Yes | Yes |
| `Country` | Yes | Yes | Yes | Yes |

### Extension Attributes

| CSV Field | Create (P3001) | Update (M3001) | Inventory (IR3001) | Clear in M3001 |
|---|---|---|---|---|
| `ExtensionAttribute1` to `ExtensionAttribute15` | Yes | Yes | Yes | Yes |

### Automation / Metadata Helpers

| CSV Field | Create (P3001) | Update (M3001) | Inventory (IR3001) |
|---|---|---|---|
| `Action` | Optional metadata | Optional metadata | N/A |
| `Notes` | Optional metadata | Optional metadata | N/A |
| `ClearAttributes` | N/A | Yes | N/A |

## Inventory-Only Entra Metadata

`IR3001` additionally exports:

- `UserId`
- `Mail`
- `CreatedDateTime`
- `LastPasswordChangeDateTime`
- `OnPremisesSyncEnabled`
- `OnPremisesImmutableId`
- `OnPremisesDistinguishedName`
- `OnPremisesDomainName`
- `OnPremisesSamAccountName`
- `OnPremisesSecurityIdentifier`
- `OnPremisesLastSyncDateTime`
- `AssignedLicenseSkuIds`
- `AssignedPlanServiceIds`

## AD to Entra Mapping Guidance

- Direct or close semantic parity examples:
  - `UserPrincipalName`, `DisplayName`, `GivenName`, `Surname`, `Department`, `JobTitle`, `CompanyName`, `OfficeLocation`, `MobilePhone`, `StreetAddress`, `City`, `State`, `PostalCode`, `Country`, `ExtensionAttribute1-15`.
- Entra-first differences:
  - `BusinessPhones`, `OtherMails`, `PreferredLanguage`, `UsageLocation`, `PasswordPolicies`.
- AD-centric fields without direct Entra equivalents remain out of Entra user scripts:
  - `SamAccountName`, `UserWorkstations`, `ScriptPath`, `ProfilePath`, `HomeDirectory`, `HomeDrive`, most domain-controller operational timestamps.

## Notes on `ClearAttributes`

- `ClearAttributes` accepts semicolon-delimited logical CSV names.
- Supported clear set in `M3001` includes nullable user profile/contact/org fields and `ExtensionAttribute1-15`.
- A field cannot be both set and cleared in the same row.
