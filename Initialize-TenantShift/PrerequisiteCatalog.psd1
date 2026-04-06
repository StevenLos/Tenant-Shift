#
# PrerequisiteCatalog.psd1 — Shared prerequisite definitions for SharedModule.
#
# This catalog is consumed by PrerequisiteEngine.psm1. It is the single source
# of truth for profile-specific runtime, module, and configuration checks.
#
@{
    CatalogVersion = '1.2.0'

    Profiles = @{
        Contributor = @{
            Description = 'Contributor workstation validation for local development and quality-gate execution.'

            Checks = @(
                @{
                    Name            = 'PowerShellVersion'
                    DisplayName     = 'PowerShell version'
                    Category        = 'Runtime'
                    FactName        = 'PowerShellVersion'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '5.1'
                    Severity        = 'Error'
                    Remediation     = 'Use Windows PowerShell 5.1 or PowerShell 7.'
                }
                @{
                    Name            = 'Pester'
                    DisplayName     = 'Pester module'
                    Category        = 'Module'
                    FactName        = 'Pester'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '5.7.1'
                    Severity        = 'Error'
                    Remediation     = 'Install-Module Pester -MinimumVersion 5.7.1 -Scope CurrentUser -Force'
                }
                @{
                    Name            = 'PSScriptAnalyzer'
                    DisplayName     = 'PSScriptAnalyzer module'
                    Category        = 'Module'
                    FactName        = 'PSScriptAnalyzer'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '1.25.0'
                    Severity        = 'Error'
                    Remediation     = 'Install-Module PSScriptAnalyzer -MinimumVersion 1.25.0 -Scope CurrentUser -Force'
                }
            )
        }

        OnlineOperator = @{
            Description = 'Online operator validation for Microsoft 365 automation scripts.'

            Checks = @(
                @{
                    Name            = 'PowerShellVersion'
                    DisplayName     = 'PowerShell version'
                    Category        = 'Runtime'
                    FactName        = 'PowerShellVersion'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '7.4.6'
                    Severity        = 'Error'
                    Remediation     = 'Install PowerShell 7.4.6 or later: https://aka.ms/powershell'
                }
                @{
                    Name            = 'Microsoft.Graph.Authentication'
                    DisplayName     = 'Microsoft.Graph.Authentication module'
                    Category        = 'Module'
                    FactName        = 'Microsoft.Graph.Authentication'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '2.36.0'
                    Severity        = 'Error'
                    Remediation     = 'Install-Module Microsoft.Graph.Authentication -MinimumVersion 2.36.0 -Scope CurrentUser -Force'
                }
                @{
                    Name            = 'Microsoft.Graph.Users'
                    DisplayName     = 'Microsoft.Graph.Users module'
                    Category        = 'Module'
                    FactName        = 'Microsoft.Graph.Users'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '2.36.0'
                    Severity        = 'Error'
                    Remediation     = 'Install-Module Microsoft.Graph.Users -MinimumVersion 2.36.0 -Scope CurrentUser -Force'
                }
                @{
                    Name            = 'Microsoft.Graph.Groups'
                    DisplayName     = 'Microsoft.Graph.Groups module'
                    Category        = 'Module'
                    FactName        = 'Microsoft.Graph.Groups'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '2.36.0'
                    Severity        = 'Error'
                    Remediation     = 'Install-Module Microsoft.Graph.Groups -MinimumVersion 2.36.0 -Scope CurrentUser -Force'
                }
                @{
                    Name            = 'Microsoft.Graph.Teams'
                    DisplayName     = 'Microsoft.Graph.Teams module'
                    Category        = 'Module'
                    FactName        = 'Microsoft.Graph.Teams'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '2.36.0'
                    Severity        = 'Error'
                    Remediation     = 'Install-Module Microsoft.Graph.Teams -MinimumVersion 2.36.0 -Scope CurrentUser -Force'
                }
                @{
                    Name            = 'Microsoft.Graph.Identity.DirectoryManagement'
                    DisplayName     = 'Microsoft.Graph.Identity.DirectoryManagement module'
                    Category        = 'Module'
                    FactName        = 'Microsoft.Graph.Identity.DirectoryManagement'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '2.36.0'
                    Severity        = 'Error'
                    Remediation     = 'Install-Module Microsoft.Graph.Identity.DirectoryManagement -MinimumVersion 2.36.0 -Scope CurrentUser -Force'
                }
                @{
                    Name            = 'Microsoft.Graph.Identity.SignIns'
                    DisplayName     = 'Microsoft.Graph.Identity.SignIns module'
                    Category        = 'Module'
                    FactName        = 'Microsoft.Graph.Identity.SignIns'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '2.36.0'
                    Severity        = 'Error'
                    Remediation     = 'Install-Module Microsoft.Graph.Identity.SignIns -MinimumVersion 2.36.0 -Scope CurrentUser -Force'
                }
                @{
                    Name            = 'Microsoft.Graph.Users.Actions'
                    DisplayName     = 'Microsoft.Graph.Users.Actions module'
                    Category        = 'Module'
                    FactName        = 'Microsoft.Graph.Users.Actions'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '2.36.0'
                    Severity        = 'Error'
                    Remediation     = 'Install-Module Microsoft.Graph.Users.Actions -MinimumVersion 2.36.0 -Scope CurrentUser -Force'
                }
                @{
                    Name            = 'ExchangeOnlineManagement'
                    DisplayName     = 'ExchangeOnlineManagement module'
                    Category        = 'Module'
                    FactName        = 'ExchangeOnlineManagement'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '3.9.2'
                    Severity        = 'Error'
                    Remediation     = 'Install-Module ExchangeOnlineManagement -MinimumVersion 3.9.2 -Scope CurrentUser -Force'
                }
                @{
                    Name            = 'Microsoft.Online.SharePoint.PowerShell'
                    DisplayName     = 'Microsoft.Online.SharePoint.PowerShell module'
                    Category        = 'Module'
                    FactName        = 'Microsoft.Online.SharePoint.PowerShell'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '16.0.27011.12008'
                    Severity        = 'Error'
                    Remediation     = 'Install-Module Microsoft.Online.SharePoint.PowerShell -MinimumVersion 16.0.27011.12008 -Scope CurrentUser -Force'
                }
                @{
                    Name            = 'PnP.PowerShell'
                    DisplayName     = 'PnP.PowerShell module'
                    Category        = 'Module'
                    FactName        = 'PnP.PowerShell'
                    RequirementType = 'VersionRange'
                    MinimumVersion  = '3.1.0'
                    Severity        = 'Error'
                    Remediation     = 'Install-Module PnP.PowerShell -MinimumVersion 3.1.0 -Scope CurrentUser -Force'
                }
                @{
                    Name            = 'PNP_CLIENT_ID'
                    DisplayName     = 'PNP_CLIENT_ID environment variable'
                    Category        = 'Configuration'
                    FactName        = 'PNP_CLIENT_ID'
                    RequirementType = 'NonEmpty'
                    Severity        = 'Error'
                    Remediation     = 'Set $env:PNP_CLIENT_ID = "<your-entra-app-client-id>" or add it to your profile'
                }
            )
        }

        OnPremOperator = @{
            Description = 'On-prem operator validation for Active Directory and Exchange Management Shell scripts.'

            Checks = @(
                @{
                    Name            = 'OperatingSystem'
                    DisplayName     = 'Windows host'
                    Category        = 'Runtime'
                    FactName        = 'IsWindows'
                    RequirementType = 'AllowedValues'
                    AllowedValues   = @($true)
                    ExpectedDescription = 'a Windows host'
                    ActualValueMap  = @{
                        True  = 'Windows'
                        False = 'Non-Windows'
                    }
                    Severity        = 'Error'
                    Remediation     = 'Run OnPrem scripts from a Windows host.'
                }
                @{
                    Name                    = 'PowerShellVersion'
                    DisplayName             = 'PowerShell version'
                    Category                = 'Runtime'
                    FactName                = 'PowerShellVersion'
                    RequirementType         = 'VersionRange'
                    MinimumVersion          = '5.1'
                    MaximumVersionExclusive = '5.2'
                    Severity                = 'Error'
                    Remediation             = 'Run OnPrem scripts from Windows PowerShell 5.1.'
                }
                @{
                    Name            = 'PowerShellEdition'
                    DisplayName     = 'PowerShell edition'
                    Category        = 'Runtime'
                    FactName        = 'PSEdition'
                    RequirementType = 'AllowedValues'
                    AllowedValues   = @('Desktop')
                    ExpectedDescription = 'Desktop edition'
                    Severity        = 'Error'
                    Remediation     = 'Run OnPrem scripts from Windows PowerShell 5.1 (Desktop edition).'
                }
                @{
                    Name            = 'ExchangeManagementShell'
                    DisplayName     = 'Exchange Management Shell session'
                    Category        = 'Runtime'
                    FactName        = 'ExchangeManagementShell'
                    RequirementType = 'Present'
                    ExpectedDescription = 'Exchange Management Shell cmdlets loaded'
                    ActualValueMap  = @{
                        True  = 'Available'
                        False = 'Unavailable'
                    }
                    Severity        = 'Error'
                    Remediation     = 'Launch Exchange Management Shell before running Exchange OnPrem scripts.'
                }
                @{
                    Name            = 'ActiveDirectory'
                    DisplayName     = 'ActiveDirectory module'
                    Category        = 'Module'
                    FactName        = 'ActiveDirectory'
                    RequirementType = 'Present'
                    Severity        = 'Error'
                    Remediation     = 'Install RSAT AD module (Server: Install-WindowsFeature RSAT-AD-PowerShell | Client: Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0)'
                }
            )
        }

        RepoScan = @{
            Description = 'Repository-wide prerequisite discovery profile.'
            Checks      = @()
        }
    }
}
