param(
    [Parameter(Mandatory = $false)]
    [string]$InputCsv = ".\EntraUserInput.csv",

    [Parameter(Mandatory = $false)]
    [string]$OutputCsv = ".\EntraUserInventory.csv"
)

# ============================================================
# Bulk Entra User Inventory / Local Email Matching
#
# Purpose:
#   1. Read email addresses from a CSV.
#   2. Download all Entra Member users once.
#   3. Build local indexes using:
#        - UserPrincipalName
#        - Mail
#        - ProxyAddresses
#        - OtherMails
#   4. Match input addresses locally.
#   5. Export extensive Entra / hybrid identity information.
#
# Requires:
#   Microsoft.Graph.Users
#
# Install if required:
#   Install-Module Microsoft.Graph -Scope CurrentUser
# ============================================================

Connect-MgGraph -Scopes @(
    "User.Read.All"
    "Directory.Read.All"
    "AuditLog.Read.All"
)

if (-not (Test-Path $InputCsv)) {
    throw "Input CSV not found: $InputCsv"
}

$InputUsers = @(Import-Csv $InputCsv)

if ($InputUsers.Count -eq 0) {
    throw "Input CSV contains no rows."
}

if (-not ($InputUsers[0].PSObject.Properties.Name -contains "EmailAddress")) {
    throw "Input CSV must contain a column named EmailAddress."
}

$Properties = @(
    "id"
    "accountEnabled"
    "ageGroup"
    "businessPhones"
    "city"
    "companyName"
    "consentProvidedForMinor"
    "country"
    "createdDateTime"
    "creationType"
    "department"
    "displayName"
    "employeeHireDate"
    "employeeId"
    "employeeLeaveDateTime"
    "employeeOrgData"
    "employeeType"
    "externalUserState"
    "externalUserStateChangeDateTime"
    "faxNumber"
    "givenName"
    "identities"
    "imAddresses"
    "isResourceAccount"
    "jobTitle"
    "lastPasswordChangeDateTime"
    "legalAgeGroupClassification"
    "mail"
    "mailNickname"
    "mobilePhone"
    "officeLocation"
    "onPremisesDistinguishedName"
    "onPremisesDomainName"
    "onPremisesExtensionAttributes"
    "onPremisesImmutableId"
    "onPremisesLastSyncDateTime"
    "onPremisesProvisioningErrors"
    "onPremisesSamAccountName"
    "onPremisesSecurityIdentifier"
    "onPremisesSyncEnabled"
    "onPremisesUserPrincipalName"
    "otherMails"
    "passwordPolicies"
    "postalCode"
    "preferredLanguage"
    "proxyAddresses"
    "signInActivity"
    "state"
    "streetAddress"
    "surname"
    "usageLocation"
    "userPrincipalName"
    "userType"
)

Write-Host ""
Write-Host "Downloading Entra Member users..." -ForegroundColor Cyan

$StartTime = Get-Date

$AllEntraUsers = @(
    Get-MgUser `
        -All `
        -Filter "userType eq 'Member'" `
        -Property $Properties `
        -ErrorAction Stop
)

$Elapsed = (Get-Date) - $StartTime

Write-Host "Downloaded $($AllEntraUsers.Count) Entra Member users." -ForegroundColor Green
Write-Host "Download duration: $($Elapsed.ToString('hh\:mm\:ss'))"
Write-Host ""

$UPNIndex       = @{}
$MailIndex      = @{}
$ProxyIndex     = @{}
$OtherMailIndex = @{}

function Add-ToIndex {
    param(
        [hashtable]$Index,
        [string]$Key,
        [object]$User
    )

    if ([string]::IsNullOrWhiteSpace($Key)) {
        return
    }

    $NormalizedKey = $Key.Trim().ToLowerInvariant()

    if (-not $Index.ContainsKey($NormalizedKey)) {
        $Index[$NormalizedKey] = [System.Collections.ArrayList]::new()
    }

    [void]$Index[$NormalizedKey].Add($User)
}

Write-Host "Building local lookup indexes..." -ForegroundColor Cyan

$Counter = 0
$TotalUsers = $AllEntraUsers.Count

foreach ($User in $AllEntraUsers) {
    $Counter++

    if (($Counter % 500) -eq 0 -or $Counter -eq $TotalUsers) {
        $Percent = [math]::Round(($Counter / $TotalUsers) * 100, 0)

        Write-Progress `
            -Activity "Building Entra lookup indexes" `
            -Status "$Counter of $TotalUsers" `
            -PercentComplete $Percent
    }

    Add-ToIndex -Index $UPNIndex -Key $User.UserPrincipalName -User $User
    Add-ToIndex -Index $MailIndex -Key $User.Mail -User $User

    foreach ($ProxyAddress in @($User.ProxyAddresses)) {
        if ([string]::IsNullOrWhiteSpace($ProxyAddress)) {
            continue
        }

        $Address = $ProxyAddress

        if ($Address -match '^(?i)smtp:') {
            $Address = $Address.Substring(5)
        }

        Add-ToIndex -Index $ProxyIndex -Key $Address -User $User
    }

    foreach ($OtherMail in @($User.OtherMails)) {
        Add-ToIndex -Index $OtherMailIndex -Key $OtherMail -User $User
    }
}

Write-Progress -Activity "Building Entra lookup indexes" -Completed
Write-Host "Indexes built." -ForegroundColor Green
Write-Host ""

function Add-UniqueMatches {
    param(
        [System.Collections.ArrayList]$Destination,
        [object[]]$Users,
        [string]$MatchType
    )

    foreach ($User in @($Users)) {
        if ($null -eq $User) {
            continue
        }

        $Existing = $Destination | Where-Object { $_.User.Id -eq $User.Id }

        if (-not $Existing) {
            [void]$Destination.Add(
                [PSCustomObject]@{
                    User      = $User
                    MatchType = $MatchType
                }
            )
        }
    }
}

Write-Host "Matching input email addresses locally..." -ForegroundColor Cyan

$Results = [System.Collections.ArrayList]::new()

$InputCounter = 0
$InputTotal = $InputUsers.Count

foreach ($Row in $InputUsers) {
    $InputCounter++

    $EmailAddress = [string]$Row.EmailAddress

    if ([string]::IsNullOrWhiteSpace($EmailAddress)) {
        continue
    }

    $EmailAddress = $EmailAddress.Trim()
    $NormalizedEmail = $EmailAddress.ToLowerInvariant()

    $Percent = [math]::Round(($InputCounter / $InputTotal) * 100, 0)

    Write-Progress `
        -Activity "Matching input addresses" `
        -Status "$InputCounter of $InputTotal - $EmailAddress" `
        -PercentComplete $Percent

    $Matches = [System.Collections.ArrayList]::new()

    if ($UPNIndex.ContainsKey($NormalizedEmail)) {
        Add-UniqueMatches -Destination $Matches -Users $UPNIndex[$NormalizedEmail] -MatchType "UserPrincipalName"
    }

    if ($MailIndex.ContainsKey($NormalizedEmail)) {
        Add-UniqueMatches -Destination $Matches -Users $MailIndex[$NormalizedEmail] -MatchType "Mail"
    }

    if ($ProxyIndex.ContainsKey($NormalizedEmail)) {
        Add-UniqueMatches -Destination $Matches -Users $ProxyIndex[$NormalizedEmail] -MatchType "ProxyAddress"
    }

    if ($OtherMailIndex.ContainsKey($NormalizedEmail)) {
        Add-UniqueMatches -Destination $Matches -Users $OtherMailIndex[$NormalizedEmail] -MatchType "OtherMail"
    }

    if ($Matches.Count -eq 0) {
        [void]$Results.Add(
            [PSCustomObject]@{
                InputEmailAddress = $EmailAddress
                Found             = $false
                MatchCount        = 0
                MatchType         = $null
                AmbiguousMatch    = $false
                DisplayName       = $null
                EntraObjectId     = $null
                UserPrincipalName = $null
                IdentityState     = "Not Found"
            }
        )

        continue
    }

    foreach ($Match in $Matches) {
        $U = $Match.User
        $Ext = $U.OnPremisesExtensionAttributes

        if ($U.OnPremisesSyncEnabled -eq $true) {
            $IdentityState = "Currently Synced From On-Premises AD"
        }
        elseif (
            $U.OnPremisesImmutableId -or
            $U.OnPremisesSecurityIdentifier -or
            $U.OnPremisesSamAccountName -or
            $U.OnPremisesDomainName -or
            $U.OnPremisesDistinguishedName
        ) {
            $IdentityState = "Cloud Managed - On-Premises Metadata Remains"
        }
        else {
            $IdentityState = "Cloud Only"
        }

        $PrimarySMTP = (
            @($U.ProxyAddresses) |
            Where-Object { $_ -cmatch '^SMTP:' } |
            Select-Object -First 1
        )

        if ($PrimarySMTP) {
            $PrimarySMTP = $PrimarySMTP.Substring(5)
        }

        $IdentityStrings = @()
        foreach ($Identity in @($U.Identities)) {
            if ($Identity) {
                $IdentityStrings += (
                    "{0}|{1}|{2}" -f `
                        $Identity.SignInType,
                        $Identity.Issuer,
                        $Identity.IssuerAssignedId
                )
            }
        }

        $ProvisioningErrors = @()
        foreach ($ProvisioningError in @($U.OnPremisesProvisioningErrors)) {
            if ($ProvisioningError) {
                $ProvisioningErrors += (
                    "{0}|{1}|{2}" -f `
                        $ProvisioningError.Category,
                        $ProvisioningError.PropertyCausingError,
                        $ProvisioningError.Value
                )
            }
        }

        [void]$Results.Add(
            [PSCustomObject]@{
                InputEmailAddress                = $EmailAddress
                Found                            = $true
                MatchCount                       = $Matches.Count
                MatchType                        = $Match.MatchType
                AmbiguousMatch                   = ($Matches.Count -gt 1)

                DisplayName                      = $U.DisplayName
                GivenName                        = $U.GivenName
                Surname                          = $U.Surname
                EntraObjectId                    = $U.Id
                UserPrincipalName                = $U.UserPrincipalName
                AccountEnabled                   = $U.AccountEnabled
                UserType                         = $U.UserType
                IsResourceAccount                = $U.IsResourceAccount
                CreatedDateTime                  = $U.CreatedDateTime
                CreationType                     = $U.CreationType

                Mail                             = $U.Mail
                MailNickname                     = $U.MailNickname
                PrimarySMTPAddress               = $PrimarySMTP
                ProxyAddresses                   = (@($U.ProxyAddresses) -join ";")
                OtherMails                       = (@($U.OtherMails) -join ";")
                IMAddresses                      = (@($U.ImAddresses) -join ";")

                IdentityState                    = $IdentityState

                OnPremisesSyncEnabled            = $U.OnPremisesSyncEnabled
                OnPremisesLastSyncDateTime       = $U.OnPremisesLastSyncDateTime
                OnPremisesImmutableId            = $U.OnPremisesImmutableId
                OnPremisesSID                    = $U.OnPremisesSecurityIdentifier
                OnPremisesSamAccountName         = $U.OnPremisesSamAccountName
                OnPremisesUserPrincipalName      = $U.OnPremisesUserPrincipalName
                OnPremisesDomainName             = $U.OnPremisesDomainName
                OnPremisesDistinguishedName      = $U.OnPremisesDistinguishedName
                OnPremisesProvisioningErrors     = ($ProvisioningErrors -join ";")

                Identities                       = ($IdentityStrings -join ";")

                EmployeeId                       = $U.EmployeeId
                EmployeeType                     = $U.EmployeeType
                EmployeeHireDate                 = $U.EmployeeHireDate
                EmployeeLeaveDateTime            = $U.EmployeeLeaveDateTime
                CompanyName                      = $U.CompanyName
                Department                       = $U.Department
                JobTitle                         = $U.JobTitle
                OfficeLocation                   = $U.OfficeLocation

                MobilePhone                      = $U.MobilePhone
                BusinessPhones                   = (@($U.BusinessPhones) -join ";")
                FaxNumber                        = $U.FaxNumber

                StreetAddress                    = $U.StreetAddress
                City                             = $U.City
                State                            = $U.State
                PostalCode                       = $U.PostalCode
                Country                          = $U.Country
                UsageLocation                    = $U.UsageLocation
                PreferredLanguage                = $U.PreferredLanguage

                LastPasswordChangeDateTime       = $U.LastPasswordChangeDateTime
                PasswordPolicies                 = $U.PasswordPolicies

                LastSignInDateTime               = $U.SignInActivity.LastSignInDateTime
                LastSuccessfulSignInDateTime     = $U.SignInActivity.LastSuccessfulSignInDateTime
                LastNonInteractiveSignInDateTime = $U.SignInActivity.LastNonInteractiveSignInDateTime

                ExternalUserState                = $U.ExternalUserState
                ExternalUserStateChangeDateTime  = $U.ExternalUserStateChangeDateTime

                ExtensionAttribute1              = $Ext.ExtensionAttribute1
                ExtensionAttribute2              = $Ext.ExtensionAttribute2
                ExtensionAttribute3              = $Ext.ExtensionAttribute3
                ExtensionAttribute4              = $Ext.ExtensionAttribute4
                ExtensionAttribute5              = $Ext.ExtensionAttribute5
                ExtensionAttribute6              = $Ext.ExtensionAttribute6
                ExtensionAttribute7              = $Ext.ExtensionAttribute7
                ExtensionAttribute8              = $Ext.ExtensionAttribute8
                ExtensionAttribute9              = $Ext.ExtensionAttribute9
                ExtensionAttribute10             = $Ext.ExtensionAttribute10
                ExtensionAttribute11             = $Ext.ExtensionAttribute11
                ExtensionAttribute12             = $Ext.ExtensionAttribute12
                ExtensionAttribute13             = $Ext.ExtensionAttribute13
                ExtensionAttribute14             = $Ext.ExtensionAttribute14
                ExtensionAttribute15             = $Ext.ExtensionAttribute15
            }
        )
    }
}

Write-Progress -Activity "Matching input addresses" -Completed

$Results |
    Export-Csv `
        -Path $OutputCsv `
        -NoTypeInformation `
        -Encoding UTF8

$FoundInputAddresses = @(
    $Results |
    Where-Object { $_.Found -eq $true } |
    Select-Object -ExpandProperty InputEmailAddress -Unique
).Count

$NotFound = @(
    $Results |
    Where-Object { $_.Found -eq $false }
).Count

$Ambiguous = @(
    $Results |
    Where-Object { $_.AmbiguousMatch -eq $true } |
    Select-Object -ExpandProperty InputEmailAddress -Unique
).Count

$CurrentlySynced = @(
    $Results |
    Where-Object { $_.IdentityState -eq "Currently Synced From On-Premises AD" }
).Count

$CloudWithMetadata = @(
    $Results |
    Where-Object { $_.IdentityState -eq "Cloud Managed - On-Premises Metadata Remains" }
).Count

$CloudOnly = @(
    $Results |
    Where-Object { $_.IdentityState -eq "Cloud Only" }
).Count

Write-Host ""
Write-Host "============================================================"
Write-Host "Complete" -ForegroundColor Green
Write-Host "============================================================"
Write-Host "Input email addresses : $InputTotal"
Write-Host "Matched addresses     : $FoundInputAddresses"
Write-Host "Not found             : $NotFound"
Write-Host "Ambiguous matches     : $Ambiguous"
Write-Host ""
Write-Host "Currently AD synced   : $CurrentlySynced"
Write-Host "Cloud + AD metadata   : $CloudWithMetadata"
Write-Host "Cloud only            : $CloudOnly"
Write-Host ""
Write-Host "Entra users downloaded: $TotalUsers"
Write-Host "Output file           : $OutputCsv"
Write-Host "============================================================"
