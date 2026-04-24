<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260423-000000

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement
Microsoft.Graph.Authentication
Microsoft.Graph.Groups
Microsoft.Graph.Users

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Gets ExchangeOnlineMailboxAccessByUser and exports results to CSV.

.DESCRIPTION
    For each in-scope user, discovers every shared mailbox, resource mailbox (room), and
    equipment mailbox where that user holds any permission — Full Access, ReadOnly, Send As,
    or Send on Behalf — whether that permission was granted directly to the user or to a
    group the user belongs to (resolved transitively via Microsoft Graph).
    Accepts target input either from a CSV file (FromCsv parameter set, using -InputCsvPath)
    or by enumerating all licensed users in the tenant (-DiscoverAll parameter set).
    All results — including rows that could not be processed — are written to the output CSV.

.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER DiscoverAll
    Enumerate all licensed users in the tenant rather than processing from an input CSV file.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.

.EXAMPLE
    .\D-EXOL-0500-Get-ExchangeOnlineMailboxAccessByUser.ps1 -InputCsvPath .\Scope-Users.input.csv
    Inventory the objects listed in the input CSV file.

.EXAMPLE
    .\D-EXOL-0500-Get-ExchangeOnlineMailboxAccessByUser.ps1 -DiscoverAll
    Discover and inventory all objects in scope, writing results to the default output path.

.NOTES
    Version:          1.0
    Required modules: ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users
    Required roles:   Exchange Recipient Administrator (read-only), Global Reader
    Limitations:      None known.

    CSV Fields:
    Column              Type    Required  Description
    ------              ----    --------  -----------
    UserPrincipalName   String  Yes       UPN of the user to evaluate
#>
#Requires -Version 7.0

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'DiscoverAll')]
    [switch]$DiscoverAll,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Discover_OutputCsvPath') -ChildPath ("Results_D-EXOL-0500-Get-ExchangeOnlineMailboxAccessByUser_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function New-InventoryResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [int]$RowNumber,
        [Parameter(Mandatory)] [string]$PrimaryKey,
        [Parameter(Mandatory)] [string]$Action,
        [Parameter(Mandatory)] [string]$Status,
        [Parameter(Mandatory)] [string]$Message,
        [Parameter(Mandatory)] [hashtable]$Data
    )

    $base = New-ResultObject -RowNumber $RowNumber -PrimaryKey $PrimaryKey -Action $Action -Status $Status -Message $Message
    $ordered = [ordered]@{}
    foreach ($prop in $base.PSObject.Properties.Name) { $ordered[$prop] = $base.$prop }
    foreach ($key in $Data.Keys) { $ordered[$key] = $Data[$key] }
    return [PSCustomObject]$ordered
}

function New-EmptyAccessRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$UserPrincipalName,
        [Parameter(Mandatory)] [string]$MailboxIdentity,
        [Parameter(Mandatory)] [string]$MailboxDisplayName,
        [Parameter(Mandatory)] [string]$MailboxType,
        [Parameter(Mandatory)] [string]$PermissionType,
        [string]$AccessSource                = '',
        [string]$GrantingGroupIdentity       = '',
        [string]$GrantingGroupDisplayName    = ''
    )

    return [ordered]@{
        UserPrincipalName        = $UserPrincipalName
        MailboxIdentity          = $MailboxIdentity
        MailboxDisplayName       = $MailboxDisplayName
        MailboxType              = $MailboxType
        PermissionType           = $PermissionType
        AccessSource             = $AccessSource
        GrantingGroupIdentity    = $GrantingGroupIdentity
        GrantingGroupDisplayName = $GrantingGroupDisplayName
    }
}

function Get-ObjectPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$InputObject,
        [Parameter(Mandatory)] [string]$PropertyName
    )

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($property) {
        return $property.Value
    }

    return $null
}

function Get-StringPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$InputObject,
        [Parameter(Mandatory)] [string]$PropertyName
    )

    return ([string](Get-ObjectPropertyValue -InputObject $InputObject -PropertyName $PropertyName)).Trim()
}

function Test-IsTrusteeExcluded {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Trustee
    )

    if ([string]::IsNullOrWhiteSpace($Trustee))               { return $true }
    if ($Trustee -match '^NT AUTHORITY\\')                     { return $true }
    if ($Trustee -match '^S-1-5-')                             { return $true }
    if ($Trustee -eq 'SELF')                                   { return $true }

    return $false
}

function Test-IsGroupRecipientType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$RecipientTypeDetails
    )

    return ($RecipientTypeDetails -match 'Group')
}

$requiredHeaders = @('UserPrincipalName')

$reportPropertyOrder = @(
    'TimestampUtc', 'RowNumber', 'PrimaryKey', 'Action', 'Status', 'Message', 'ScopeMode',
    'UserPrincipalName', 'MailboxIdentity', 'MailboxDisplayName', 'MailboxType',
    'PermissionType', 'AccessSource', 'GrantingGroupIdentity', 'GrantingGroupDisplayName'
)

Write-Status -Message 'Starting Exchange Online mailbox access by user inventory script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('GroupMember.Read.All')
Ensure-ExchangeConnection

$scopeMode = 'Csv'
if ($PSCmdlet.ParameterSetName -eq 'DiscoverAll') {
    $scopeMode = 'DiscoverAll'
    Write-Status -Message 'DiscoverAll enabled. Enumerating all licensed users in the tenant.'

    $licensedUsers = @(Invoke-WithRetry -OperationName 'Enumerate all licensed users' -ScriptBlock {
        Get-MgUser `
            -Filter "assignedLicenses/`$count ne 0" `
            -CountVariable lc `
            -ConsistencyLevel eventual `
            -All `
            -Property 'Id,UserPrincipalName,DisplayName' `
            -ErrorAction Stop
    })

    Write-Status -Message "DiscoverAll: found $($licensedUsers.Count) licensed users."
    $rows = @($licensedUsers | ForEach-Object { [PSCustomObject]@{ UserPrincipalName = $_.UserPrincipalName } })
}
else {
    $rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
}

# Cache all shared, resource, and equipment mailboxes once before the user loop.
Write-Status -Message 'Loading all shared, resource, and equipment mailboxes. This may take several minutes for large tenants.'
$allMailboxes = @(Invoke-WithRetry -OperationName 'Enumerate shared/resource/equipment mailboxes' -ScriptBlock {
    Get-EXOMailbox `
        -ResultSize Unlimited `
        -RecipientTypeDetails SharedMailbox, RoomMailbox, EquipmentMailbox `
        -Properties Identity, DisplayName, RecipientTypeDetails, GrantSendOnBehalfTo `
        -ErrorAction Stop
})
Write-Status -Message "Loaded $($allMailboxes.Count) mailboxes for access evaluation."

# Recipient type cache: keyed by trustee identity string, value is RecipientTypeDetails string.
$recipientTypeCache = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)

# EXO ObjectId cache: keyed by trustee identity, value is Entra ObjectId string (may be empty string if unresolvable).
$trusteeObjectIdCache = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)

$results   = [System.Collections.Generic.List[object]]::new()
$rowNumber = 1

foreach ($row in $rows) {
    $upn = ([string]$row.UserPrincipalName).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw 'UserPrincipalName is required and cannot be blank.'
        }

        # ── Step 1: Resolve the user's Entra ObjectId ──────────────────────────────
        $userObj = Invoke-WithRetry -OperationName "Resolve Entra user $upn" -ScriptBlock {
            Get-MgUser -UserId $upn -Property 'Id,UserPrincipalName,DisplayName' -ErrorAction Stop
        }
        $userId = ([string]$userObj.Id).Trim()

        # ── Step 2: Fetch all transitive group memberships for this user ───────────
        $transitiveMemberships = @(Invoke-WithRetry -OperationName "Get transitive group memberships for $upn" -ScriptBlock {
            Get-MgUserTransitiveMemberOf -UserId $userId -All -Property 'Id,DisplayName' -ErrorAction Stop
        })

        # Build a HashSet of group ObjectIds the user belongs to (transitively).
        $userGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        # Build a lookup of groupId -> displayName for reporting.
        $groupDisplayNameById = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($membership in $transitiveMemberships) {
            $gid   = ([string]$membership.Id).Trim()
            $gname = ([string]$membership.DisplayName).Trim()
            [void]$userGroupIds.Add($gid)
            if (-not $groupDisplayNameById.ContainsKey($gid)) {
                $groupDisplayNameById[$gid] = $gname
            }
        }

        # ── Step 3: Evaluate each mailbox ─────────────────────────────────────────
        $accessRows = [System.Collections.Generic.List[object]]::new()

        foreach ($mailbox in $allMailboxes) {
            $mailboxIdentity    = ([string]$mailbox.Identity).Trim()
            $mailboxDisplayName = ([string]$mailbox.DisplayName).Trim()
            $mailboxType        = ([string]$mailbox.RecipientTypeDetails).Trim()

            # ── Helper: resolve a trustee string to (isUser, isGroup, groupId, groupDisplayName) ──
            $resolveTrustee = {
                param(
                    [Parameter(Mandatory)]
                    [string]$TrusteeHint
                )

                $result = [PSCustomObject]@{
                    IsDirectUser         = $false
                    IsGroupMember        = $false
                    GroupObjectId        = ''
                    GroupDisplayName     = ''
                }

                # Direct user match: UPN or identity contains the UPN.
                if ($TrusteeHint -eq $upn -or
                    $TrusteeHint.StartsWith($upn + ' ', [System.StringComparison]::OrdinalIgnoreCase)) {
                    $result.IsDirectUser = $true
                    return $result
                }

                # Attempt to determine RecipientTypeDetails and ObjectId for this trustee.
                $rtd = ''
                $oid = ''

                if ($trusteeObjectIdCache.ContainsKey($TrusteeHint)) {
                    $oid = $trusteeObjectIdCache[$TrusteeHint]
                }

                if ($recipientTypeCache.ContainsKey($TrusteeHint)) {
                    $rtd = $recipientTypeCache[$TrusteeHint]
                }
                else {
                    try {
                        $recipient = Invoke-WithRetry -OperationName "Lookup EXO recipient $TrusteeHint" -ScriptBlock {
                            Get-EXORecipient -Identity $TrusteeHint -Properties 'RecipientTypeDetails,ExternalDirectoryObjectId' -ErrorAction Stop
                        }
                        $rtd = ([string]$recipient.RecipientTypeDetails).Trim()
                        $oid = ([string]$recipient.ExternalDirectoryObjectId).Trim()
                        $recipientTypeCache[$TrusteeHint] = $rtd
                        $trusteeObjectIdCache[$TrusteeHint] = $oid
                    }
                    catch {
                        $recipientTypeCache[$TrusteeHint] = ''
                        $trusteeObjectIdCache[$TrusteeHint] = ''
                    }
                }

                # Check if the trustee is the user themselves by ObjectId or UPN match.
                if (-not [string]::IsNullOrWhiteSpace($oid) -and
                    $oid -eq $userId) {
                    $result.IsDirectUser = $true
                    return $result
                }

                # Check if the trustee is a group and the user is a transitive member.
                if (Test-IsGroupRecipientType -RecipientTypeDetails $rtd) {
                    if (-not [string]::IsNullOrWhiteSpace($oid) -and $userGroupIds.Contains($oid)) {
                        $gdn = ''
                        if ($groupDisplayNameById.ContainsKey($oid)) { $gdn = $groupDisplayNameById[$oid] }
                        $result.IsGroupMember    = $true
                        $result.GroupObjectId    = $oid
                        $result.GroupDisplayName = $gdn
                        return $result
                    }
                }

                return $result
            }

            # ── FullAccess and ReadOnly ────────────────────────────────────────────
            $mailboxPermissions = @(Invoke-WithRetry -OperationName "Get mailbox permissions for $mailboxIdentity" -ScriptBlock {
                Get-EXOMailboxPermission -Identity $mailboxIdentity -ErrorAction Stop
            })

            foreach ($perm in $mailboxPermissions) {
                if ([bool](Get-ObjectPropertyValue -InputObject $perm -PropertyName 'Deny') -eq $true)        { continue }
                if ([bool](Get-ObjectPropertyValue -InputObject $perm -PropertyName 'IsInherited') -eq $true) { continue }

                $trustee = Get-StringPropertyValue -InputObject $perm -PropertyName 'User'
                if (Test-IsTrusteeExcluded -Trustee $trustee) { continue }

                $accessRights = @((Get-ObjectPropertyValue -InputObject $perm -PropertyName 'AccessRights') | ForEach-Object { ([string]$_).Trim() })

                $permissionTypes = @()
                if ($accessRights -contains 'FullAccess')     { $permissionTypes += 'FullAccess' }
                if ($accessRights -contains 'ReadPermission') { $permissionTypes += 'ReadOnly' }

                if ($permissionTypes.Count -eq 0) { continue }

                $resolved = & $resolveTrustee -TrusteeHint $trustee

                foreach ($permType in $permissionTypes) {
                    if ($resolved.IsDirectUser) {
                        $accessRows.Add([ordered]@{
                            UserPrincipalName        = $upn
                            MailboxIdentity          = $mailboxIdentity
                            MailboxDisplayName       = $mailboxDisplayName
                            MailboxType              = $mailboxType
                            PermissionType           = $permType
                            AccessSource             = 'Direct'
                            GrantingGroupIdentity    = ''
                            GrantingGroupDisplayName = ''
                        })
                    }
                    elseif ($resolved.IsGroupMember) {
                        $accessRows.Add([ordered]@{
                            UserPrincipalName        = $upn
                            MailboxIdentity          = $mailboxIdentity
                            MailboxDisplayName       = $mailboxDisplayName
                            MailboxType              = $mailboxType
                            PermissionType           = $permType
                            AccessSource             = 'Group'
                            GrantingGroupIdentity    = $resolved.GroupObjectId
                            GrantingGroupDisplayName = $resolved.GroupDisplayName
                        })
                    }
                }
            }

            # ── Send As ───────────────────────────────────────────────────────────
            $recipientPermissions = @(Invoke-WithRetry -OperationName "Get recipient permissions (SendAs) for $mailboxIdentity" -ScriptBlock {
                Get-EXORecipientPermission -Identity $mailboxIdentity -AccessRights SendAs -ErrorAction Stop
            })

            foreach ($perm in $recipientPermissions) {
                if ([bool](Get-ObjectPropertyValue -InputObject $perm -PropertyName 'Deny') -eq $true) { continue }

                $trustee = Get-StringPropertyValue -InputObject $perm -PropertyName 'Trustee'
                if (Test-IsTrusteeExcluded -Trustee $trustee) { continue }

                $resolved = & $resolveTrustee -TrusteeHint $trustee

                if ($resolved.IsDirectUser) {
                    $accessRows.Add([ordered]@{
                        UserPrincipalName        = $upn
                        MailboxIdentity          = $mailboxIdentity
                        MailboxDisplayName       = $mailboxDisplayName
                        MailboxType              = $mailboxType
                        PermissionType           = 'SendAs'
                        AccessSource             = 'Direct'
                        GrantingGroupIdentity    = ''
                        GrantingGroupDisplayName = ''
                    })
                }
                elseif ($resolved.IsGroupMember) {
                    $accessRows.Add([ordered]@{
                        UserPrincipalName        = $upn
                        MailboxIdentity          = $mailboxIdentity
                        MailboxDisplayName       = $mailboxDisplayName
                        MailboxType              = $mailboxType
                        PermissionType           = 'SendAs'
                        AccessSource             = 'Group'
                        GrantingGroupIdentity    = $resolved.GroupObjectId
                        GrantingGroupDisplayName = $resolved.GroupDisplayName
                    })
                }
            }

            # ── Send on Behalf ────────────────────────────────────────────────────
            $sendOnBehalfList = @(Get-ObjectPropertyValue -InputObject $mailbox -PropertyName 'GrantSendOnBehalfTo')

            foreach ($delegate in $sendOnBehalfList) {
                $delegateHint = ''
                if ($delegate -is [string]) {
                    $delegateHint = ([string]$delegate).Trim()
                }
                else {
                    $delegateHint = Get-StringPropertyValue -InputObject $delegate -PropertyName 'DistinguishedName'
                    if ([string]::IsNullOrWhiteSpace($delegateHint)) {
                        $delegateHint = Get-StringPropertyValue -InputObject $delegate -PropertyName 'Name'
                    }
                }

                if ([string]::IsNullOrWhiteSpace($delegateHint)) { continue }
                if (Test-IsTrusteeExcluded -Trustee $delegateHint) { continue }

                $resolved = & $resolveTrustee -TrusteeHint $delegateHint

                if ($resolved.IsDirectUser) {
                    $accessRows.Add([ordered]@{
                        UserPrincipalName        = $upn
                        MailboxIdentity          = $mailboxIdentity
                        MailboxDisplayName       = $mailboxDisplayName
                        MailboxType              = $mailboxType
                        PermissionType           = 'SendOnBehalf'
                        AccessSource             = 'Direct'
                        GrantingGroupIdentity    = ''
                        GrantingGroupDisplayName = ''
                    })
                }
                elseif ($resolved.IsGroupMember) {
                    $accessRows.Add([ordered]@{
                        UserPrincipalName        = $upn
                        MailboxIdentity          = $mailboxIdentity
                        MailboxDisplayName       = $mailboxDisplayName
                        MailboxType              = $mailboxType
                        PermissionType           = 'SendOnBehalf'
                        AccessSource             = 'Group'
                        GrantingGroupIdentity    = $resolved.GroupObjectId
                        GrantingGroupDisplayName = $resolved.GroupDisplayName
                    })
                }
            }
        }

        # ── Emit results for this user ─────────────────────────────────────────────
        if ($accessRows.Count -eq 0) {
            $results.Add((New-InventoryResult `
                -RowNumber  $rowNumber `
                -PrimaryKey "$upn||" `
                -Action     'GetExchangeMailboxAccessByUser' `
                -Status     'Success' `
                -Message    'No mailbox access entitlements found.' `
                -Data       ([ordered]@{
                    UserPrincipalName        = $upn
                    MailboxIdentity          = ''
                    MailboxDisplayName       = ''
                    MailboxType              = ''
                    PermissionType           = ''
                    AccessSource             = ''
                    GrantingGroupIdentity    = ''
                    GrantingGroupDisplayName = ''
                })))
        }
        else {
            foreach ($accessRow in $accessRows) {
                $pk = "$upn|$($accessRow.MailboxIdentity)|$($accessRow.PermissionType)"
                $results.Add((New-InventoryResult `
                    -RowNumber  $rowNumber `
                    -PrimaryKey $pk `
                    -Action     'GetExchangeMailboxAccessByUser' `
                    -Status     'Success' `
                    -Message    'Mailbox access entitlement exported.' `
                    -Data       ([ordered]@{
                        UserPrincipalName        = $accessRow.UserPrincipalName
                        MailboxIdentity          = $accessRow.MailboxIdentity
                        MailboxDisplayName       = $accessRow.MailboxDisplayName
                        MailboxType              = $accessRow.MailboxType
                        PermissionType           = $accessRow.PermissionType
                        AccessSource             = $accessRow.AccessSource
                        GrantingGroupIdentity    = $accessRow.GrantingGroupIdentity
                        GrantingGroupDisplayName = $accessRow.GrantingGroupDisplayName
                    })))
            }
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($upn) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-InventoryResult `
            -RowNumber  $rowNumber `
            -PrimaryKey $upn `
            -Action     'GetExchangeMailboxAccessByUser' `
            -Status     'Failed' `
            -Message    $_.Exception.Message `
            -Data       ([ordered]@{
                UserPrincipalName        = $upn
                MailboxIdentity          = ''
                MailboxDisplayName       = ''
                MailboxType              = ''
                PermissionType           = ''
                AccessSource             = ''
                GrantingGroupIdentity    = ''
                GrantingGroupDisplayName = ''
            })))
    }

    $rowNumber++
}

foreach ($result in $results) {
    Add-Member -InputObject $result -NotePropertyName 'ScopeMode' -NotePropertyValue $scopeMode -Force
}

$orderedResults = foreach ($result in $results) {
    Convert-ToOrderedReportObject -InputObject $result -PropertyOrder $reportPropertyOrder
}

Export-ResultsCsv -Results @($orderedResults) -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online mailbox access by user inventory script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
