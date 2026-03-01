#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath ("Results_B08-Create-EntraMicrosoft365Groups_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'M365.Common.psm1') -Force -DisableNameChecking

$requiredHeaders = @(
    'GroupDisplayName',
    'MailNickname',
    'Description',
    'Visibility',
    'OwnerUserPrincipalNames',
    'MemberUserPrincipalNames'
)

Write-Status -Message 'Starting Microsoft 365 group creation script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users')
Ensure-GraphConnection -RequiredScopes @('Group.ReadWrite.All', 'User.Read.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()
$userByUpn = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

function Resolve-UserByUpn {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    if ($userByUpn.ContainsKey($UserPrincipalName)) {
        return $userByUpn[$UserPrincipalName]
    }

    $escapedUpn = Escape-ODataString -Value $UserPrincipalName
    $users = @(Invoke-WithRetry -OperationName "Lookup user $UserPrincipalName" -ScriptBlock {
        Get-MgUser -Filter "userPrincipalName eq '$escapedUpn'" -ConsistencyLevel eventual -ErrorAction Stop
    })

    if ($users.Count -eq 0) {
        throw "User '$UserPrincipalName' was not found."
    }

    if ($users.Count -gt 1) {
        throw "Multiple users were returned for UPN '$UserPrincipalName'. Resolve duplicate directory objects before retrying."
    }

    $user = $users[0]
    $userByUpn[$UserPrincipalName] = $user
    return $user
}

$rowNumber = 1
foreach ($row in $rows) {
    $groupDisplayName = ([string]$row.GroupDisplayName).Trim()
    $mailNickname = ([string]$row.MailNickname).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($groupDisplayName) -or [string]::IsNullOrWhiteSpace($mailNickname)) {
            throw 'GroupDisplayName and MailNickname are required.'
        }

        $description = ([string]$row.Description).Trim()
        $visibility = ([string]$row.Visibility).Trim()
        if ([string]::IsNullOrWhiteSpace($visibility)) {
            $visibility = 'Private'
        }

        if ($visibility -notin @('Private', 'Public')) {
            throw "Visibility '$visibility' is invalid. Use Private or Public."
        }

        $ownerUpns = ConvertTo-Array -Value ([string]$row.OwnerUserPrincipalNames)
        $memberUpns = ConvertTo-Array -Value ([string]$row.MemberUserPrincipalNames)

        $escapedMailNickname = Escape-ODataString -Value $mailNickname
        $existingGroupsByAlias = @(Invoke-WithRetry -OperationName "Lookup Microsoft 365 group alias $mailNickname" -ScriptBlock {
            Get-MgGroup -Filter "mailNickname eq '$escapedMailNickname'" -ConsistencyLevel eventual -ErrorAction Stop
        })

        if ($existingGroupsByAlias.Count -gt 1) {
            throw "Multiple groups found with mailNickname '$mailNickname'. Resolve duplicate aliases before running this script."
        }

        if ($existingGroupsByAlias.Count -eq 1) {
            $existingGroup = Invoke-WithRetry -OperationName "Load group details for alias $mailNickname" -ScriptBlock {
                Get-MgGroup -GroupId $existingGroupsByAlias[0].Id -Property 'id,displayName,mailNickname,groupTypes,securityEnabled,mailEnabled' -ErrorAction Stop
            }

            $existingGroupTypes = @($existingGroup.GroupTypes)
            $isMicrosoft365Group = ($existingGroupTypes -contains 'Unified') -and ($existingGroup.MailEnabled -eq $true) -and ($existingGroup.SecurityEnabled -eq $false)
            if (-not $isMicrosoft365Group) {
                throw "A group with mailNickname '$mailNickname' already exists but is not a Microsoft 365 group."
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'CreateMicrosoft365Group' -Status 'Skipped' -Message 'Microsoft 365 group already exists.'))
            $rowNumber++
            continue
        }

        $resolvedOwners = [System.Collections.Generic.List[object]]::new()
        $seenOwnerIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($ownerUpn in $ownerUpns) {
            $ownerUser = Resolve-UserByUpn -UserPrincipalName $ownerUpn
            $ownerId = ([string]$ownerUser.Id).Trim()
            if ([string]::IsNullOrWhiteSpace($ownerId)) {
                throw "Owner '$ownerUpn' has an empty Id value."
            }

            if ($seenOwnerIds.Add($ownerId)) {
                $resolvedOwners.Add($ownerUser)
            }
        }

        $resolvedMembers = [System.Collections.Generic.List[object]]::new()
        $seenMemberIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($memberUpn in $memberUpns) {
            $memberUser = Resolve-UserByUpn -UserPrincipalName $memberUpn
            $memberId = ([string]$memberUser.Id).Trim()
            if ([string]::IsNullOrWhiteSpace($memberId)) {
                throw "Member '$memberUpn' has an empty Id value."
            }

            if ($seenMemberIds.Add($memberId)) {
                $resolvedMembers.Add($memberUser)
            }
        }

        $body = @{
            displayName     = $groupDisplayName
            mailEnabled     = $true
            mailNickname    = $mailNickname
            securityEnabled = $false
            groupTypes      = @('Unified')
            visibility      = $visibility
        }

        if (-not [string]::IsNullOrWhiteSpace($description)) {
            $body.description = $description
        }

        if ($PSCmdlet.ShouldProcess($groupDisplayName, 'Create Microsoft 365 group')) {
            $createdGroup = Invoke-WithRetry -OperationName "Create Microsoft 365 group $groupDisplayName" -ScriptBlock {
                New-MgGroup -BodyParameter $body -ErrorAction Stop
            }

            $messages = [System.Collections.Generic.List[string]]::new()
            $rowHadError = $false

            if ($resolvedOwners.Count -gt 0) {
                try {
                    $existingOwnerIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                    $existingOwners = @(Invoke-WithRetry -OperationName "Load owners for $groupDisplayName" -ScriptBlock {
                        Get-MgGroupOwner -GroupId $createdGroup.Id -All -ErrorAction Stop
                    })

                    foreach ($existingOwner in $existingOwners) {
                        $existingOwnerId = ([string]$existingOwner.Id).Trim()
                        if (-not [string]::IsNullOrWhiteSpace($existingOwnerId)) {
                            $null = $existingOwnerIds.Add($existingOwnerId)
                        }
                    }

                    foreach ($ownerUser in $resolvedOwners) {
                        $ownerUserId = ([string]$ownerUser.Id).Trim()
                        $ownerUserUpn = ([string]$ownerUser.UserPrincipalName).Trim()
                        if ($existingOwnerIds.Contains($ownerUserId)) {
                            $messages.Add("Owner '$ownerUserUpn' already present (skipped).")
                            continue
                        }

                        if ($PSCmdlet.ShouldProcess("$groupDisplayName -> $ownerUserUpn", 'Add Microsoft 365 group owner')) {
                            $ownerRef = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$ownerUserId" }
                            Invoke-WithRetry -OperationName "Add owner $ownerUserUpn to $groupDisplayName" -ScriptBlock {
                                New-MgGroupOwnerByRef -GroupId $createdGroup.Id -BodyParameter $ownerRef -ErrorAction Stop
                            }
                            $null = $existingOwnerIds.Add($ownerUserId)
                            $messages.Add("Owner '$ownerUserUpn' added.")
                        }
                        else {
                            $messages.Add("Owner '$ownerUserUpn' skipped due to WhatIf.")
                        }
                    }
                }
                catch {
                    $rowHadError = $true
                    $messages.Add("Owner assignment failed ($($_.Exception.Message)).")
                }
            }

            if ($resolvedMembers.Count -gt 0) {
                try {
                    $existingMemberIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                    $existingMembers = @(Invoke-WithRetry -OperationName "Load members for $groupDisplayName" -ScriptBlock {
                        Get-MgGroupMember -GroupId $createdGroup.Id -All -ErrorAction Stop
                    })

                    foreach ($existingMember in $existingMembers) {
                        $existingMemberId = ([string]$existingMember.Id).Trim()
                        if (-not [string]::IsNullOrWhiteSpace($existingMemberId)) {
                            $null = $existingMemberIds.Add($existingMemberId)
                        }
                    }

                    foreach ($memberUser in $resolvedMembers) {
                        $memberUserId = ([string]$memberUser.Id).Trim()
                        $memberUserUpn = ([string]$memberUser.UserPrincipalName).Trim()
                        if ($existingMemberIds.Contains($memberUserId)) {
                            $messages.Add("Member '$memberUserUpn' already present (skipped).")
                            continue
                        }

                        if ($PSCmdlet.ShouldProcess("$groupDisplayName -> $memberUserUpn", 'Add Microsoft 365 group member')) {
                            $memberRef = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$memberUserId" }
                            Invoke-WithRetry -OperationName "Add member $memberUserUpn to $groupDisplayName" -ScriptBlock {
                                New-MgGroupMemberByRef -GroupId $createdGroup.Id -BodyParameter $memberRef -ErrorAction Stop
                            }
                            $null = $existingMemberIds.Add($memberUserId)
                            $messages.Add("Member '$memberUserUpn' added.")
                        }
                        else {
                            $messages.Add("Member '$memberUserUpn' skipped due to WhatIf.")
                        }
                    }
                }
                catch {
                    $rowHadError = $true
                    $messages.Add("Member assignment failed ($($_.Exception.Message)).")
                }
            }

            $status = if ($rowHadError) { 'CreatedWithErrors' } else { 'Created' }
            if ($messages.Count -eq 0) {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'CreateMicrosoft365Group' -Status $status -Message 'Microsoft 365 group created successfully.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'CreateMicrosoft365Group' -Status $status -Message ("Microsoft 365 group created successfully. {0}" -f ($messages -join ' '))))
            }
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'CreateMicrosoft365Group' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($groupDisplayName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $groupDisplayName -Action 'CreateMicrosoft365Group' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Microsoft 365 group creation script completed.' -Level SUCCESS

