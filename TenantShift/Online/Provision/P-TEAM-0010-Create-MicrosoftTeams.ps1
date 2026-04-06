<#
.LICENSE
MIT License
Copyright (c) 2014–2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260301-004416

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
Microsoft.Graph.Authentication
Microsoft.Graph.Groups
Microsoft.Graph.Users
Microsoft.Graph.Teams

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)

.SYNOPSIS
    Provisions MicrosoftTeams in Microsoft 365.

.DESCRIPTION
    Creates MicrosoftTeams in Microsoft 365 based on records provided in the input CSV file.
    Each row in the input file corresponds to one provisioning operation. Results are written
    to the output CSV, one row per processed record, with a Status column indicating success
    or failure.
    Supports -WhatIf for dry-run validation before committing changes.
.PARAMETER InputCsvPath
    Path to the input CSV file. Each row must include the required fields documented in the .NOTES section.

.PARAMETER OutputCsvPath
    Path for the results CSV output file. Defaults to a timestamped file in a sub-folder of the script directory.


.EXAMPLE
    .\SM-P3309-Create-MicrosoftTeams.ps1 -InputCsvPath .\3309.input.csv

    Process all records in the input CSV file.

.EXAMPLE
    .\SM-P3309-Create-MicrosoftTeams.ps1 -InputCsvPath .\3309.input.csv -WhatIf

    Dry-run: reports what would change without making any modifications.

.NOTES
    Version:          1.0
    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users, Microsoft.Graph.Teams
    Required roles:   Global Administrator or appropriate workload-specific role
    Limitations:      None known.

    CSV Fields:
    Column                             Type      Required  Description
    ---------------------------------  ----      --------  -----------
    TeamDisplayName                    String    Yes       <fill in description>
    MailNickname                       String    Yes       <fill in description>
    Description                        String    Yes       <fill in description>
    Visibility                         String    Yes       <fill in description>
    OwnerUserPrincipalNames            String    Yes       <fill in description>
    MemberUserPrincipalNames           String    Yes       <fill in description>
    AllowCreateUpdateChannels          String    Yes       <fill in description>
    AllowDeleteChannels                String    Yes       <fill in description>
    AllowAddRemoveApps                 String    Yes       <fill in description>
    AllowCreateUpdateRemoveTabs        String    Yes       <fill in description>
    AllowCreateUpdateRemoveConnectors  String    Yes       <fill in description>
    AllowUserEditMessages              String    Yes       <fill in description>
    AllowUserDeleteMessages            String    Yes       <fill in description>
    AllowOwnerDeleteMessages           String    Yes       <fill in description>
    AllowTeamMentions                  String    Yes       <fill in description>
    AllowChannelMentions               String    Yes       <fill in description>
    AllowGiphy                         String    Yes       <fill in description>
    GiphyContentRating                 String    Yes       <fill in description>
    AllowStickersAndMemes              String    Yes       <fill in description>
    AllowCustomMemes                   String    Yes       <fill in description>
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Provision_OutputCsvPath') -ChildPath ("Results_SM-P3309-Create-MicrosoftTeams_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {


function Resolve-UserByUpn {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [System.Collections.Generic.Dictionary[string, object]]$UserCache
    )

    if ($UserCache.ContainsKey($UserPrincipalName)) {
        return $UserCache[$UserPrincipalName]
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
    $UserCache[$UserPrincipalName] = $user
    return $user
}

function New-TeamFromGroupWithPropagationRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        [hashtable]$TeamBody,

        [Parameter(Mandatory)]
        [string]$OperationName
    )

    $maxAttempts = 6
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            Invoke-WithRetry -OperationName $OperationName -ScriptBlock {
                New-MgGroupTeam -GroupId $GroupId -BodyParameter $TeamBody -ErrorAction Stop | Out-Null
            }
            return
        }
        catch {
            $message = [string]$_.Exception.Message
            $messageLower = $message.ToLowerInvariant()
            $likelyPropagationDelay = $messageLower -match 'not found|resource.*not exist|does not exist|directory object|replica|replication'

            if ($attempt -ge $maxAttempts -or -not $likelyPropagationDelay) {
                throw
            }

            $delaySeconds = [Math]::Min($attempt * 5, 30)
            Write-Status -Level WARN -Message "Team creation is waiting for directory replication (attempt $attempt/$maxAttempts). Retrying in $delaySeconds second(s)."
            Start-Sleep -Seconds $delaySeconds
        }
    }
}

$requiredHeaders = @(
    'TeamDisplayName',
    'MailNickname',
    'Description',
    'Visibility',
    'OwnerUserPrincipalNames',
    'MemberUserPrincipalNames',
    'AllowCreateUpdateChannels',
    'AllowDeleteChannels',
    'AllowAddRemoveApps',
    'AllowCreateUpdateRemoveTabs',
    'AllowCreateUpdateRemoveConnectors',
    'AllowUserEditMessages',
    'AllowUserDeleteMessages',
    'AllowOwnerDeleteMessages',
    'AllowTeamMentions',
    'AllowChannelMentions',
    'AllowGiphy',
    'GiphyContentRating',
    'AllowStickersAndMemes',
    'AllowCustomMemes'
)

Write-Status -Message 'Starting Microsoft Teams creation script.'
Assert-ModuleCurrent -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users', 'Microsoft.Graph.Teams')
Ensure-GraphConnection -RequiredScopes @('Group.ReadWrite.All', 'Team.Create', 'TeamSettings.ReadWrite.All', 'User.Read.All', 'Directory.Read.All')

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()
$userByUpn = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

$rowNumber = 1
foreach ($row in $rows) {
    $teamDisplayName = ([string]$row.TeamDisplayName).Trim()
    $mailNickname = ([string]$row.MailNickname).Trim()

    try {
        if ([string]::IsNullOrWhiteSpace($teamDisplayName) -or [string]::IsNullOrWhiteSpace($mailNickname)) {
            throw 'TeamDisplayName and MailNickname are required.'
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

        $allowCreateUpdateChannels = ConvertTo-Bool -Value $row.AllowCreateUpdateChannels -Default $true
        $allowDeleteChannels = ConvertTo-Bool -Value $row.AllowDeleteChannels -Default $true
        $allowAddRemoveApps = ConvertTo-Bool -Value $row.AllowAddRemoveApps -Default $true
        $allowCreateUpdateRemoveTabs = ConvertTo-Bool -Value $row.AllowCreateUpdateRemoveTabs -Default $true
        $allowCreateUpdateRemoveConnectors = ConvertTo-Bool -Value $row.AllowCreateUpdateRemoveConnectors -Default $true
        $allowUserEditMessages = ConvertTo-Bool -Value $row.AllowUserEditMessages -Default $true
        $allowUserDeleteMessages = ConvertTo-Bool -Value $row.AllowUserDeleteMessages -Default $true
        $allowOwnerDeleteMessages = ConvertTo-Bool -Value $row.AllowOwnerDeleteMessages -Default $true
        $allowTeamMentions = ConvertTo-Bool -Value $row.AllowTeamMentions -Default $true
        $allowChannelMentions = ConvertTo-Bool -Value $row.AllowChannelMentions -Default $true
        $allowGiphy = ConvertTo-Bool -Value $row.AllowGiphy -Default $true
        $allowStickersAndMemes = ConvertTo-Bool -Value $row.AllowStickersAndMemes -Default $true
        $allowCustomMemes = ConvertTo-Bool -Value $row.AllowCustomMemes -Default $true

        $giphyContentRatingRaw = ([string]$row.GiphyContentRating).Trim()
        if ([string]::IsNullOrWhiteSpace($giphyContentRatingRaw)) {
            $giphyContentRatingRaw = 'Moderate'
        }

        $giphyContentRating = switch -Regex ($giphyContentRatingRaw.ToLowerInvariant()) {
            '^strict$' { 'strict'; break }
            '^moderate$' { 'moderate'; break }
            default { throw "GiphyContentRating '$giphyContentRatingRaw' is invalid. Use Strict or Moderate." }
        }

        $escapedMailNickname = Escape-ODataString -Value $mailNickname
        $groupsByAlias = @(Invoke-WithRetry -OperationName "Lookup Microsoft 365 group alias $mailNickname" -ScriptBlock {
            Get-MgGroup -Filter "mailNickname eq '$escapedMailNickname'" -ConsistencyLevel eventual -ErrorAction Stop
        })

        if ($groupsByAlias.Count -gt 1) {
            throw "Multiple groups found with mailNickname '$mailNickname'. Resolve duplicate aliases before running this script."
        }

        $group = $null
        $groupWasCreated = $false

        if ($groupsByAlias.Count -eq 1) {
            $group = Invoke-WithRetry -OperationName "Load group details for alias $mailNickname" -ScriptBlock {
                Get-MgGroup -GroupId $groupsByAlias[0].Id -Property 'id,displayName,mailNickname,groupTypes,securityEnabled,mailEnabled' -ErrorAction Stop
            }

            $groupTypes = @($group.GroupTypes)
            $isMicrosoft365Group = ($groupTypes -contains 'Unified') -and ($group.MailEnabled -eq $true) -and ($group.SecurityEnabled -eq $false)
            if (-not $isMicrosoft365Group) {
                throw "A group with mailNickname '$mailNickname' already exists but is not a Microsoft 365 group."
            }
        }

        $existingTeam = $null
        if ($group) {
            $existingTeam = Invoke-WithRetry -OperationName "Check existing Team for group $mailNickname" -ScriptBlock {
                Get-MgGroupTeam -GroupId $group.Id -ErrorAction SilentlyContinue
            }
        }

        if ($existingTeam) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $teamDisplayName -Action 'CreateMicrosoftTeam' -Status 'Skipped' -Message 'Microsoft Team already exists for the group alias.'))
            $rowNumber++
            continue
        }

        if ($PSCmdlet.ShouldProcess($teamDisplayName, 'Create Microsoft Team')) {
            if (-not $group) {
                $groupBody = @{
                    displayName     = $teamDisplayName
                    description     = $description
                    mailEnabled     = $true
                    mailNickname    = $mailNickname
                    securityEnabled = $false
                    groupTypes      = @('Unified')
                    visibility      = $visibility
                }

                if ([string]::IsNullOrWhiteSpace($description)) {
                    $groupBody.Remove('description') | Out-Null
                }

                $group = Invoke-WithRetry -OperationName "Create Microsoft 365 group $teamDisplayName" -ScriptBlock {
                    New-MgGroup -BodyParameter $groupBody -ErrorAction Stop
                }
                $groupWasCreated = $true
            }

            $ownersAdded = 0
            $membersAdded = 0

            if ($ownerUpns.Count -gt 0) {
                $existingOwnerIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $existingOwners = @(Invoke-WithRetry -OperationName "Load owners for $teamDisplayName" -ScriptBlock {
                    Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction Stop
                })
                foreach ($owner in $existingOwners) {
                    $ownerId = ([string]$owner.Id).Trim()
                    if (-not [string]::IsNullOrWhiteSpace($ownerId)) {
                        $null = $existingOwnerIds.Add($ownerId)
                    }
                }

                foreach ($ownerUpn in $ownerUpns) {
                    $ownerUser = Resolve-UserByUpn -UserPrincipalName $ownerUpn -UserCache $userByUpn
                    $ownerUserId = ([string]$ownerUser.Id).Trim()
                    if ([string]::IsNullOrWhiteSpace($ownerUserId) -or $existingOwnerIds.Contains($ownerUserId)) {
                        continue
                    }

                    $ownerRef = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$ownerUserId" }
                    Invoke-WithRetry -OperationName "Add owner $ownerUpn to $teamDisplayName" -ScriptBlock {
                        New-MgGroupOwnerByRef -GroupId $group.Id -BodyParameter $ownerRef -ErrorAction Stop
                    }
                    $null = $existingOwnerIds.Add($ownerUserId)
                    $ownersAdded++
                }
            }

            if ($memberUpns.Count -gt 0) {
                $existingMemberIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $existingMembers = @(Invoke-WithRetry -OperationName "Load members for $teamDisplayName" -ScriptBlock {
                    Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
                })
                foreach ($member in $existingMembers) {
                    $memberId = ([string]$member.Id).Trim()
                    if (-not [string]::IsNullOrWhiteSpace($memberId)) {
                        $null = $existingMemberIds.Add($memberId)
                    }
                }

                foreach ($memberUpn in $memberUpns) {
                    $memberUser = Resolve-UserByUpn -UserPrincipalName $memberUpn -UserCache $userByUpn
                    $memberUserId = ([string]$memberUser.Id).Trim()
                    if ([string]::IsNullOrWhiteSpace($memberUserId) -or $existingMemberIds.Contains($memberUserId)) {
                        continue
                    }

                    $memberRef = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$memberUserId" }
                    Invoke-WithRetry -OperationName "Add member $memberUpn to $teamDisplayName" -ScriptBlock {
                        New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter $memberRef -ErrorAction Stop
                    }
                    $null = $existingMemberIds.Add($memberUserId)
                    $membersAdded++
                }
            }

            $teamBody = @{
                memberSettings = @{
                    allowCreateUpdateChannels         = $allowCreateUpdateChannels
                    allowDeleteChannels               = $allowDeleteChannels
                    allowAddRemoveApps                = $allowAddRemoveApps
                    allowCreateUpdateRemoveTabs       = $allowCreateUpdateRemoveTabs
                    allowCreateUpdateRemoveConnectors = $allowCreateUpdateRemoveConnectors
                }
                messagingSettings = @{
                    allowUserEditMessages   = $allowUserEditMessages
                    allowUserDeleteMessages = $allowUserDeleteMessages
                    allowOwnerDeleteMessages = $allowOwnerDeleteMessages
                    allowTeamMentions       = $allowTeamMentions
                    allowChannelMentions    = $allowChannelMentions
                }
                funSettings = @{
                    allowGiphy             = $allowGiphy
                    giphyContentRating     = $giphyContentRating
                    allowStickersAndMemes  = $allowStickersAndMemes
                    allowCustomMemes       = $allowCustomMemes
                }
            }

            New-TeamFromGroupWithPropagationRetry -GroupId $group.Id -TeamBody $teamBody -OperationName "Create Team for $teamDisplayName"

            $messageParts = [System.Collections.Generic.List[string]]::new()
            if ($groupWasCreated) {
                $messageParts.Add('Created backing Microsoft 365 group.')
            }
            else {
                $messageParts.Add('Used existing Microsoft 365 group.')
            }
            $messageParts.Add("Owners added: $ownersAdded.")
            $messageParts.Add("Members added: $membersAdded.")
            $messageParts.Add('Microsoft Team created successfully.')

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $teamDisplayName -Action 'CreateMicrosoftTeam' -Status 'Created' -Message ($messageParts -join ' ')))
        }
        else {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $teamDisplayName -Action 'CreateMicrosoftTeam' -Status 'WhatIf' -Message 'Creation skipped due to WhatIf.'))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($teamDisplayName) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $teamDisplayName -Action 'CreateMicrosoftTeam' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Microsoft Teams creation script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}







