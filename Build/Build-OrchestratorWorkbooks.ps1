#Requires -Version 7.0

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$rootDefault = 'C:\temp\EntraID_EXO_Basic'
$pwshDefault = 'pwsh'
$spAdminDefault = 'https://contoso-admin.sharepoint.com'

$provisionRows = @(
    @{ Id='P3001'; Script='Online\Provision\P3001-Create-EntraUsers.ps1'; Input='Online\Provision\P3001-Create-EntraUsers.input.csv'; NeedsSp=$false; Notes='Create Entra users' },
    @{ Id='P3002'; Script='Online\Provision\P3002-Invite-EntraGuestUsers.ps1'; Input='Online\Provision\P3002-Invite-EntraGuestUsers.input.csv'; NeedsSp=$false; Notes='Invite guest users' },
    @{ Id='P3005'; Script='Online\Provision\P3005-Create-EntraAssignedSecurityGroups.ps1'; Input='Online\Provision\P3005-Create-EntraAssignedSecurityGroups.input.csv'; NeedsSp=$false; Notes='Create assigned security groups' },
    @{ Id='P3006'; Script='Online\Provision\P3006-Create-EntraDynamicUserSecurityGroups.ps1'; Input='Online\Provision\P3006-Create-EntraDynamicUserSecurityGroups.input.csv'; NeedsSp=$false; Notes='Create dynamic user security groups' },
    @{ Id='P3008'; Script='Online\Provision\P3008-Create-EntraMicrosoft365Groups.ps1'; Input='Online\Provision\P3008-Create-EntraMicrosoft365Groups.input.csv'; NeedsSp=$false; Notes='Create Microsoft 365 groups' },
    @{ Id='P3113'; Script='Online\Provision\P3113-Create-ExchangeOnlineMailContacts.ps1'; Input='Online\Provision\P3113-Create-ExchangeOnlineMailContacts.input.csv'; NeedsSp=$false; Notes='Create mail contacts' },
    @{ Id='P3114'; Script='Online\Provision\P3114-Create-ExchangeOnlineDistributionLists.ps1'; Input='Online\Provision\P3114-Create-ExchangeOnlineDistributionLists.input.csv'; NeedsSp=$false; Notes='Create distribution lists' },
    @{ Id='P3115'; Script='Online\Provision\P3115-Create-ExchangeOnlineMailEnabledSecurityGroups.ps1'; Input='Online\Provision\P3115-Create-ExchangeOnlineMailEnabledSecurityGroups.input.csv'; NeedsSp=$false; Notes='Create mail-enabled security groups' },
    @{ Id='P3116'; Script='Online\Provision\P3116-Create-ExchangeOnlineSharedMailboxes.ps1'; Input='Online\Provision\P3116-Create-ExchangeOnlineSharedMailboxes.input.csv'; NeedsSp=$false; Notes='Create shared mailboxes' },
    @{ Id='P3118'; Script='Online\Provision\P3118-Create-ExchangeOnlineResourceMailboxes.ps1'; Input='Online\Provision\P3118-Create-ExchangeOnlineResourceMailboxes.input.csv'; NeedsSp=$false; Notes='Create resource mailboxes' },
    @{ Id='P3119'; Script='Online\Provision\P3119-Create-ExchangeOnlineDynamicDistributionGroups.ps1'; Input='Online\Provision\P3119-Create-ExchangeOnlineDynamicDistributionGroups.input.csv'; NeedsSp=$false; Notes='Create dynamic distribution groups' },
    @{ Id='P3240'; Script='Online\Provision\P3240-Create-SharePointSites.ps1'; Input='Online\Provision\P3240-Create-SharePointSites.input.csv'; NeedsSp=$true; Notes='Create SharePoint sites' },
    @{ Id='P3242'; Script='Online\Provision\P3242-Create-SharePointHubSites.ps1'; Input='Online\Provision\P3242-Create-SharePointHubSites.input.csv'; NeedsSp=$true; Notes='Register hub sites' },
    @{ Id='P3309'; Script='Online\Provision\P3309-Create-MicrosoftTeams.ps1'; Input='Online\Provision\P3309-Create-MicrosoftTeams.input.csv'; NeedsSp=$false; Notes='Create Teams' }
)

$modifyRows = @(
    @{ Id='M3003'; Script='Online\Modify\M3003-Set-EntraUserLicenses.ps1'; Input='Online\Modify\M3003-Set-EntraUserLicenses.input.csv'; NeedsSp=$false; Notes='Assign user licenses' },
    @{ Id='M3007'; Script='Online\Modify\M3007-Set-EntraSecurityGroupMembers.ps1'; Input='Online\Modify\M3007-Set-EntraSecurityGroupMembers.input.csv'; NeedsSp=$false; Notes='Update Entra security group membership' },
    @{ Id='M3113'; Script='Online\Modify\M3113-Update-ExchangeOnlineMailContacts.ps1'; Input='Online\Modify\M3113-Update-ExchangeOnlineMailContacts.input.csv'; NeedsSp=$false; Notes='Update mail contacts' },
    @{ Id='M3114'; Script='Online\Modify\M3114-Update-ExchangeOnlineDistributionLists.ps1'; Input='Online\Modify\M3114-Update-ExchangeOnlineDistributionLists.input.csv'; NeedsSp=$false; Notes='Update distribution lists' },
    @{ Id='M3115'; Script='Online\Modify\M3115-Set-ExchangeOnlineDistributionListMembers.ps1'; Input='Online\Modify\M3115-Set-ExchangeOnlineDistributionListMembers.input.csv'; NeedsSp=$false; Notes='Update DL membership' },
    @{ Id='M3116'; Script='Online\Modify\M3116-Update-ExchangeOnlineSharedMailboxes.ps1'; Input='Online\Modify\M3116-Update-ExchangeOnlineSharedMailboxes.input.csv'; NeedsSp=$false; Notes='Update shared mailbox settings' },
    @{ Id='M3117'; Script='Online\Modify\M3117-Set-ExchangeOnlineSharedMailboxPermissions.ps1'; Input='Online\Modify\M3117-Set-ExchangeOnlineSharedMailboxPermissions.input.csv'; NeedsSp=$false; Notes='Update shared mailbox permissions' },
    @{ Id='M3118'; Script='Online\Modify\M3118-Update-ExchangeOnlineResourceMailboxes.ps1'; Input='Online\Modify\M3118-Update-ExchangeOnlineResourceMailboxes.input.csv'; NeedsSp=$false; Notes='Update resource mailbox settings' },
    @{ Id='M3119'; Script='Online\Modify\M3119-Set-ExchangeOnlineResourceMailboxBookingDelegates.ps1'; Input='Online\Modify\M3119-Set-ExchangeOnlineResourceMailboxBookingDelegates.input.csv'; NeedsSp=$false; Notes='Update resource booking delegates' },
    @{ Id='M3120'; Script='Online\Modify\M3120-Set-ExchangeOnlineMailboxDelegations.ps1'; Input='Online\Modify\M3120-Set-ExchangeOnlineMailboxDelegations.input.csv'; NeedsSp=$false; Notes='Update mailbox delegations' },
    @{ Id='M3121'; Script='Online\Modify\M3121-Set-ExchangeOnlineMailboxFolderPermissions.ps1'; Input='Online\Modify\M3121-Set-ExchangeOnlineMailboxFolderPermissions.input.csv'; NeedsSp=$false; Notes='Update mailbox folder permissions' },
    @{ Id='M3122'; Script='Online\Modify\M3122-Update-ExchangeOnlineMailEnabledSecurityGroups.ps1'; Input='Online\Modify\M3122-Update-ExchangeOnlineMailEnabledSecurityGroups.input.csv'; NeedsSp=$false; Notes='Update mail-enabled security groups' },
    @{ Id='M3123'; Script='Online\Modify\M3123-Update-ExchangeOnlineDynamicDistributionGroups.ps1'; Input='Online\Modify\M3123-Update-ExchangeOnlineDynamicDistributionGroups.input.csv'; NeedsSp=$false; Notes='Update dynamic distribution groups' },
    @{ Id='M3204'; Script='Online\Modify\M3204-PreProvision-OneDrive.ps1'; Input='Online\Modify\M3204-PreProvision-OneDrive.input.csv'; NeedsSp=$true; Notes='Pre-provision OneDrive sites' },
    @{ Id='M3205'; Script='Online\Modify\M3205-Set-OneDriveStorageQuota.ps1'; Input='Online\Modify\M3205-Set-OneDriveStorageQuota.input.csv'; NeedsSp=$true; Notes='Update OneDrive storage quotas' },
    @{ Id='M3206'; Script='Online\Modify\M3206-Set-OneDriveSiteCollectionAdmins.ps1'; Input='Online\Modify\M3206-Set-OneDriveSiteCollectionAdmins.input.csv'; NeedsSp=$true; Notes='Update OneDrive site collection admins' },
    @{ Id='M3241'; Script='Online\Modify\M3241-Set-SharePointSiteAdmins.ps1'; Input='Online\Modify\M3241-Set-SharePointSiteAdmins.input.csv'; NeedsSp=$true; Notes='Update SharePoint site admins' },
    @{ Id='M3243'; Script='Online\Modify\M3243-Associate-SharePointSitesToHub.ps1'; Input='Online\Modify\M3243-Associate-SharePointSitesToHub.input.csv'; NeedsSp=$true; Notes='Associate sites to hub' },
    @{ Id='M3310'; Script='Online\Modify\M3310-Set-MicrosoftTeamMembers.ps1'; Input='Online\Modify\M3310-Set-MicrosoftTeamMembers.input.csv'; NeedsSp=$false; Notes='Update Team members' },
    @{ Id='M3311'; Script='Online\Modify\M3311-Update-MicrosoftTeamChannels.ps1'; Input='Online\Modify\M3311-Update-MicrosoftTeamChannels.input.csv'; NeedsSp=$false; Notes='Update Team channels' },
    @{ Id='M3312'; Script='Online\Modify\M3312-Set-MicrosoftTeamChannelMembers.ps1'; Input='Online\Modify\M3312-Set-MicrosoftTeamChannelMembers.input.csv'; NeedsSp=$false; Notes='Update Team channel members' }
)

$irRows = @(
    @{ Id='IR3001'; Script='Online\InventoryAndReport\IR3001-Get-EntraUsers.ps1'; Input='Online\InventoryAndReport\Scope-Users.input.csv'; NeedsSp=$false; Notes='Inventory users' },
    @{ Id='IR3002'; Script='Online\InventoryAndReport\IR3002-Get-EntraGuestUsers.ps1'; Input='Online\InventoryAndReport\Scope-GuestUsers.input.csv'; NeedsSp=$false; Notes='Inventory guest users' },
    @{ Id='IR3003'; Script='Online\InventoryAndReport\IR3003-Get-EntraUserLicenses.ps1'; Input='Online\InventoryAndReport\Scope-Users.input.csv'; NeedsSp=$false; Notes='Inventory user licenses' },
    @{ Id='IR3204'; Script='Online\InventoryAndReport\IR3204-Get-OneDriveProvisioningStatus.ps1'; Input='Online\InventoryAndReport\Scope-Users.input.csv'; NeedsSp=$true; Notes='Inventory OneDrive URL/provisioning status by user' },
    @{ Id='IR3205'; Script='Online\InventoryAndReport\IR3205-Get-OneDriveStorageAndQuota.ps1'; Input='Online\InventoryAndReport\Scope-Users.input.csv'; NeedsSp=$true; Notes='Inventory OneDrive storage and quota by user' },
    @{ Id='IR3206'; Script='Online\InventoryAndReport\IR3206-Get-OneDriveSiteCollectionAdmins.ps1'; Input='Online\InventoryAndReport\Scope-Users.input.csv'; NeedsSp=$true; Notes='Inventory OneDrive site collection admins by user' },
    @{ Id='IR3005'; Script='Online\InventoryAndReport\IR3005-Get-EntraSecurityGroups.ps1'; Input='Online\InventoryAndReport\Scope-EntraSecurityGroups.input.csv'; NeedsSp=$false; Notes='Inventory assigned security groups' },
    @{ Id='IR3006'; Script='Online\InventoryAndReport\IR3006-Get-EntraDynamicUserSecurityGroups.ps1'; Input='Online\InventoryAndReport\Scope-EntraDynamicUserSecurityGroups.input.csv'; NeedsSp=$false; Notes='Inventory dynamic user security groups' },
    @{ Id='IR3007'; Script='Online\InventoryAndReport\IR3007-Get-EntraSecurityGroupMembers.ps1'; Input='Online\InventoryAndReport\Scope-EntraSecurityGroups.input.csv'; NeedsSp=$false; Notes='Inventory security group members' },
    @{ Id='IR3008'; Script='Online\InventoryAndReport\IR3008-Get-EntraMicrosoft365Groups.ps1'; Input='Online\InventoryAndReport\Scope-M365Groups.input.csv'; NeedsSp=$false; Notes='Inventory M365 groups' },
    @{ Id='IR3113'; Script='Online\InventoryAndReport\IR3113-Get-ExchangeOnlineMailContacts.ps1'; Input='Online\InventoryAndReport\Scope-MailContacts.input.csv'; NeedsSp=$false; Notes='Inventory mail contacts' },
    @{ Id='IR3114'; Script='Online\InventoryAndReport\IR3114-Get-ExchangeOnlineDistributionLists.ps1'; Input='Online\InventoryAndReport\Scope-DistributionLists.input.csv'; NeedsSp=$false; Notes='Inventory distribution lists' },
    @{ Id='IR3115'; Script='Online\InventoryAndReport\IR3115-Get-ExchangeOnlineDistributionListMembers.ps1'; Input='Online\InventoryAndReport\Scope-DistributionLists.input.csv'; NeedsSp=$false; Notes='Inventory DL members' },
    @{ Id='IR3116'; Script='Online\InventoryAndReport\IR3116-Get-ExchangeOnlineSharedMailboxes.ps1'; Input='Online\InventoryAndReport\Scope-SharedMailboxes.input.csv'; NeedsSp=$false; Notes='Inventory shared mailboxes' },
    @{ Id='IR3117'; Script='Online\InventoryAndReport\IR3117-Get-ExchangeOnlineSharedMailboxPermissions.ps1'; Input='Online\InventoryAndReport\Scope-SharedMailboxes.input.csv'; NeedsSp=$false; Notes='Inventory shared mailbox permissions' },
    @{ Id='IR3118'; Script='Online\InventoryAndReport\IR3118-Get-ExchangeOnlineResourceMailboxes.ps1'; Input='Online\InventoryAndReport\Scope-ResourceMailboxes.input.csv'; NeedsSp=$false; Notes='Inventory resource mailboxes' },
    @{ Id='IR3119'; Script='Online\InventoryAndReport\IR3119-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1'; Input='Online\InventoryAndReport\Scope-ResourceMailboxes.input.csv'; NeedsSp=$false; Notes='Inventory resource booking settings' },
    @{ Id='IR3120'; Script='Online\InventoryAndReport\IR3120-Get-ExchangeOnlineMailboxDelegations.ps1'; Input='Online\InventoryAndReport\Scope-Mailboxes.input.csv'; NeedsSp=$false; Notes='Inventory mailbox delegations' },
    @{ Id='IR3121'; Script='Online\InventoryAndReport\IR3121-Get-ExchangeOnlineMailboxFolderPermissions.ps1'; Input='Online\InventoryAndReport\Scope-Mailboxes.input.csv'; NeedsSp=$false; Notes='Inventory mailbox folder permissions' },
    @{ Id='IR3122'; Script='Online\InventoryAndReport\IR3122-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1'; Input='Online\InventoryAndReport\Scope-MailEnabledSecurityGroups.input.csv'; NeedsSp=$false; Notes='Inventory mail-enabled security groups' },
    @{ Id='IR3123'; Script='Online\InventoryAndReport\IR3123-Get-ExchangeOnlineDynamicDistributionGroups.ps1'; Input='Online\InventoryAndReport\Scope-DynamicDistributionGroups.input.csv'; NeedsSp=$false; Notes='Inventory dynamic distribution groups' },
    @{ Id='IR3240'; Script='Online\InventoryAndReport\IR3240-Get-SharePointSites.ps1'; Input='Online\InventoryAndReport\Scope-SharePointSites.input.csv'; NeedsSp=$true; Notes='Inventory SharePoint sites' },
    @{ Id='IR3309'; Script='Online\InventoryAndReport\IR3309-Get-MicrosoftTeams.ps1'; Input='Online\InventoryAndReport\Scope-Teams.input.csv'; NeedsSp=$false; Notes='Inventory Teams' },
    @{ Id='IR3310'; Script='Online\InventoryAndReport\IR3310-Get-MicrosoftTeamMembers.ps1'; Input='Online\InventoryAndReport\Scope-Teams.input.csv'; NeedsSp=$false; Notes='Inventory Team members' },
    @{ Id='IR3311'; Script='Online\InventoryAndReport\IR3311-Get-MicrosoftTeamChannels.ps1'; Input='Online\InventoryAndReport\Scope-Teams.input.csv'; NeedsSp=$false; Notes='Inventory Team channels' }
)

function New-OrchestratorWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Excel,

        [Parameter(Mandatory)]
        [string]$OutputPath,

        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter(Mandatory)]
        [object[]]$Rows,

        [Parameter(Mandatory)]
        [bool]$DefaultWhatIf
    )

    $workbook = $Excel.Workbooks.Add()
    $configSheet = $null
    $commandsSheet = $null

    try {
        $configSheet = $workbook.Worksheets.Item(1)
        if ($workbook.Worksheets.Count -lt 2) {
            $commandsSheet = $workbook.Worksheets.Add()
        }
        else {
            $commandsSheet = $workbook.Worksheets.Item(2)
        }

        $configSheet.Name = 'Config'
        $commandsSheet.Name = 'Commands'

        while ($workbook.Worksheets.Count -gt 2) {
            $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
        }

        $configSheet.Cells.Item(1, 1).Value2 = 'Setting'
        $configSheet.Cells.Item(1, 2).Value2 = 'Value'
        $configSheet.Cells.Item(1, 3).Value2 = 'Notes'

        $configRows = @(
            @('WorkbookTitle', $Title, 'Read-only label for this workbook'),
            @('RootPath', $rootDefault, 'Change this to where the repo was extracted, for example C:\temp\EntraID_EXO_Basic'),
            @('PwshExe', $pwshDefault, 'pwsh or full path to pwsh.exe'),
            @('SharePointAdminUrl', $spAdminDefault, 'Used only when NeedsSharePointAdminUrl is TRUE'),
            @('UseWhatIfDefault', $(if ($DefaultWhatIf) { 'TRUE' } else { 'FALSE' }), 'Default for UseWhatIf column; can be changed per row')
        )

        $configRowIndex = 2
        foreach ($row in $configRows) {
            $configSheet.Cells.Item($configRowIndex, 1).Value2 = $row[0]
            $configSheet.Cells.Item($configRowIndex, 2).Value2 = $row[1]
            $configSheet.Cells.Item($configRowIndex, 3).Value2 = $row[2]
            $configRowIndex++
        }

        $headers = @(
            'Enabled',
            'ScriptId',
            'ScriptRelativePath',
            'InputCsvRelativePath',
            'OutputCsvPathOverride',
            'NeedsSharePointAdminUrl',
            'UseWhatIf',
            'AdditionalArgs',
            'Notes',
            'ReadyState',
            'Command'
        )

        for ($i = 0; $i -lt $headers.Count; $i++) {
            $commandsSheet.Cells.Item(1, $i + 1).Value2 = $headers[$i]
        }

        $rowIndex = 2
        foreach ($row in $Rows) {
            $commandsSheet.Cells.Item($rowIndex, 1).Formula = '=TRUE'
            $commandsSheet.Cells.Item($rowIndex, 2).Value2 = $row.Id
            $commandsSheet.Cells.Item($rowIndex, 3).Value2 = $row.Script
            $commandsSheet.Cells.Item($rowIndex, 4).Value2 = $row.Input
            $commandsSheet.Cells.Item($rowIndex, 5).Value2 = ''
            $commandsSheet.Cells.Item($rowIndex, 6).Formula = $(if ($row.NeedsSp) { '=TRUE' } else { '=FALSE' })
            $commandsSheet.Cells.Item($rowIndex, 7).Formula = '=UPPER(Config!$B$6)="TRUE"'
            $commandsSheet.Cells.Item($rowIndex, 8).Value2 = ''
            $commandsSheet.Cells.Item($rowIndex, 9).Value2 = $row.Notes
            $commandsSheet.Cells.Item($rowIndex, 10).Formula = '=IF(A' + $rowIndex + '<>TRUE,"Skip",IF(OR(C' + $rowIndex + '="",D' + $rowIndex + '="",AND(F' + $rowIndex + '=TRUE,Config!$B$5="")),"Missing required","Ready"))'

            $commandFormula = '=IF(A' + $rowIndex + '<>TRUE,"",TRIM(Config!$B$4 & " -File """ & Config!$B$3 & "\" & C' + $rowIndex + ' & """ -InputCsvPath """ & Config!$B$3 & "\" & D' + $rowIndex + ' & """" & IF(E' + $rowIndex + '<>""," -OutputCsvPath """ & IF(OR(LEFT(E' + $rowIndex + ',2)="\\",MID(E' + $rowIndex + ',2,1)=":"),E' + $rowIndex + ',Config!$B$3 & "\" & E' + $rowIndex + ') & """","") & IF(F' + $rowIndex + '=TRUE," -SharePointAdminUrl """ & Config!$B$5 & """","") & IF(G' + $rowIndex + '=TRUE," -WhatIf","") & IF(H' + $rowIndex + '<>""," " & H' + $rowIndex + ',"")))'
            $commandsSheet.Cells.Item($rowIndex, 11).Formula = $commandFormula

            $rowIndex++
        }

        $configSheet.Range('A1:C1').Font.Bold = $true
        $commandsSheet.Range('A1:K1').Font.Bold = $true
        $commandsSheet.Range('A1:K1').Interior.ColorIndex = 15
        $commandsSheet.Range('J2:J' + ($rowIndex - 1)).Interior.ColorIndex = 36

        $commandsSheet.Range('A2:A' + ($rowIndex - 1)).HorizontalAlignment = -4108
        $commandsSheet.Range('F2:G' + ($rowIndex - 1)).HorizontalAlignment = -4108
        $commandsSheet.Range('K2:K' + ($rowIndex - 1)).WrapText = $true

        $commandsSheet.Range('A1:K1').AutoFilter() | Out-Null
        $commandsSheet.Application.ActiveWindow.SplitRow = 1
        $commandsSheet.Application.ActiveWindow.FreezePanes = $true

        $configSheet.Columns.AutoFit() | Out-Null
        $commandsSheet.Columns.AutoFit() | Out-Null
        $commandsSheet.Columns.Item(11).ColumnWidth = 120
        $commandsSheet.Columns.Item(9).ColumnWidth = 38

        $outputDir = Split-Path -Path $OutputPath -Parent
        if (-not (Test-Path -LiteralPath $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }

        if (Test-Path -LiteralPath $OutputPath) {
            Remove-Item -LiteralPath $OutputPath -Force
        }

        $workbook.SaveAs($OutputPath, 51)
        Write-Output "CREATED: $OutputPath"
    }
    finally {
        if ($workbook) {
            $workbook.Close($false)
        }
        if ($commandsSheet) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($commandsSheet) | Out-Null
        }
        if ($configSheet) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($configSheet) | Out-Null
        }
        if ($workbook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
    }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$repoRoot = Split-Path -Parent $PSScriptRoot

try {
    New-OrchestratorWorkbook -Excel $excel -OutputPath (Join-Path -Path $repoRoot -ChildPath 'Online\Provision\Provision-Orchestrator.xlsx') -Title 'Provision Orchestrator' -Rows $provisionRows -DefaultWhatIf $true
    New-OrchestratorWorkbook -Excel $excel -OutputPath (Join-Path -Path $repoRoot -ChildPath 'Online\Modify\Modify-Orchestrator.xlsx') -Title 'Modify Orchestrator' -Rows $modifyRows -DefaultWhatIf $true
    New-OrchestratorWorkbook -Excel $excel -OutputPath (Join-Path -Path $repoRoot -ChildPath 'Online\InventoryAndReport\InventoryAndReport-Orchestrator.xlsx') -Title 'InventoryAndReport Orchestrator' -Rows $irRows -DefaultWhatIf $false
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}





