<#
.SYNOPSIS
    Generates InventoryAndReport-Orchestrator.xlsx via Excel COM automation.
#>
[CmdletBinding()]
param(
    [string]$OutputPath = "$PSScriptRoot\InventoryAndReport-Orchestrator.xlsx"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─── Script catalog ───────────────────────────────────────────────────────────
$catalog = @(
    [pscustomobject]@{Workload='MEID';ID='D-MEID-0010';Desc='Get Entra Users';                                        File='D-MEID-0010-Get-EntraUsers.ps1';                                           InputCsv='Scope-Users.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='MEID';ID='D-MEID-0020';Desc='Get Entra Guest Users';                                  File='D-MEID-0020-Get-EntraGuestUsers.ps1';                                       InputCsv='Scope-GuestUsers.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='MEID';ID='D-MEID-0030';Desc='Get Entra User Licenses';                                File='D-MEID-0030-Get-EntraUserLicenses.ps1';                                     InputCsv='Scope-Users.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='MEID';ID='D-MEID-0060';Desc='Get Entra Privileged Roles';                             File='D-MEID-0060-Get-EntraPrivilegedRoles.ps1';                                  InputCsv='Scope-EntraPrivilegedRoles.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='MEID';ID='D-MEID-0070';Desc='Get Entra Security Groups';                              File='D-MEID-0070-Get-EntraSecurityGroups.ps1';                                   InputCsv='Scope-EntraSecurityGroups.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='MEID';ID='D-MEID-0080';Desc='Get Entra Dynamic User Security Groups';                 File='D-MEID-0080-Get-EntraDynamicUserSecurityGroups.ps1';                        InputCsv='Scope-EntraDynamicUserSecurityGroups.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='MEID';ID='D-MEID-0090';Desc='Get Entra Security Group Members';                       File='D-MEID-0090-Get-EntraSecurityGroupMembers.ps1';                             InputCsv='Scope-EntraSecurityGroups.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='MEID';ID='D-MEID-0100';Desc='Get Entra Microsoft 365 Groups';                         File='D-MEID-0100-Get-EntraMicrosoft365Groups.ps1';                               InputCsv='Scope-M365Groups.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='MEID';ID='D-MEID-0110';Desc='Get Entra Microsoft 365 Group Members';                  File='D-MEID-0110-Get-EntraMicrosoft365GroupMembers.ps1';                         InputCsv='Scope-M365Groups.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='MEID';ID='D-MEID-0120';Desc='Get Entra Microsoft 365 Group Owners';                   File='D-MEID-0120-Get-EntraMicrosoft365GroupOwners.ps1';                          InputCsv='Scope-M365Groups.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0010';Desc='Get Exchange Online Domain Verification Records';        File='D-EXOL-0010-Get-ExchangeOnlineDomainVerificationRecords.ps1';               InputCsv='Scope-AcceptedDomains.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0030';Desc='Get Exchange Online Mail Contacts';                      File='D-EXOL-0030-Get-ExchangeOnlineMailContacts.ps1';                            InputCsv='Scope-MailContacts.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0040';Desc='Get Exchange Online Distribution Lists';                 File='D-EXOL-0040-Get-ExchangeOnlineDistributionLists.ps1';                       InputCsv='Scope-DistributionLists.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0050';Desc='Get Exchange Online Mail-Enabled Security Groups';       File='D-EXOL-0050-Get-ExchangeOnlineMailEnabledSecurityGroups.ps1';               InputCsv='Scope-MailEnabledSecurityGroups.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0060';Desc='Get Exchange Online Dynamic Distribution Groups';        File='D-EXOL-0060-Get-ExchangeOnlineDynamicDistributionGroups.ps1';               InputCsv='Scope-DynamicDistributionGroups.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0070';Desc='Get Exchange Online Shared Mailboxes';                   File='D-EXOL-0070-Get-ExchangeOnlineSharedMailboxes.ps1';                         InputCsv='Scope-SharedMailboxes.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0080';Desc='Get Exchange Online Resource Mailboxes';                 File='D-EXOL-0080-Get-ExchangeOnlineResourceMailboxes.ps1';                       InputCsv='Scope-ResourceMailboxes.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0090';Desc='Get Exchange Online Distribution List Members';          File='D-EXOL-0090-Get-ExchangeOnlineDistributionListMembers.ps1';                 InputCsv='Scope-DistributionLists.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0100';Desc='Get Exchange Online Mail-Enabled Security Group Members';File='D-EXOL-0100-Get-ExchangeOnlineMailEnabledSecurityGroupMembers.ps1';         InputCsv='Scope-MailEnabledSecurityGroups.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0110';Desc='Get Exchange Online Shared Mailbox Permissions';         File='D-EXOL-0110-Get-ExchangeOnlineSharedMailboxPermissions.ps1';                InputCsv='Scope-SharedMailboxes.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0120';Desc='Get Exchange Online Resource Mailbox Booking Delegates'; File='D-EXOL-0120-Get-ExchangeOnlineResourceMailboxBookingDelegates.ps1';         InputCsv='Scope-ResourceMailboxes.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0130';Desc='Get Exchange Online Mailbox Delegations';                File='D-EXOL-0130-Get-ExchangeOnlineMailboxDelegations.ps1';                      InputCsv='Scope-Mailboxes.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0140';Desc='Get Exchange Online Mailbox Folder Permissions';         File='D-EXOL-0140-Get-ExchangeOnlineMailboxFolderPermissions.ps1';                InputCsv='Scope-Mailboxes.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0210';Desc='Get Exchange Online Recipient Type Counts';              File='D-EXOL-0210-Get-ExchangeOnlineRecipientTypeCounts.ps1';                     InputCsv='D-EXOL-0210-Get-ExchangeOnlineRecipientTypeCounts.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0220';Desc='Get Exchange Online Mailbox High-Level Stats';           File='D-EXOL-0220-Get-ExchangeOnlineMailboxHighLevelStats.ps1';                   InputCsv='Scope-Mailboxes.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0230';Desc='Get Exchange Online Mailbox Sizes';                      File='D-EXOL-0230-Get-ExchangeOnlineMailboxSizes.ps1';                            InputCsv='Scope-Mailboxes.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0240';Desc='Get Exchange Online Mailbox Stats Per Mailbox';          File='D-EXOL-0240-Get-ExchangeOnlineMailboxStatsPerMailbox.ps1';                  InputCsv='Scope-Mailboxes.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0250';Desc='Get Exchange Online Mailbox Permissions Consolidated';   File='D-EXOL-0250-Get-ExchangeOnlineMailboxPermissionsConsolidated.ps1';          InputCsv='Scope-Mailboxes.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0260';Desc='Get Exchange Online User Mailbox SMTP Addresses';        File='D-EXOL-0260-Get-ExchangeOnlineUserMailboxSmtpAddresses.ps1';                InputCsv='Scope-Mailboxes.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0270';Desc='Get Exchange Online Shared Mailbox SMTP Addresses';      File='D-EXOL-0270-Get-ExchangeOnlineSharedMailboxSmtpAddresses.ps1';              InputCsv='Scope-SharedMailboxes.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='EXOL';ID='D-EXOL-0280';Desc='Test Exchange Online Unexpected Retention Policy Tags';  File='D-EXOL-0280-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.ps1';          InputCsv='D-EXOL-0280-Test-ExchangeOnlineUnexpectedRetentionPolicyTags.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='ONDR';ID='D-ONDR-0010';Desc='Get OneDrive Provisioning Status';                       File='D-ONDR-0010-Get-OneDriveProvisioningStatus.ps1';                            InputCsv='Scope-Users.input.csv';NeedsSP='Yes'}
    [pscustomobject]@{Workload='ONDR';ID='D-ONDR-0020';Desc='Get OneDrive Storage and Quota';                         File='D-ONDR-0020-Get-OneDriveStorageAndQuota.ps1';                               InputCsv='Scope-Users.input.csv';NeedsSP='Yes'}
    [pscustomobject]@{Workload='ONDR';ID='D-ONDR-0030';Desc='Get OneDrive Site Collection Admins';                    File='D-ONDR-0030-Get-OneDriveSiteCollectionAdmins.ps1';                          InputCsv='Scope-Users.input.csv';NeedsSP='Yes'}
    [pscustomobject]@{Workload='ONDR';ID='D-ONDR-0040';Desc='Get OneDrive Sharing Settings';                          File='D-ONDR-0040-Get-OneDriveSharingSettings.ps1';                               InputCsv='Scope-Users.input.csv';NeedsSP='Yes'}
    [pscustomobject]@{Workload='ONDR';ID='D-ONDR-0050';Desc='Get OneDrive External Sharing Links';                    File='D-ONDR-0050-Get-OneDriveExternalSharingLinks.ps1';                          InputCsv='Scope-Users.input.csv';NeedsSP='Yes'}
    [pscustomobject]@{Workload='ONDR';ID='D-ONDR-0060';Desc='Get OneDrive Site Lock State';                           File='D-ONDR-0060-Get-OneDriveSiteLockState.ps1';                                 InputCsv='Scope-Users.input.csv';NeedsSP='Yes'}
    [pscustomobject]@{Workload='SPOL';ID='D-SPOL-0010';Desc='Get SharePoint Sites';                                   File='D-SPOL-0010-Get-SharePointSites.ps1';                                       InputCsv='Scope-SharePointSites.input.csv';NeedsSP='Yes'}
    [pscustomobject]@{Workload='TEAM';ID='D-TEAM-0010';Desc='Get Microsoft Teams';                                    File='D-TEAM-0010-Get-MicrosoftTeams.ps1';                                        InputCsv='Scope-Teams.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='TEAM';ID='D-TEAM-0020';Desc='Get Microsoft Team Members';                             File='D-TEAM-0020-Get-MicrosoftTeamMembers.ps1';                                  InputCsv='Scope-Teams.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='TEAM';ID='D-TEAM-0030';Desc='Get Microsoft Team Channels';                            File='D-TEAM-0030-Get-MicrosoftTeamChannels.ps1';                                 InputCsv='Scope-Teams.input.csv';NeedsSP='No'}
    [pscustomobject]@{Workload='TEAM';ID='D-TEAM-0040';Desc='Get Microsoft Team Channel Members';                     File='D-TEAM-0040-Get-MicrosoftTeamChannelMembers.ps1';                           InputCsv='D-TEAM-0040-Get-MicrosoftTeamChannelMembers.input.csv';NeedsSP='No'}
)

# ─── Color helper: RRGGBB hex string → Excel BGR decimal ──────────────────────
function clr([string]$hex) {
    $r = [Convert]::ToInt32($hex.Substring(0,2),16)
    $g = [Convert]::ToInt32($hex.Substring(2,2),16)
    $b = [Convert]::ToInt32($hex.Substring(4,2),16)
    return ($b -shl 16) + ($g -shl 8) + $r
}

$wlColors = @{
    MEID = clr 'D6E4FF'
    EXOL = clr 'D6F5D6'
    ONDR = clr 'FFF3CC'
    SPOL = clr 'FFE0CC'
    TEAM = clr 'ECD9FF'
}
$navyBg   = clr '1F3864'
$blueBg   = clr '2E75B6'
$white    = 16777215
$editBg   = clr 'FFFCD6'
$lockBg   = clr 'F2F2F2'
$darkGrey = clr '595959'
$warnBg   = clr 'FCE4D6'
$warnFg   = clr '843C0C'
$greenBg  = clr 'E2EFDA'
$greenFg  = clr '375623'
$greyText = clr 'A6A6A6'

# ─── Excel COM bootstrap ──────────────────────────────────────────────────────
$xl = New-Object -ComObject Excel.Application
$xl.Visible        = $false
$xl.DisplayAlerts  = $false
$xl.ScreenUpdating = $false

try {
    $wb = $xl.Workbooks.Add()

    # Reduce to 1 sheet, name it Config; add Scripts and RunCommands after
    while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }
    $wb.Sheets.Item(1).Name = 'Config'
    $cfgSh = $wb.Sheets.Item('Config')
    $scrSh = $wb.Sheets.Add([System.Reflection.Missing]::Value, $cfgSh)
    $scrSh.Name = 'Scripts'
    $runSh = $wb.Sheets.Add([System.Reflection.Missing]::Value, $scrSh)
    $runSh.Name = 'RunCommands'

    # ══════════════════════════════════════════════════════════════════════════
    # CONFIG SHEET
    # ══════════════════════════════════════════════════════════════════════════

    $cfgSh.Range('A1:C1').Merge()
    $cell = $cfgSh.Range('A1')
    $cell.Value2 = 'InventoryAndReport Orchestrator — Configuration'
    $cell.Font.Bold = $true; $cell.Font.Size = 14
    $cell.Interior.Color = $navyBg; $cell.Font.Color = $white
    $cfgSh.Rows.Item(1).RowHeight = 28

    # Column headers row 3 — individual assignments to avoid array flattening
    $cfgSh.Range('A3').Value2 = 'Setting';  $cfgSh.Range('A3').Font.Bold = $true; $cfgSh.Range('A3').Interior.Color = $blueBg; $cfgSh.Range('A3').Font.Color = $white
    $cfgSh.Range('B3').Value2 = 'Value';    $cfgSh.Range('B3').Font.Bold = $true; $cfgSh.Range('B3').Interior.Color = $blueBg; $cfgSh.Range('B3').Font.Color = $white
    $cfgSh.Range('C3').Value2 = 'Notes';    $cfgSh.Range('C3').Font.Bold = $true; $cfgSh.Range('C3').Interior.Color = $blueBg; $cfgSh.Range('C3').Font.Color = $white
    $cfgSh.Rows.Item(3).RowHeight = 18

    # Setting rows 4–6
    $cfgSh.Range('A4').Value2 = 'Script Root Path';     $cfgSh.Range('A4').Font.Bold = $true
    $cfgSh.Range('B4').Interior.Color = $editBg
    $cfgSh.Range('C4').Value2 = 'Repo root path to run scripts from. Use . for current directory, or an absolute path. Use forward slashes.'
    $cfgSh.Range('C4').Font.Italic = $true; $cfgSh.Range('C4').Font.Color = $darkGrey

    $cfgSh.Range('A5').Value2 = 'SharePoint Admin URL'; $cfgSh.Range('A5').Font.Bold = $true
    $cfgSh.Range('B5').Interior.Color = $editBg
    $cfgSh.Range('C5').Value2 = 'Required for ONDR and SPOL scripts. e.g. https://contoso-admin.sharepoint.com'
    $cfgSh.Range('C5').Font.Italic = $true; $cfgSh.Range('C5').Font.Color = $darkGrey

    $cfgSh.Range('A6').Value2 = 'Output CSV Path';      $cfgSh.Range('A6').Font.Bold = $true
    $cfgSh.Range('B6').Interior.Color = $editBg
    $cfgSh.Range('C6').Value2 = 'Optional. Leave blank to use script default (timestamped file in InventoryAndReport_OutputCsvPath\).'
    $cfgSh.Range('C6').Font.Italic = $true; $cfgSh.Range('C6').Font.Color = $darkGrey

    foreach ($r in 4,5,6) { $cfgSh.Rows.Item($r).RowHeight = 18 }

    # PnP note row 8
    $cfgSh.Range('A8').Value2 = 'PnP Note'; $cfgSh.Range('A8').Font.Bold = $true
    $cfgSh.Range('B8:C8').Merge()
    $cfgSh.Range('B8').Value2 = 'ONDR and SPOL scripts use PnP.PowerShell and require a registered Entra app. Set the PNP_CLIENT_ID environment variable before running (or pass -ClientId to Connect-PnPOnline). See RUNBOOK-InventoryAndReport.md §1 for setup.'
    $cfgSh.Range('B8').WrapText = $true
    $cfgSh.Range('B8').Font.Italic = $true; $cfgSh.Range('B8').Font.Color = $warnFg
    $cfgSh.Range('B8').Interior.Color = $warnBg
    $cfgSh.Rows.Item(8).RowHeight = 48

    $cfgSh.Columns.Item(1).ColumnWidth = 24
    $cfgSh.Columns.Item(2).ColumnWidth = 52
    $cfgSh.Columns.Item(3).ColumnWidth = 72

    # Named ranges pointing to config value cells (use formula string, not range object)
    $wb.Names.Add('cfg_ScriptRoot',         '=Config!$B$4') | Out-Null
    $wb.Names.Add('cfg_SharePointAdminUrl', '=Config!$B$5') | Out-Null
    $wb.Names.Add('cfg_OutputCsvPath',      '=Config!$B$6') | Out-Null

    # ══════════════════════════════════════════════════════════════════════════
    # SCRIPTS SHEET
    # ══════════════════════════════════════════════════════════════════════════

    # Header row 1 — use explicit Range addresses
    $scrSh.Range('A1').Value2 = '#';               $scrSh.Range('A1').Font.Bold = $true; $scrSh.Range('A1').Interior.Color = $navyBg; $scrSh.Range('A1').Font.Color = $white
    $scrSh.Range('B1').Value2 = 'Include';          $scrSh.Range('B1').Font.Bold = $true; $scrSh.Range('B1').Interior.Color = $navyBg; $scrSh.Range('B1').Font.Color = $white
    $scrSh.Range('C1').Value2 = 'Workload';         $scrSh.Range('C1').Font.Bold = $true; $scrSh.Range('C1').Interior.Color = $navyBg; $scrSh.Range('C1').Font.Color = $white
    $scrSh.Range('D1').Value2 = 'Script ID';        $scrSh.Range('D1').Font.Bold = $true; $scrSh.Range('D1').Interior.Color = $navyBg; $scrSh.Range('D1').Font.Color = $white
    $scrSh.Range('E1').Value2 = 'Description';      $scrSh.Range('E1').Font.Bold = $true; $scrSh.Range('E1').Interior.Color = $navyBg; $scrSh.Range('E1').Font.Color = $white
    $scrSh.Range('F1').Value2 = 'Script File';      $scrSh.Range('F1').Font.Bold = $true; $scrSh.Range('F1').Interior.Color = $navyBg; $scrSh.Range('F1').Font.Color = $white
    $scrSh.Range('G1').Value2 = 'Scope Mode';       $scrSh.Range('G1').Font.Bold = $true; $scrSh.Range('G1').Interior.Color = $navyBg; $scrSh.Range('G1').Font.Color = $white
    $scrSh.Range('H1').Value2 = 'Input CSV';        $scrSh.Range('H1').Font.Bold = $true; $scrSh.Range('H1').Interior.Color = $navyBg; $scrSh.Range('H1').Font.Color = $white
    $scrSh.Range('I1').Value2 = 'SP Admin Req';     $scrSh.Range('I1').Font.Bold = $true; $scrSh.Range('I1').Interior.Color = $navyBg; $scrSh.Range('I1').Font.Color = $white
    $scrSh.Range('J1').Value2 = 'Generated Command';$scrSh.Range('J1').Font.Bold = $true; $scrSh.Range('J1').Interior.Color = $navyBg; $scrSh.Range('J1').Font.Color = $white
    $scrSh.Rows.Item(1).RowHeight = 20

    # Data rows 2–41
    $DR      = 2
    $lastRow = $DR + $catalog.Count - 1

    for ($i = 0; $i -lt $catalog.Count; $i++) {
        $s   = $catalog[$i]
        $row = $DR + $i
        $wc  = $wlColors[$s.Workload]

        # Use Range("XN") addressing — avoids Cells.Item COM dispatch ambiguity
        $scrSh.Range("A${row}").Value2 = [string]($i + 1)
        $scrSh.Range("B${row}").Value2 = 'Yes'
        $scrSh.Range("C${row}").Value2 = $s.Workload
        $scrSh.Range("D${row}").Value2 = $s.ID
        $scrSh.Range("E${row}").Value2 = $s.Desc
        $scrSh.Range("F${row}").Value2 = $s.File
        $scrSh.Range("G${row}").Value2 = 'InputCsvPath'
        $scrSh.Range("H${row}").Value2 = $s.InputCsv
        $scrSh.Range("I${row}").Value2 = $s.NeedsSP

        # Workload color baseline for full row, then override specific columns
        $scrSh.Range("A${row}:J${row}").Interior.Color = $wc
        $scrSh.Range("B${row}").Interior.Color = $editBg   # editable
        $scrSh.Range("G${row}").Interior.Color = $editBg   # editable
        $scrSh.Range("I${row}").Interior.Color = $lockBg   # computed/locked
        $scrSh.Range("I${row}").Font.Italic    = $true
        $scrSh.Range("J${row}").Interior.Color = $lockBg   # computed/locked
    }

    # Generated Command formulas — column J
    # Note: """ in PS single-quoted string = 3 literal " chars = Excel escaped-quote idiom
    for ($row = $DR; $row -le $lastRow; $row++) {
        $f = '=IF(B' + $row + '="Yes","pwsh "&cfg_ScriptRoot&"/SharedModule/Online/InventoryAndReport/"&F' + $row +
             '&IF(G' + $row + '="DiscoverAll"," -DiscoverAll"," -InputCsvPath """&cfg_ScriptRoot&"/SharedModule/Online/InventoryAndReport/"&H' + $row + '&"""")' +
             '&IF(AND(I' + $row + '="Yes",cfg_SharePointAdminUrl<>"")," -SharePointAdminUrl """&cfg_SharePointAdminUrl&"""","")' +
             '&IF(cfg_OutputCsvPath<>""," -OutputCsvPath """&cfg_OutputCsvPath&"""",""),"")'
        $scrSh.Range("J${row}").Formula   = $f
        $scrSh.Range("J${row}").Font.Name = 'Consolas'
        $scrSh.Range("J${row}").Font.Size = 9
        $scrSh.Range("J${row}").Font.Color = clr '1F3864'
    }

    # Data validation: Include (B) — Yes/No dropdown
    $bRng = $scrSh.Range("B${DR}:B${lastRow}")
    $bRng.Validation.Delete()
    $bRng.Validation.Add(3, 1, 1, 'Yes,No')

    # Data validation: Scope Mode (G) — InputCsvPath/DiscoverAll dropdown
    $gRng = $scrSh.Range("G${DR}:G${lastRow}")
    $gRng.Validation.Delete()
    $gRng.Validation.Add(3, 1, 1, 'InputCsvPath,DiscoverAll')

    # Conditional formatting: Include col B — Yes = green, No = grey
    $cfRng = $scrSh.Range("B${DR}:B${lastRow}")
    $cfRng.FormatConditions.Delete()
    $fc = $cfRng.FormatConditions.Add(1, 3, 'Yes')
    $fc.Interior.Color = $greenBg; $fc.Font.Color = $greenFg; $fc.Font.Bold = $true
    $fc = $cfRng.FormatConditions.Add(1, 3, 'No')
    $fc.Interior.Color = $lockBg;  $fc.Font.Color = $greyText

    # Column widths (integer indices: A=1 … J=10)
    $scrSh.Columns.Item(1).ColumnWidth  =  4   # #
    $scrSh.Columns.Item(2).ColumnWidth  = 10   # Include
    $scrSh.Columns.Item(3).ColumnWidth  =  9   # Workload
    $scrSh.Columns.Item(4).ColumnWidth  = 16   # Script ID
    $scrSh.Columns.Item(5).ColumnWidth  = 46   # Description
    $scrSh.Columns.Item(6).ColumnWidth  = 55   # Script File
    $scrSh.Columns.Item(7).ColumnWidth  = 16   # Scope Mode
    $scrSh.Columns.Item(8).ColumnWidth  = 56   # Input CSV
    $scrSh.Columns.Item(9).ColumnWidth  = 14   # SP Admin Req
    $scrSh.Columns.Item(10).ColumnWidth = 90   # Generated Command

    # Freeze header row
    $scrSh.Activate()
    $xl.ActiveWindow.SplitRow    = 1
    $xl.ActiveWindow.SplitColumn = 0
    $xl.ActiveWindow.FreezePanes = $true

    # ══════════════════════════════════════════════════════════════════════════
    # RUNCOMMANDS SHEET
    # ══════════════════════════════════════════════════════════════════════════

    $runSh.Range('A1:B1').Merge()
    $cell = $runSh.Range('A1')
    $cell.Value2 = 'Run Commands'
    $cell.Font.Bold = $true; $cell.Font.Size = 14
    $cell.Interior.Color = $navyBg; $cell.Font.Color = $white
    $runSh.Rows.Item(1).RowHeight = 28

    $runSh.Range('A2').Value2 = 'Commands below reflect all scripts with Include = Yes from the Scripts sheet, built from current Config values. Copy and paste into a PowerShell 7 terminal from the repository root directory.'
    $runSh.Range('A2').Font.Italic = $true
    $runSh.Range('A2').Font.Color  = $darkGrey
    $runSh.Rows.Item(2).RowHeight  = 24

    # FILTER formula — spills one command per included script
    $runSh.Range('A3').Formula   = '=IFERROR(FILTER(Scripts!J$2:J$41,Scripts!B$2:B$41="Yes"),"No scripts selected — set Include = Yes on the Scripts sheet.")'
    $runSh.Range('A3').Font.Name = 'Consolas'
    $runSh.Range('A3').Font.Size = 9
    $runSh.Columns.Item(1).ColumnWidth = 160

    # Freeze rows 1–2
    $runSh.Activate()
    $xl.ActiveWindow.SplitRow    = 2
    $xl.ActiveWindow.SplitColumn = 0
    $xl.ActiveWindow.FreezePanes = $true

    # Open on Config tab, cursor on first editable cell
    $cfgSh.Activate()
    $cfgSh.Cells.Item(4,2).Select()

    # ── Save ─────────────────────────────────────────────────────────────────
    if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
    $wb.SaveAs($OutputPath, 51)   # 51 = xlOpenXMLWorkbook
    $wb.Close($false)
    Write-Host "Saved: $OutputPath" -ForegroundColor Green

} catch {
    Write-Host "ERROR at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)" -ForegroundColor Red
    throw
} finally {
    $xl.ScreenUpdating = $true
    $xl.DisplayAlerts  = $true
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
