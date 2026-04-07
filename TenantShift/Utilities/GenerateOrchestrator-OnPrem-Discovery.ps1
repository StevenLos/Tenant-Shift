<#
.SYNOPSIS
    Generates Discover-Orchestrator-OnPrem.xlsx for OnPrem scripts via Excel COM automation.
#>
[CmdletBinding()]
param(
    [string]$OutputPath = "$PSScriptRoot\Discover-Orchestrator-OnPrem.xlsx"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─── Script catalog ───────────────────────────────────────────────────────────
# HasServer: script accepts -Server (DC hostname)
# HasSBase:  script accepts -SearchBase (OU DN; DiscoverAll mode only)
$catalog = @(
    [pscustomobject]@{Workload='ADUC';ID='D-ADUC-0010';Desc='Get Active Directory Organizational Units';                   File='D-ADUC-0010-Get-ActiveDirectoryOrganizationalUnits.ps1';                    InputCsv='Scope-ActiveDirectoryOUs.input.csv';                      HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='ADUC';ID='D-ADUC-0020';Desc='Get Active Directory Users';                                  File='D-ADUC-0020-Get-ActiveDirectoryUsers.ps1';                                  InputCsv='Scope-ActiveDirectoryUsers.input.csv';                    HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='ADUC';ID='D-ADUC-0030';Desc='Get Active Directory Contacts';                               File='D-ADUC-0030-Get-ActiveDirectoryContacts.ps1';                               InputCsv='Scope-ActiveDirectoryContacts.input.csv';                 HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='ADUC';ID='D-ADUC-0040';Desc='Get Active Directory Security Groups';                        File='D-ADUC-0040-Get-ActiveDirectorySecurityGroups.ps1';                         InputCsv='Scope-ActiveDirectorySecurityGroups.input.csv';           HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='ADUC';ID='D-ADUC-0050';Desc='Get Active Directory Distribution Groups';                    File='D-ADUC-0050-Get-ActiveDirectoryDistributionGroups.ps1';                     InputCsv='Scope-ActiveDirectoryDistributionGroups.input.csv';       HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='ADUC';ID='D-ADUC-0060';Desc='Get Active Directory Security Group Members';                 File='D-ADUC-0060-Get-ActiveDirectorySecurityGroupMembers.ps1';                   InputCsv='Scope-ActiveDirectorySecurityGroups.input.csv';           HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='ADUC';ID='D-ADUC-0070';Desc='Get Active Directory Distribution Group Members';             File='D-ADUC-0070-Get-ActiveDirectoryDistributionGroupMembers.ps1';               InputCsv='Scope-ActiveDirectoryDistributionGroups.input.csv';       HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='ADUC';ID='D-ADUC-0100';Desc='Get Active Directory User Recursive Group Memberships';       File='D-ADUC-0100-Get-ActiveDirectoryUserRecursiveGroupMemberships.ps1';          InputCsv='Scope-ActiveDirectoryUsers.input.csv';                    HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='ADUC';ID='D-ADUC-0110';Desc='Get Active Directory Users Without Group Memberships';        File='D-ADUC-0110-Get-ActiveDirectoryUsersWithoutGroupMemberships.ps1';           InputCsv='Scope-ActiveDirectoryUsers.input.csv';                    HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='ADUC';ID='D-ADUC-0120';Desc='Get Active Directory Groups Without Members';                 File='D-ADUC-0120-Get-ActiveDirectoryGroupsWithoutMembers.ps1';                   InputCsv='Scope-ActiveDirectorySecurityGroups.input.csv';           HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0010';Desc='Get Exchange On-Prem Mail Contacts';                          File='D-EXOP-0010-Get-ExchangeOnPremMailContacts.ps1';                            InputCsv='Scope-ExchangeOnPremMailContacts.input.csv';              HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0020';Desc='Get Exchange On-Prem Distribution Lists';                     File='D-EXOP-0020-Get-ExchangeOnPremDistributionLists.ps1';                       InputCsv='Scope-ExchangeOnPremDistributionLists.input.csv';         HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0030';Desc='Get Exchange On-Prem Mail-Enabled Security Groups';           File='D-EXOP-0030-Get-ExchangeOnPremMailEnabledSecurityGroups.ps1';               InputCsv='Scope-ExchangeOnPremMailEnabledSecurityGroups.input.csv'; HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0040';Desc='Get Exchange On-Prem Dynamic Distribution Groups';            File='D-EXOP-0040-Get-ExchangeOnPremDynamicDistributionGroups.ps1';               InputCsv='Scope-ExchangeOnPremDynamicDistributionGroups.input.csv'; HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0050';Desc='Get Exchange On-Prem Shared Mailboxes';                       File='D-EXOP-0050-Get-ExchangeOnPremSharedMailboxes.ps1';                         InputCsv='Scope-ExchangeOnPremSharedMailboxes.input.csv';           HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0060';Desc='Get Exchange On-Prem Resource Mailboxes';                     File='D-EXOP-0060-Get-ExchangeOnPremResourceMailboxes.ps1';                       InputCsv='Scope-ExchangeOnPremResourceMailboxes.input.csv';         HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0070';Desc='Get Exchange On-Prem Distribution List Members';              File='D-EXOP-0070-Get-ExchangeOnPremDistributionListMembers.ps1';                 InputCsv='Scope-ExchangeOnPremDistributionLists.input.csv';         HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0080';Desc='Get Exchange On-Prem Shared Mailbox Permissions';             File='D-EXOP-0080-Get-ExchangeOnPremSharedMailboxPermissions.ps1';               InputCsv='Scope-ExchangeOnPremSharedMailboxes.input.csv';           HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0090';Desc='Get Exchange On-Prem Resource Mailbox Booking Delegates';     File='D-EXOP-0090-Get-ExchangeOnPremResourceMailboxBookingDelegates.ps1';         InputCsv='Scope-ExchangeOnPremResourceMailboxes.input.csv';         HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0100';Desc='Get Exchange On-Prem Mailbox Delegations';                    File='D-EXOP-0100-Get-ExchangeOnPremMailboxDelegations.ps1';                     InputCsv='Scope-ExchangeOnPremMailboxes.input.csv';                 HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0110';Desc='Get Exchange On-Prem Mailbox Folder Permissions';             File='D-EXOP-0110-Get-ExchangeOnPremMailboxFolderPermissions.ps1';               InputCsv='Scope-ExchangeOnPremMailboxes.input.csv';                 HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0150';Desc='Get Exchange On-Prem Inbound Connector Details';              File='D-EXOP-0150-Get-ExchangeOnPremInboundConnectorDetails.ps1';                 InputCsv='Scope-ExchangeOnPremInboundConnectors.input.csv';         HasServer='Yes';HasSBase='Yes'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0160';Desc='Get Exchange On-Prem Outlook Client Versions From RPC Logs';  File='D-EXOP-0160-Get-ExchangeOnPremOutlookClientVersionsFromRpcLogs.ps1';       InputCsv='Scope-ExchangeOnPremRpcLogs.input.csv';                   HasServer='No'; HasSBase='No'}
    [pscustomobject]@{Workload='EXOP';ID='D-EXOP-0170';Desc='Get Exchange On-Prem RPC Log Export';                         File='D-EXOP-0170-Get-ExchangeOnPremRpcLogExport.ps1';                           InputCsv='Scope-ExchangeOnPremRpcLogs.input.csv';                   HasServer='No'; HasSBase='No'}
    [pscustomobject]@{Workload='ADGP';ID='D-ADGP-0010';Desc='Get Group Policy Objects';                                    File='D-ADGP-0010-Get-GroupPolicyObjects.ps1';                                   InputCsv='D-ADGP-0010-Get-GroupPolicyObjects.input.csv';            HasServer='Yes';HasSBase='No'}
    [pscustomobject]@{Workload='ADGP';ID='D-ADGP-0020';Desc='Get Group Policy Links';                                      File='D-ADGP-0020-Get-GroupPolicyLinks.ps1';                                     InputCsv='D-ADGP-0020-Get-GroupPolicyLinks.input.csv';              HasServer='Yes';HasSBase='No'}
)

# ─── Color helper: RRGGBB hex string → Excel BGR decimal ──────────────────────
function clr([string]$hex) {
    $r = [Convert]::ToInt32($hex.Substring(0,2),16)
    $g = [Convert]::ToInt32($hex.Substring(2,2),16)
    $b = [Convert]::ToInt32($hex.Substring(4,2),16)
    return ($b -shl 16) + ($g -shl 8) + $r
}

$wlColors = @{
    ADUC = clr 'CCDCF5'   # light steel blue
    EXOP = clr 'CCEAD9'   # light sage green
    ADGP = clr 'E8D5F5'   # light lavender
}
$navyBg   = clr '1F3864'
$blueBg   = clr '2E75B6'
$white    = 16777215
$editBg   = clr 'FFFCD6'
$lockBg   = clr 'F2F2F2'
$darkGrey = clr '595959'
$warnBg   = clr 'FCE4D6'
$warnFg   = clr '843C0C'
$infoBg   = clr 'DEEAF1'
$infoFg   = clr '1F4E79'
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

    # Sheets: Config, Scripts, RunCommands, PostProcess
    while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }
    $wb.Sheets.Item(1).Name = 'Config'
    $cfgSh = $wb.Sheets.Item('Config')
    $scrSh = $wb.Sheets.Add([System.Reflection.Missing]::Value, $cfgSh)
    $scrSh.Name = 'Scripts'
    $runSh = $wb.Sheets.Add([System.Reflection.Missing]::Value, $scrSh)
    $runSh.Name = 'RunCommands'
    $ppSh  = $wb.Sheets.Add([System.Reflection.Missing]::Value, $runSh)
    $ppSh.Name  = 'PostProcess'

    # ══════════════════════════════════════════════════════════════════════════
    # CONFIG SHEET
    # ══════════════════════════════════════════════════════════════════════════

    $cfgSh.Range('A1:C1').Merge()
    $cell = $cfgSh.Range('A1')
    $cell.Value2 = 'Discover Orchestrator (OnPrem) — Configuration'
    $cell.Font.Bold = $true; $cell.Font.Size = 14
    $cell.Interior.Color = $navyBg; $cell.Font.Color = $white
    $cfgSh.Rows.Item(1).RowHeight = 28

    # Column headers row 3
    $cfgSh.Range('A3').Value2 = 'Setting'; $cfgSh.Range('A3').Font.Bold = $true; $cfgSh.Range('A3').Interior.Color = $blueBg; $cfgSh.Range('A3').Font.Color = $white
    $cfgSh.Range('B3').Value2 = 'Value';   $cfgSh.Range('B3').Font.Bold = $true; $cfgSh.Range('B3').Interior.Color = $blueBg; $cfgSh.Range('B3').Font.Color = $white
    $cfgSh.Range('C3').Value2 = 'Notes';   $cfgSh.Range('C3').Font.Bold = $true; $cfgSh.Range('C3').Interior.Color = $blueBg; $cfgSh.Range('C3').Font.Color = $white
    $cfgSh.Rows.Item(3).RowHeight = 18

    # Row 4 — Script Root Path
    $cfgSh.Range('A4').Value2 = 'Script Root Path';        $cfgSh.Range('A4').Font.Bold = $true
    $cfgSh.Range('B4').Interior.Color = $editBg
    $cfgSh.Range('C4').Value2 = 'Repo root path to run scripts from. Use . for current directory, or an absolute path. Use forward slashes.'
    $cfgSh.Range('C4').Font.Italic = $true; $cfgSh.Range('C4').Font.Color = $darkGrey

    # Row 5 — AD Server
    $cfgSh.Range('A5').Value2 = 'AD Server';               $cfgSh.Range('A5').Font.Bold = $true
    $cfgSh.Range('B5').Interior.Color = $editBg
    $cfgSh.Range('C5').Value2 = 'Optional. Domain controller hostname to target (passed as -Server). If blank, scripts use the default DC for the current domain. Applies to all ADUC, EXOP, and ADGP scripts that accept -Server.'
    $cfgSh.Range('C5').Font.Italic = $true; $cfgSh.Range('C5').Font.Color = $darkGrey

    # Row 6 — AD SearchBase
    $cfgSh.Range('A6').Value2 = 'AD SearchBase (DN)';      $cfgSh.Range('A6').Font.Bold = $true
    $cfgSh.Range('B6').Interior.Color = $editBg
    $cfgSh.Range('C6').Value2 = 'Optional. Distinguished name of the OU to scope DiscoverAll discovery (passed as -SearchBase). Example: OU=Corp,DC=contoso,DC=com. Injected only when Scope Mode = DiscoverAll and the script supports -SearchBase (see Has SBase column).'
    $cfgSh.Range('C6').Font.Italic = $true; $cfgSh.Range('C6').Font.Color = $darkGrey

    # Row 7 — Output CSV Path
    $cfgSh.Range('A7').Value2 = 'Output CSV Path';         $cfgSh.Range('A7').Font.Bold = $true
    $cfgSh.Range('B7').Interior.Color = $editBg
    $cfgSh.Range('C7').Value2 = 'Optional. Leave blank to use each script''s default (timestamped file in Discover_OutputCsvPath\).'
    $cfgSh.Range('C7').Font.Italic = $true; $cfgSh.Range('C7').Font.Color = $darkGrey

    # Row 8 — Include powershell Prefix
    $cfgSh.Range('A8').Value2 = 'Include powershell Prefix'; $cfgSh.Range('A8').Font.Bold = $true
    $cfgSh.Range('B8').Value2 = 'Yes'
    $cfgSh.Range('B8').Interior.Color = $editBg
    $cfgSh.Range('C8').Value2 = 'Set to Yes to prepend powershell to each generated command. All OnPrem scripts require Windows PowerShell 5.1 — do not use pwsh (PowerShell 7).'
    $cfgSh.Range('C8').Font.Italic = $true; $cfgSh.Range('C8').Font.Color = $darkGrey
    $cfgSh.Range('B8').Validation.Delete()
    $cfgSh.Range('B8').Validation.Add(3, 1, 1, 'Yes,No')

    foreach ($r in 4,5,6,7,8) { $cfgSh.Rows.Item($r).RowHeight = 18 }

    # Row 10 — EXOP note
    $cfgSh.Range('A10').Value2 = 'EXOP Note'; $cfgSh.Range('A10').Font.Bold = $true
    $cfgSh.Range('B10:C10').Merge()
    $cfgSh.Range('B10').Value2 = 'EXOP scripts (D-EXOP-*) must be run inside an active Exchange Management Shell (EMS) session connected to the on-prem Exchange organisation. Start the EMS and connect to the target server before running any D-EXOP command. D-EXOP-0160 and D-EXOP-0170 parse RPC client access log files and support additional parameters (-LogPath, -LookbackDays, -ClientSoftware) not included in the generated command — add them manually if needed.'
    $cfgSh.Range('B10').WrapText = $true
    $cfgSh.Range('B10').Font.Italic = $true; $cfgSh.Range('B10').Font.Color = $warnFg
    $cfgSh.Range('B10').Interior.Color = $warnBg
    $cfgSh.Rows.Item(10).RowHeight = 60

    # Row 11 — ADGP note
    $cfgSh.Range('A11').Value2 = 'ADGP Note'; $cfgSh.Range('A11').Font.Bold = $true
    $cfgSh.Range('B11:C11').Merge()
    $cfgSh.Range('B11').Value2 = 'ADGP scripts (D-ADGP-*) require the GroupPolicy RSAT module (RSAT Group Policy Management Tools) and must run under Windows PowerShell 5.1 — the GroupPolicy module does not load under pwsh (PowerShell 7). Install RSAT on the machine before running these scripts. Post-processing scripts D-ADGP-0030, D-ADGP-0040, and D-ADGP-0050 are not included here — see the PostProcess sheet for their example commands and prerequisites.'
    $cfgSh.Range('B11').WrapText = $true
    $cfgSh.Range('B11').Font.Italic = $true; $cfgSh.Range('B11').Font.Color = $infoFg
    $cfgSh.Range('B11').Interior.Color = $infoBg
    $cfgSh.Rows.Item(11).RowHeight = 60

    $cfgSh.Columns.Item(1).ColumnWidth = 26
    $cfgSh.Columns.Item(2).ColumnWidth = 52
    $cfgSh.Columns.Item(3).ColumnWidth = 72

    # Named ranges
    $wb.Names.Add('cfg_ScriptRoot',    '=Config!$B$4') | Out-Null
    $wb.Names.Add('cfg_Server',        '=Config!$B$5') | Out-Null
    $wb.Names.Add('cfg_SearchBase',    '=Config!$B$6') | Out-Null
    $wb.Names.Add('cfg_OutputCsvPath', '=Config!$B$7') | Out-Null
    $wb.Names.Add('cfg_PsPrefix',      '=Config!$B$8') | Out-Null

    # ══════════════════════════════════════════════════════════════════════════
    # SCRIPTS SHEET
    # ══════════════════════════════════════════════════════════════════════════

    # Header row 1 — columns A through K
    $scrSh.Range('A1').Value2 = '#';                  $scrSh.Range('A1').Font.Bold = $true; $scrSh.Range('A1').Interior.Color = $navyBg; $scrSh.Range('A1').Font.Color = $white
    $scrSh.Range('B1').Value2 = 'Include';             $scrSh.Range('B1').Font.Bold = $true; $scrSh.Range('B1').Interior.Color = $navyBg; $scrSh.Range('B1').Font.Color = $white
    $scrSh.Range('C1').Value2 = 'Workload';            $scrSh.Range('C1').Font.Bold = $true; $scrSh.Range('C1').Interior.Color = $navyBg; $scrSh.Range('C1').Font.Color = $white
    $scrSh.Range('D1').Value2 = 'Script ID';           $scrSh.Range('D1').Font.Bold = $true; $scrSh.Range('D1').Interior.Color = $navyBg; $scrSh.Range('D1').Font.Color = $white
    $scrSh.Range('E1').Value2 = 'Description';         $scrSh.Range('E1').Font.Bold = $true; $scrSh.Range('E1').Interior.Color = $navyBg; $scrSh.Range('E1').Font.Color = $white
    $scrSh.Range('F1').Value2 = 'Script File';         $scrSh.Range('F1').Font.Bold = $true; $scrSh.Range('F1').Interior.Color = $navyBg; $scrSh.Range('F1').Font.Color = $white
    $scrSh.Range('G1').Value2 = 'Scope Mode';          $scrSh.Range('G1').Font.Bold = $true; $scrSh.Range('G1').Interior.Color = $navyBg; $scrSh.Range('G1').Font.Color = $white
    $scrSh.Range('H1').Value2 = 'Input CSV';           $scrSh.Range('H1').Font.Bold = $true; $scrSh.Range('H1').Interior.Color = $navyBg; $scrSh.Range('H1').Font.Color = $white
    $scrSh.Range('I1').Value2 = 'Has Server';          $scrSh.Range('I1').Font.Bold = $true; $scrSh.Range('I1').Interior.Color = $navyBg; $scrSh.Range('I1').Font.Color = $white
    $scrSh.Range('J1').Value2 = 'Has SBase';           $scrSh.Range('J1').Font.Bold = $true; $scrSh.Range('J1').Interior.Color = $navyBg; $scrSh.Range('J1').Font.Color = $white
    $scrSh.Range('K1').Value2 = 'Generated Command';   $scrSh.Range('K1').Font.Bold = $true; $scrSh.Range('K1').Interior.Color = $navyBg; $scrSh.Range('K1').Font.Color = $white
    $scrSh.Rows.Item(1).RowHeight = 20

    # Data rows 2–27
    $DR      = 2
    $lastRow = $DR + $catalog.Count - 1

    for ($i = 0; $i -lt $catalog.Count; $i++) {
        $s   = $catalog[$i]
        $row = $DR + $i
        $wc  = $wlColors[$s.Workload]

        $scrSh.Range("A${row}").Value2 = [string]($i + 1)
        $scrSh.Range("B${row}").Value2 = 'Yes'
        $scrSh.Range("C${row}").Value2 = $s.Workload
        $scrSh.Range("D${row}").Value2 = $s.ID
        $scrSh.Range("E${row}").Value2 = $s.Desc
        $scrSh.Range("F${row}").Value2 = $s.File
        $scrSh.Range("G${row}").Value2 = 'InputCsvPath'
        $scrSh.Range("H${row}").Value2 = $s.InputCsv
        $scrSh.Range("I${row}").Value2 = $s.HasServer
        $scrSh.Range("J${row}").Value2 = $s.HasSBase

        # Workload color baseline for full row, then override specific columns
        $scrSh.Range("A${row}:K${row}").Interior.Color = $wc
        $scrSh.Range("B${row}").Interior.Color = $editBg   # editable
        $scrSh.Range("G${row}").Interior.Color = $editBg   # editable
        $scrSh.Range("I${row}").Interior.Color = $lockBg   # locked metadata
        $scrSh.Range("I${row}").Font.Italic    = $true
        $scrSh.Range("J${row}").Interior.Color = $lockBg   # locked metadata
        $scrSh.Range("J${row}").Font.Italic    = $true
        $scrSh.Range("K${row}").Interior.Color = $lockBg   # generated/locked
    }

    # Generated Command formulas — column K
    # Formula injects: powershell prefix, script path, scope mode, -Server (if HasServer=Yes and cfg_Server set),
    # -SearchBase (if HasSBase=Yes and DiscoverAll and cfg_SearchBase set), -OutputCsvPath (if cfg_OutputCsvPath set)
    for ($row = $DR; $row -le $lastRow; $row++) {
        $f = '=IF(B' + $row + '="Yes",' +
             'IF(cfg_PsPrefix="Yes","powershell ","")' +
             '&cfg_ScriptRoot&"/TenantShift/OnPrem/Discover/"&F' + $row +
             '&IF(G' + $row + '="DiscoverAll"," -DiscoverAll"," -InputCsvPath """&cfg_ScriptRoot&"/TenantShift/OnPrem/Discover/"&H' + $row + '&"""")' +
             '&IF(AND(I' + $row + '="Yes",cfg_Server<>"")," -Server """&cfg_Server&"""","")' +
             '&IF(AND(J' + $row + '="Yes",G' + $row + '="DiscoverAll",cfg_SearchBase<>"")," -SearchBase """&cfg_SearchBase&"""","")' +
             '&IF(cfg_OutputCsvPath<>""," -OutputCsvPath """&cfg_OutputCsvPath&"""","")' +
             ',"")'
        $scrSh.Range("K${row}").Formula   = $f
        $scrSh.Range("K${row}").Font.Name = 'Consolas'
        $scrSh.Range("K${row}").Font.Size = 9
        $scrSh.Range("K${row}").Font.Color = clr '1F3864'
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

    # Column widths (integer indices: A=1 … K=11)
    $scrSh.Columns.Item(1).ColumnWidth  =  4   # #
    $scrSh.Columns.Item(2).ColumnWidth  = 10   # Include
    $scrSh.Columns.Item(3).ColumnWidth  =  9   # Workload
    $scrSh.Columns.Item(4).ColumnWidth  = 16   # Script ID
    $scrSh.Columns.Item(5).ColumnWidth  = 52   # Description
    $scrSh.Columns.Item(6).ColumnWidth  = 64   # Script File
    $scrSh.Columns.Item(7).ColumnWidth  = 16   # Scope Mode
    $scrSh.Columns.Item(8).ColumnWidth  = 56   # Input CSV
    $scrSh.Columns.Item(9).ColumnWidth  = 12   # Has Server
    $scrSh.Columns.Item(10).ColumnWidth = 12   # Has SBase
    $scrSh.Columns.Item(11).ColumnWidth = 90   # Generated Command

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

    $runSh.Range('A2').Value2 = 'Commands below reflect all scripts with Include = Yes from the Scripts sheet, built from current Config values. Copy and paste into a Windows PowerShell 5.1 terminal from the repository root directory.'
    $runSh.Range('A2').Font.Italic = $true
    $runSh.Range('A2').Font.Color  = $darkGrey
    $runSh.Rows.Item(2).RowHeight  = 24

    # FILTER formula — spills one command per included script
    $runSh.Range('A3').Formula   = '=IFERROR(FILTER(Scripts!K$2:K$27,Scripts!B$2:B$27="Yes"),"No scripts selected — set Include = Yes on the Scripts sheet.")'
    $runSh.Range('A3').Font.Name = 'Consolas'
    $runSh.Range('A3').Font.Size = 9
    $runSh.Columns.Item(1).ColumnWidth = 160

    # Freeze rows 1–2
    $runSh.Activate()
    $xl.ActiveWindow.SplitRow    = 2
    $xl.ActiveWindow.SplitColumn = 0
    $xl.ActiveWindow.FreezePanes = $true

    # ══════════════════════════════════════════════════════════════════════════
    # POSTPROCESS SHEET
    # ══════════════════════════════════════════════════════════════════════════

    $ppSh.Range('A1:B1').Merge()
    $cell = $ppSh.Range('A1')
    $cell.Value2 = 'ADGP Post-Processing Scripts'
    $cell.Font.Bold = $true; $cell.Font.Size = 14
    $cell.Interior.Color = $navyBg; $cell.Font.Color = $white
    $ppSh.Rows.Item(1).RowHeight = 28

    $ppSh.Range('A2:B2').Merge()
    $ppSh.Range('A2').Value2 = 'The scripts below are pure CSV-to-CSV transformers. They take the output files of earlier ADGP scripts as input and do not connect to Active Directory. Run them in sequence after D-ADGP-0010 and D-ADGP-0020 have completed. Commands shown assume you are running from the repository root directory with the default output path. Adjust CSV path glob patterns to match your actual output filenames if needed.'
    $ppSh.Range('A2').WrapText = $true
    $ppSh.Range('A2').Font.Italic = $true
    $ppSh.Range('A2').Font.Color  = $darkGrey
    $ppSh.Rows.Item(2).RowHeight  = 48

    # Helper: write a section block
    # Section rows layout: header | Description | Prerequisite(s) | Example Command
    function Write-PpSection {
        param($Sheet, [int]$HeaderRow, [string]$Title, [string]$Desc, [string]$Prereq, [string]$ExCmd)

        $Sheet.Range("A${HeaderRow}:B${HeaderRow}").Merge()
        $hdr = $Sheet.Range("A${HeaderRow}")
        $hdr.Value2 = $Title
        $hdr.Font.Bold = $true; $hdr.Font.Size = 11
        $hdr.Interior.Color = $blueBg; $hdr.Font.Color = $white
        $Sheet.Rows.Item($HeaderRow).RowHeight = 22

        $dRow = $HeaderRow + 1
        $Sheet.Range("A${dRow}").Value2 = 'Description'; $Sheet.Range("A${dRow}").Font.Bold = $true
        $Sheet.Range("B${dRow}").Value2 = $Desc
        $Sheet.Range("B${dRow}").WrapText = $true
        $Sheet.Rows.Item($dRow).RowHeight = 36

        $pRow = $HeaderRow + 2
        $Sheet.Range("A${pRow}").Value2 = 'Prerequisites'; $Sheet.Range("A${pRow}").Font.Bold = $true
        $Sheet.Range("B${pRow}").Value2 = $Prereq
        $Sheet.Range("B${pRow}").WrapText = $true
        $Sheet.Rows.Item($pRow).RowHeight = 30

        $cRow = $HeaderRow + 3
        $Sheet.Range("A${cRow}").Value2 = 'Example Command'; $Sheet.Range("A${cRow}").Font.Bold = $true
        $Sheet.Range("B${cRow}").Value2 = $ExCmd
        $Sheet.Range("B${cRow}").Font.Name = 'Consolas'
        $Sheet.Range("B${cRow}").Font.Size = 9
        $Sheet.Range("B${cRow}").Font.Color = clr '1F3864'
        $Sheet.Range("B${cRow}").Interior.Color = $lockBg
        $Sheet.Range("B${cRow}").WrapText = $true
        $Sheet.Rows.Item($cRow).RowHeight = 42
    }

    Write-PpSection -Sheet $ppSh -HeaderRow 4 `
        -Title 'D-ADGP-0030 — Get Group Policy Scope Tree' `
        -Desc  'Reads the output CSV from D-ADGP-0020-Get-GroupPolicyLinks.ps1 and produces a hierarchical GPO scope tree with tree-key sort order and GPO precedence per scope. No AD calls are made — pure CSV post-processor.' `
        -Prereq 'Run D-ADGP-0020 first. Provide its output CSV via -LinksCsvPath.' `
        -ExCmd  'powershell .\TenantShift\OnPrem\Discover\D-ADGP-0030-Get-GroupPolicyScopeTree.ps1 -LinksCsvPath .\TenantShift\OnPrem\Discover\Discover_OutputCsvPath\Results_D-ADGP-0020*.csv'

    Write-PpSection -Sheet $ppSh -HeaderRow 9 `
        -Title 'D-ADGP-0040 — Get Group Policy Summary' `
        -Desc  'Reads the output CSV from D-ADGP-0020-Get-GroupPolicyLinks.ps1 and produces per-scope and per-GPO summary statistics including enforced link counts and a risk tier (High/Medium/Low). Optionally enriches output with empty-GPO flags from D-ADGP-0010 when -GpoObjectsCsvPath is supplied. No AD calls are made — pure CSV post-processor.' `
        -Prereq 'Run D-ADGP-0020 first (mandatory). Run D-ADGP-0010 first for empty-GPO enrichment (optional).' `
        -ExCmd  'powershell .\TenantShift\OnPrem\Discover\D-ADGP-0040-Get-GroupPolicySummary.ps1 -LinksCsvPath .\TenantShift\OnPrem\Discover\Discover_OutputCsvPath\Results_D-ADGP-0020*.csv -GpoObjectsCsvPath .\TenantShift\OnPrem\Discover\Discover_OutputCsvPath\Results_D-ADGP-0010*.csv'

    Write-PpSection -Sheet $ppSh -HeaderRow 14 `
        -Title 'D-ADGP-0050 — Get Group Policy Unified View' `
        -Desc  'Merges the output CSVs from D-ADGP-0010, D-ADGP-0020, D-ADGP-0030, and D-ADGP-0040 into a single flat export with one row per GPO-scope link. Designed for Excel pivot table analysis and Power BI ingestion. D-ADGP-0030 and D-ADGP-0040 inputs are optional. No AD calls are made — pure CSV post-processor.' `
        -Prereq 'Run D-ADGP-0010 and D-ADGP-0020 first (mandatory). Run D-ADGP-0030 and D-ADGP-0040 first for TreeKey, precedence, risk tier, and empty-GPO enrichment (optional but recommended).' `
        -ExCmd  'powershell .\TenantShift\OnPrem\Discover\D-ADGP-0050-Get-GroupPolicyUnifiedView.ps1 -GpoObjectsCsvPath .\TenantShift\OnPrem\Discover\Discover_OutputCsvPath\Results_D-ADGP-0010*.csv -LinksCsvPath .\TenantShift\OnPrem\Discover\Discover_OutputCsvPath\Results_D-ADGP-0020*.csv -ScopeTreeCsvPath .\TenantShift\OnPrem\Discover\Discover_OutputCsvPath\Results_D-ADGP-0030*.csv -SummaryCsvPath .\TenantShift\OnPrem\Discover\Discover_OutputCsvPath\Results_D-ADGP-0040*.csv'

    $ppSh.Columns.Item(1).ColumnWidth =  20
    $ppSh.Columns.Item(2).ColumnWidth = 140

    # ── Save ─────────────────────────────────────────────────────────────────
    $cfgSh.Activate()
    $cfgSh.Cells.Item(4,2).Select()

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
