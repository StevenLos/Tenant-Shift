#
# Shared.Common.psd1 — Version manifest for Shared.Common.ps1
#
# NOTE: This manifest is a versioning and documentation artifact only.
# Shared.Common.ps1 is loaded via dot-sourcing (. "$PSScriptRoot\Shared.Common.ps1")
# from within M365.Common.psm1 and OnPrem.Common.psm1.
# It is NOT loaded via Import-Module. Do not use New-ModuleManifest patterns
# that assume Import-Module loading (e.g., RootModule, NestedModules).
#
@{
    ModuleVersion     = '1.0.0'
    GUID              = '87e3688b-8232-4750-b995-e7933d6a4416'
    Author            = 'Steven Los'
    CompanyName       = 'CarveOutToNewCo'
    Copyright         = 'Copyright (c) 2014-2026 Steven Los. MIT License.'
    Description       = 'Shared utility functions for the CarveOutToNewCo automation platform. Dot-sourced by M365.Common.psm1 and OnPrem.Common.psm1. Contains 11 platform-wide utility functions extracted from both environment modules.'
    PowerShellVersion = '5.1'

    # FunctionsToExport is informational — dot-sourced files do not use Export-ModuleMember.
    # This list documents the public surface of Shared.Common.ps1.
    FunctionsToExport = @(
        'Write-Status'
        'Start-RunTranscript'
        'Stop-RunTranscript'
        'ConvertTo-Bool'
        'ConvertTo-Array'
        'Import-ValidatedCsv'
        'New-ResultObject'
        'Export-ResultsCsv'
        'Get-TrimmedValue'
        'Convert-MultiValueToString'
        'Convert-ToOrderedReportObject'
    )

    # PrivateData / PSData
    PrivateData = @{
        PSData = @{
            Tags        = @('SharedModule', 'CarveOutToNewCo', 'Utilities')
            ProjectUri  = ''
            ReleaseNotes = '1.0.0 — Initial extraction from M365.Common.psm1 and OnPrem.Common.psm1 (section-02).'
        }
    }
}
