#
# M365.Common.psd1 — Version manifest for M365.Common.psm1
#
# NOTE: Shared.Common.ps1 is dot-sourced at runtime via a guard at the top of
# M365.Common.psm1. The 'SharedCommonMinVersion' entry in PrivateData documents
# the minimum required version. Because dot-sourcing bypasses the PS module
# loader, this constraint cannot be enforced automatically — it is informational.
#
@{
    ModuleVersion     = '2.0.0'
    GUID              = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'
    Author            = 'Steven Los'
    CompanyName       = 'CarveOutToNewCo'
    Copyright         = 'Copyright (c) 2014-2026 Steven Los. MIT License.'
    Description       = 'Online (M365/Graph/EXO/PnP) common functions for the CarveOutToNewCo automation platform. Shared utility functions are dot-sourced from Shared.Common.ps1.'
    PowerShellVersion = '7.0'

    # PrivateData
    PrivateData = @{
        SharedCommonMinVersion = '1.0.0'  # Shared.Common.ps1 — dot-sourced at runtime; not enforced by PS loader
        PSData = @{
            Tags         = @('M365', 'Online', 'Graph', 'ExchangeOnline', 'SharePoint', 'CarveOutToNewCo')
            ProjectUri   = ''
            ReleaseNotes = '2.0.0 — Shared utility functions extracted to Shared.Common.ps1 (section-03). Added dot-source guard. Removed 8 duplicate function definitions.'
        }
    }
}
