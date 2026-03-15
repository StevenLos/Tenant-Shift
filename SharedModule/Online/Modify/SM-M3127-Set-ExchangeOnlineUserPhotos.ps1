<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260304-153500

.POWERSHELLREQUIRED
7.0+

.REQUIREDMODULES
ExchangeOnlineManagement

.MODULEVERSIONPOLICY
Latest from PSGallery (validated at runtime by Assert-ModuleCurrent)
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$InputCsvPath,

    [string]$OutputCsvPath = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Modify_OutputCsvPath') -ChildPath ("Results_SM-M3127-Set-ExchangeOnlineUserPhotos_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$commonModulePath = Join-Path -Path $repoRoot -ChildPath 'Common\Online\M365.Common.psm1'
Import-Module $commonModulePath -Force -DisableNameChecking

$transcriptPath = Start-RunTranscript -OutputCsvPath $OutputCsvPath -ScriptPath $PSCommandPath

try {

function Get-TrimmedValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    return ([string]$Value).Trim()
}

$requiredHeaders = @(
    'MailboxIdentity',
    'PhotoPath',
    'RemovePhoto',
    'PreviewOnly',
    'Notes'
)

Write-Status -Message 'Starting Exchange Online user photo update script.'
Assert-ModuleCurrent -ModuleNames @('ExchangeOnlineManagement')
Ensure-ExchangeConnection

$setUserPhotoCommand = Get-Command -Name Set-UserPhoto -ErrorAction Stop
$supportsPreview = $setUserPhotoCommand.Parameters.ContainsKey('Preview')
$supportsSave = $setUserPhotoCommand.Parameters.ContainsKey('Save')

$removeUserPhotoCommand = Get-Command -Name Remove-UserPhoto -ErrorAction SilentlyContinue
$supportsRemoveUserPhoto = $null -ne $removeUserPhotoCommand

$rows = Import-ValidatedCsv -InputCsvPath $InputCsvPath -RequiredHeaders $requiredHeaders
$results = [System.Collections.Generic.List[object]]::new()

$rowNumber = 1
foreach ($row in $rows) {
    $mailboxIdentity = Get-TrimmedValue -Value $row.MailboxIdentity

    try {
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
            throw 'MailboxIdentity is required.'
        }

        $removePhoto = ConvertTo-Bool -Value (Get-TrimmedValue -Value $row.RemovePhoto)
        $previewOnly = ConvertTo-Bool -Value (Get-TrimmedValue -Value $row.PreviewOnly)
        $photoPath = Get-TrimmedValue -Value $row.PhotoPath

        if ($removePhoto -and -not [string]::IsNullOrWhiteSpace($photoPath)) {
            throw 'PhotoPath must be empty when RemovePhoto is TRUE.'
        }

        if (-not $removePhoto -and [string]::IsNullOrWhiteSpace($photoPath)) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetUserPhoto' -Status 'Skipped' -Message 'No photo action requested. Provide PhotoPath or set RemovePhoto to TRUE.'))
            $rowNumber++
            continue
        }

        $mailbox = Invoke-WithRetry -OperationName "Lookup mailbox $mailboxIdentity" -ScriptBlock {
            Get-ExchangeOnlineMailbox -Identity $mailboxIdentity -ErrorAction SilentlyContinue
        }

        if (-not $mailbox) {
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetUserPhoto' -Status 'NotFound' -Message 'Mailbox not found.'))
            $rowNumber++
            continue
        }

        if ($removePhoto) {
            if (-not $supportsRemoveUserPhoto) {
                throw 'Remove-UserPhoto is not available in this Exchange Online session.'
            }

            if ($PSCmdlet.ShouldProcess($mailboxIdentity, 'Remove Exchange Online mailbox photo')) {
                Invoke-WithRetry -OperationName "Remove photo for $mailboxIdentity" -ScriptBlock {
                    Remove-UserPhoto -Identity $mailbox.Identity -Confirm:$false -ErrorAction Stop
                }

                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'RemoveUserPhoto' -Status 'Updated' -Message 'Mailbox photo removed successfully.'))
            }
            else {
                $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'RemoveUserPhoto' -Status 'WhatIf' -Message 'Photo removal skipped due to WhatIf.'))
            }

            $rowNumber++
            continue
        }

        $resolvedPhotoPath = $photoPath
        if (-not [System.IO.Path]::IsPathRooted($resolvedPhotoPath)) {
            $resolvedPhotoPath = Join-Path -Path $repoRoot -ChildPath $resolvedPhotoPath
        }

        if (-not (Test-Path -LiteralPath $resolvedPhotoPath -PathType Leaf)) {
            throw "PhotoPath '$photoPath' does not exist (resolved path: '$resolvedPhotoPath')."
        }

        $fileBytes = [System.IO.File]::ReadAllBytes($resolvedPhotoPath)
        if ($fileBytes.Length -eq 0) {
            throw "PhotoPath '$photoPath' resolved to an empty file."
        }

        if ($PSCmdlet.ShouldProcess($mailboxIdentity, 'Set Exchange Online mailbox photo')) {
            $warnings = [System.Collections.Generic.List[string]]::new()

            if ($supportsPreview -and $supportsSave) {
                Invoke-WithRetry -OperationName "Stage photo update for $mailboxIdentity" -ScriptBlock {
                    Set-UserPhoto -Identity $mailbox.Identity -PictureData $fileBytes -Preview -Confirm:$false -ErrorAction Stop
                }

                if (-not $previewOnly) {
                    Invoke-WithRetry -OperationName "Commit photo update for $mailboxIdentity" -ScriptBlock {
                        Set-UserPhoto -Identity $mailbox.Identity -Save -Confirm:$false -ErrorAction Stop
                    }
                }
            }
            else {
                if ($previewOnly) {
                    $warnings.Add('PreviewOnly was requested but this session does not support preview/save workflow. A direct update was performed.')
                }

                Invoke-WithRetry -OperationName "Set photo for $mailboxIdentity" -ScriptBlock {
                    Set-UserPhoto -Identity $mailbox.Identity -PictureData $fileBytes -Confirm:$false -ErrorAction Stop
                }
            }

            $message = if ($previewOnly) { 'Mailbox photo staged in preview mode.' } else { 'Mailbox photo updated successfully.' }
            if ($warnings.Count -gt 0) {
                $message = "$message Warnings: $($warnings -join ' ')"
            }

            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetUserPhoto' -Status 'Updated' -Message $message))
        }
        else {
            $whatIfMessage = if ($previewOnly) { 'Photo preview stage skipped due to WhatIf.' } else { 'Photo update skipped due to WhatIf.' }
            $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetUserPhoto' -Status 'WhatIf' -Message $whatIfMessage))
        }
    }
    catch {
        Write-Status -Message "Row $rowNumber ($mailboxIdentity) failed: $($_.Exception.Message)" -Level ERROR
        $results.Add((New-ResultObject -RowNumber $rowNumber -PrimaryKey $mailboxIdentity -Action 'SetUserPhoto' -Status 'Failed' -Message $_.Exception.Message))
    }

    $rowNumber++
}

Export-ResultsCsv -Results $results.ToArray() -OutputCsvPath $OutputCsvPath
Write-Status -Message 'Exchange Online user photo update script completed.' -Level SUCCESS
}
finally {
    Stop-RunTranscript
}
