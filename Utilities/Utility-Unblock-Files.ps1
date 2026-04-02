<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260401-000100

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
None

.MODULEVERSIONPOLICY
Not declared in this file

.SYNOPSIS
    Recursively removes internet-origin blocking metadata from files under a path.

.DESCRIPTION
    Walks a target file or directory and removes the Zone.Identifier alternate data
    stream from any files that have it. This is useful when utility or workload
    scripts were downloaded from the internet and need `Unblock-File` before they
    can run under a `RemoteSigned` execution policy.

.PARAMETER TargetPath
    File or directory to process. Directory input is scanned recursively.

.PARAMETER PassThru
    Returns structured result objects instead of exiting the host process.

.EXAMPLE
    .\Utilities\Utility-Unblock-Files\Utility-Unblock-Files.ps1 -TargetPath .\SharedModule

    Recursively unblock files under SharedModule.

.EXAMPLE
    .\Utilities\Utility-Unblock-Files\Utility-Unblock-Files.ps1 -TargetPath .\Downloads\SharedScripts -PassThru

    Recursively unblock files and return structured results.
#>
#Requires -Version 5.1

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$TargetPath,

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Status {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $color = switch ($Level) {
        'INFO' { 'Cyan' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        'SUCCESS' { 'Green' }
    }

    Write-Host "[$Level] $Message" -ForegroundColor $color
}

function Test-FileIsBlocked {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    try {
        $null = Get-Item -LiteralPath $FilePath -Stream Zone.Identifier -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Get-TargetFile {
    [CmdletBinding()]
    [OutputType([object[]])]
    param(
        [Parameter(Mandatory)]
        [string]$ResolvedPath
    )

    $item = Get-Item -LiteralPath $ResolvedPath -Force -ErrorAction Stop
    if ($item.PSIsContainer) {
        return @(Get-ChildItem -LiteralPath $ResolvedPath -File -Recurse -Force -ErrorAction Stop)
    }

    return @($item)
}

$resolvedTargetPath = (Resolve-Path -LiteralPath $TargetPath -ErrorAction Stop).Path
$targetFiles = @(Get-TargetFile -ResolvedPath $resolvedTargetPath)
$results = [System.Collections.Generic.List[object]]::new()

foreach ($file in $targetFiles) {
    $wasBlocked = Test-FileIsBlocked -FilePath $file.FullName
    $action = 'AlreadyUnblocked'

    if ($wasBlocked) {
        if ($PSCmdlet.ShouldProcess($file.FullName, 'Unblock file')) {
            Unblock-File -LiteralPath $file.FullName -ErrorAction Stop
            $action = 'Unblocked'
        }
        else {
            $action = 'WouldUnblock'
        }
    }

    $results.Add([pscustomobject]@{
            Path         = $file.FullName
            RelativePath = [System.IO.Path]::GetRelativePath($resolvedTargetPath, $file.FullName)
            WasBlocked   = $wasBlocked
            Action       = $action
        }) | Out-Null
}

if ($PassThru) {
    return @($results)
}

$blockedCount = @($results | Where-Object { $_.WasBlocked }).Count
$unblockedCount = @($results | Where-Object { $_.Action -eq 'Unblocked' }).Count
$alreadyClearCount = @($results | Where-Object { $_.Action -eq 'AlreadyUnblocked' }).Count

Write-Status -Message ("Target path: {0}" -f $resolvedTargetPath)
Write-Status -Message ("Files scanned: {0}" -f $results.Count)
Write-Status -Message ("Files previously blocked: {0}" -f $blockedCount)
Write-Status -Message ("Files unblocked: {0}" -f $unblockedCount) -Level 'SUCCESS'
Write-Status -Message ("Files already clear: {0}" -f $alreadyClearCount)

if ($results.Count -eq 0) {
    Write-Status -Message 'No files were found under the target path.' -Level 'WARN'
}

exit 0
