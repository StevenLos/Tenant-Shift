<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260315-133600

.POWERSHELLREQUIRED
5.1+

.REQUIREDMODULES
None declared in this file

.MODULEVERSIONPOLICY
Not declared in this file
#>
#Requires -Version 5.1

[CmdletBinding()]
param(
    [ValidateRange(5, 3600)]
    [int]$IntervalSeconds = 59,

    [ValidateRange(1, 10080)]
    [int]$DurationMinutes,

    [string]$KeySequence = '+{F15}'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Status {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'INFO' { 'Cyan' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        'SUCCESS' { 'Green' }
    }

    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

if ($env:OS -ne 'Windows_NT') {
    throw 'Utility-Espresso requires Windows because it uses WScript.Shell COM automation.'
}

Write-Status -Message 'Starting Utility-Espresso keep-awake loop.'
Write-Status -Message "Key sequence: $KeySequence"
Write-Status -Message "Interval seconds: $IntervalSeconds"

$stopAt = $null
if ($PSBoundParameters.ContainsKey('DurationMinutes')) {
    $stopAt = (Get-Date).AddMinutes($DurationMinutes)
    Write-Status -Message "Duration limit: $DurationMinutes minute(s) (stops around $(Get-Date -Date $stopAt -Format 'yyyy-MM-dd HH:mm:ss'))."
}
else {
    Write-Status -Message 'No duration set; running until interrupted (Ctrl+C).' -Level WARN
}

$wsh = New-Object -ComObject WScript.Shell
$sentCount = 0

while ($true) {
    if ($null -ne $stopAt -and (Get-Date) -ge $stopAt) {
        break
    }

    $wsh.SendKeys($KeySequence)
    $sentCount++
    $cycleTime = Get-Date
    Write-Status -Message ("Cycle {0}: sent keep-awake key at {1}" -f $sentCount, $cycleTime.ToString('yyyy-MM-dd HH:mm:ss'))

    Start-Sleep -Seconds $IntervalSeconds
}

Write-Status -Message "Utility-Espresso completed after sending $sentCount key sequence(s)." -Level SUCCESS
