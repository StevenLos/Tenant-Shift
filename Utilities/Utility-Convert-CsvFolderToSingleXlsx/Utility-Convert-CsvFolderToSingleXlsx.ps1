<#
.LICENSE
MIT License
Copyright (c) 2014-2026 Steven Los
See LICENSE file in repository root for full terms.

.SCRIPTVERSION
20260321-200000

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
    [Parameter(Mandatory)]
    [string]$InputFolderPath,

    [string]$OutputXlsxPath,

    [ValidateNotNullOrEmpty()]
    [string]$Delimiter = ',',

    [switch]$Recurse,

    [switch]$Overwrite,

    [switch]$NoTranscript
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

function Start-RunTranscript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OutputFilePath,

        [AllowNull()]
        [string]$ScriptPath
    )

    $directory = Split-Path -Path $OutputFilePath -Parent
    if ([string]::IsNullOrWhiteSpace($directory) -and -not [string]::IsNullOrWhiteSpace($ScriptPath)) {
        $directory = Split-Path -Path $ScriptPath -Parent
    }

    if ([string]::IsNullOrWhiteSpace($directory)) {
        throw "Unable to determine transcript directory from OutputFilePath '$OutputFilePath'."
    }

    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $scriptName = 'Script'
    if (-not [string]::IsNullOrWhiteSpace($ScriptPath)) {
        $candidate = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            $scriptName = $candidate
        }
    }

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $transcriptPath = Join-Path -Path $directory -ChildPath ("Transcript_{0}_{1}.log" -f $scriptName, $timestamp)

    Start-Transcript -LiteralPath $transcriptPath -Force -ErrorAction Stop | Out-Null
    Write-Status -Message "Transcript started at '$transcriptPath'."
    return $transcriptPath
}

function Stop-RunTranscript {
    [CmdletBinding()]
    param()

    try {
        Stop-Transcript -ErrorAction Stop | Out-Null
    }
    catch {
        $message = ([string]$_.Exception.Message).ToLowerInvariant()
        if ($message -notmatch 'not currently transcribing') {
            throw
        }
    }
}

function Resolve-FolderPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "Input folder not found: $Path"
    }

    return (Resolve-Path -LiteralPath $Path).Path
}

function Get-RelativePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$BasePath,

        [Parameter(Mandatory)]
        [string]$ChildPath
    )

    $baseResolved = (Resolve-Path -LiteralPath $BasePath).Path.TrimEnd('\', '/')
    $childResolved = (Resolve-Path -LiteralPath $ChildPath).Path

    $baseUri = [uri]($baseResolved + [System.IO.Path]::DirectorySeparatorChar)
    $childUri = [uri]$childResolved

    $relative = $baseUri.MakeRelativeUri($childUri).ToString()
    return [uri]::UnescapeDataString($relative) -replace '/', [System.IO.Path]::DirectorySeparatorChar
}

function Get-WorksheetBaseName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CsvFilePath,

        [Parameter(Mandatory)]
        [string]$RootFolderPath
    )

    $relativePath = Get-RelativePath -BasePath $RootFolderPath -ChildPath $CsvFilePath
    $withoutExtension = [System.IO.Path]::ChangeExtension($relativePath, $null)
    if ($withoutExtension.EndsWith('.')) {
        $withoutExtension = $withoutExtension.TrimEnd('.')
    }

    $normalized = ($withoutExtension -replace '[\\/]+', '__').Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        $normalized = [System.IO.Path]::GetFileNameWithoutExtension($CsvFilePath)
    }

    return $normalized
}

function Get-UniqueWorksheetName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PreferredName,

        [Parameter(Mandatory)]
        [System.Collections.Generic.HashSet[string]]$UsedNames
    )

    $candidateBase = ([string]$PreferredName).Trim()
    if ([string]::IsNullOrWhiteSpace($candidateBase)) {
        $candidateBase = 'Sheet'
    }

    $candidateBase = $candidateBase -replace '[:\\/?*\[\]]', '_'
    $candidateBase = $candidateBase -replace '\s+', ' '
    $candidateBase = $candidateBase.Trim(" '")

    if ([string]::IsNullOrWhiteSpace($candidateBase)) {
        $candidateBase = 'Sheet'
    }

    if ($candidateBase.Length -gt 31) {
        $candidateBase = $candidateBase.Substring(0, 31)
    }

    $candidate = $candidateBase
    $counter = 2

    while ($UsedNames.Contains($candidate)) {
        $suffix = "_$counter"
        $maxBaseLength = 31 - $suffix.Length
        $prefix = $candidateBase
        if ($prefix.Length -gt $maxBaseLength) {
            $prefix = $prefix.Substring(0, $maxBaseLength)
        }

        $candidate = "$prefix$suffix"
        $counter++
    }

    $null = $UsedNames.Add($candidate)
    return $candidate
}

function Read-CsvFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CsvPath,

        [Parameter(Mandatory)]
        [string]$Delimiter
    )

    try {
        Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
    }
    catch {
        # Assembly may already be loaded or available implicitly.
    }

    $fileInfo = Get-Item -LiteralPath $CsvPath -ErrorAction Stop
    if ($fileInfo.Length -eq 0) {
        return [PSCustomObject]@{
            Headers = @()
            Rows    = @()
        }
    }

    $rows = [System.Collections.Generic.List[object]]::new()
    $headers = @()
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($CsvPath)
    $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $parser.SetDelimiters($Delimiter)
    $parser.HasFieldsEnclosedInQuotes = $true
    $parser.TrimWhiteSpace = $false

    try {
        if (-not $parser.EndOfData) {
            $headers = @($parser.ReadFields())
            if ($headers.Count -gt 0) {
                $headers[0] = ([string]$headers[0]).TrimStart([char]0xFEFF)
            }
        }

        while (-not $parser.EndOfData) {
            $fields = @($parser.ReadFields())
            $rows.Add($fields) | Out-Null
        }
    }
    finally {
        $parser.Close()
    }

    return [PSCustomObject]@{
        Headers = @($headers)
        Rows    = @($rows)
    }
}

function Get-WorksheetMatrix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Headers,

        [Parameter(Mandatory)]
        [object[]]$Rows
    )

    $columnCount = $Headers.Count
    foreach ($row in $Rows) {
        $fieldCount = @($row).Count
        if ($fieldCount -gt $columnCount) {
            $columnCount = $fieldCount
        }
    }

    $headerRowCount = if ($Headers.Count -gt 0) { 1 } else { 0 }
    $rowCount = $Rows.Count + $headerRowCount

    if ($rowCount -eq 0 -or $columnCount -eq 0) {
        return [PSCustomObject]@{
            Matrix      = $null
            RowCount    = 0
            ColumnCount = 0
        }
    }

    $matrix = New-Object 'object[,]' $rowCount, $columnCount

    if ($Headers.Count -gt 0) {
        for ($columnIndex = 0; $columnIndex -lt $Headers.Count; $columnIndex++) {
            $matrix[0, $columnIndex] = [string]$Headers[$columnIndex]
        }
    }

    $dataStartRow = $headerRowCount
    for ($rowIndex = 0; $rowIndex -lt $Rows.Count; $rowIndex++) {
        $fields = @($Rows[$rowIndex])
        for ($columnIndex = 0; $columnIndex -lt $fields.Count; $columnIndex++) {
            $matrix[$rowIndex + $dataStartRow, $columnIndex] = [string]$fields[$columnIndex]
        }
    }

    return [PSCustomObject]@{
        Matrix      = $matrix
        RowCount    = $rowCount
        ColumnCount = $columnCount
    }
}

function Release-ComObjectSafely {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$ComObject
    )

    if ($null -eq $ComObject) {
        return
    }

    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
    }
    catch {
        # Best effort only.
    }
}

function Write-WorksheetData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Worksheet,

        [Parameter(Mandatory)]
        [string[]]$Headers,

        [Parameter(Mandatory)]
        [object[]]$Rows
    )

    $matrixInfo = Get-WorksheetMatrix -Headers $Headers -Rows $Rows
    if ($matrixInfo.RowCount -eq 0 -or $matrixInfo.ColumnCount -eq 0) {
        return
    }

    $topLeft = $null
    $bottomRight = $null
    $targetRange = $null
    $headerRange = $null

    try {
        $topLeft = $Worksheet.Cells.Item(1, 1)
        $bottomRight = $Worksheet.Cells.Item($matrixInfo.RowCount, $matrixInfo.ColumnCount)
        $targetRange = $Worksheet.Range($topLeft, $bottomRight)
        $targetRange.NumberFormat = '@'
        $targetRange.Value2 = $matrixInfo.Matrix
        $targetRange.EntireColumn.AutoFit() | Out-Null

        if ($Headers.Count -gt 0) {
            $headerRange = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item(1, $matrixInfo.ColumnCount))
            $headerRange.Font.Bold = $true
        }
    }
    finally {
        Release-ComObjectSafely -ComObject $headerRange
        Release-ComObjectSafely -ComObject $targetRange
        Release-ComObjectSafely -ComObject $bottomRight
        Release-ComObjectSafely -ComObject $topLeft
    }
}

function Get-CsvFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath,

        [Parameter(Mandatory)]
        [bool]$Recurse
    )

    return @(
        Get-ChildItem -LiteralPath $FolderPath -Filter '*.csv' -File -Recurse:$Recurse |
            Sort-Object -Property FullName
    )
}

if ($env:OS -ne 'Windows_NT') {
    throw 'Utility-Convert-CsvFolderToSingleXlsx requires Windows because it uses Excel COM automation to write XLSX output.'
}

$resolvedInputFolderPath = Resolve-FolderPath -Path $InputFolderPath

if (-not $PSBoundParameters.ContainsKey('OutputXlsxPath') -or [string]::IsNullOrWhiteSpace($OutputXlsxPath)) {
    $folderLeaf = Split-Path -Path $resolvedInputFolderPath -Leaf
    if ([string]::IsNullOrWhiteSpace($folderLeaf)) {
        $folderLeaf = 'CsvFolder'
    }

    $OutputXlsxPath = Join-Path -Path $resolvedInputFolderPath -ChildPath ("Results_{0}_{1}.xlsx" -f $folderLeaf, (Get-Date -Format 'yyyyMMdd-HHmmss'))
}

$outputDirectory = Split-Path -Path $OutputXlsxPath -Parent
if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

$transcriptPath = $null
if (-not $NoTranscript) {
    $transcriptPath = Start-RunTranscript -OutputFilePath $OutputXlsxPath -ScriptPath $PSCommandPath
}

$excel = $null
$workbook = $null
$usedWorksheetNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

try {
    Write-Status -Message "Scanning '$resolvedInputFolderPath' for CSV files."
    $csvFiles = Get-CsvFiles -FolderPath $resolvedInputFolderPath -Recurse:$Recurse

    if ($csvFiles.Count -eq 0) {
        throw "No CSV files were found in '$resolvedInputFolderPath'."
    }

    if (Test-Path -LiteralPath $OutputXlsxPath) {
        if ($Overwrite) {
            Remove-Item -LiteralPath $OutputXlsxPath -Force
        }
        else {
            throw "Output XLSX already exists: $OutputXlsxPath. Use -Overwrite to replace it."
        }
    }

    try {
        $excel = New-Object -ComObject Excel.Application
    }
    catch {
        throw "Microsoft Excel is required to create XLSX output with this utility. Error: $($_.Exception.Message)"
    }

    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()

    while ($workbook.Worksheets.Count -gt 1) {
        $extraWorksheet = $null
        try {
            $extraWorksheet = $workbook.Worksheets.Item($workbook.Worksheets.Count)
            $extraWorksheet.Delete()
        }
        finally {
            Release-ComObjectSafely -ComObject $extraWorksheet
        }
    }

    $sheetIndex = 0
    foreach ($csvFile in $csvFiles) {
        $worksheet = $null

        try {
            $csvData = Read-CsvFile -CsvPath $csvFile.FullName -Delimiter $Delimiter
            $worksheetBaseName = Get-WorksheetBaseName -CsvFilePath $csvFile.FullName -RootFolderPath $resolvedInputFolderPath
            $worksheetName = Get-UniqueWorksheetName -PreferredName $worksheetBaseName -UsedNames $usedWorksheetNames

            if ($sheetIndex -eq 0) {
                $worksheet = $workbook.Worksheets.Item(1)
            }
            else {
                $worksheet = $workbook.Worksheets.Add([System.Type]::Missing, $workbook.Worksheets.Item($workbook.Worksheets.Count))
            }

            $worksheet.Name = $worksheetName
            Write-WorksheetData -Worksheet $worksheet -Headers $csvData.Headers -Rows $csvData.Rows

            if ($csvData.Headers.Count -eq 0 -and $csvData.Rows.Count -eq 0) {
                Write-Status -Message "Added worksheet '$worksheetName' from empty CSV '$($csvFile.FullName)'." -Level WARN
            }
            else {
                Write-Status -Message "Added worksheet '$worksheetName' from '$($csvFile.FullName)' ($($csvData.Rows.Count) data row(s))."
            }

            $sheetIndex++
        }
        finally {
            Release-ComObjectSafely -ComObject $worksheet
        }
    }

    $workbook.SaveAs($OutputXlsxPath, 51)
    Write-Status -Message "XLSX workbook created at '$OutputXlsxPath'." -Level SUCCESS
}
finally {
    if ($workbook) {
        try {
            $workbook.Close($false)
        }
        catch {
            # Best effort only.
        }
    }

    Release-ComObjectSafely -ComObject $workbook

    if ($excel) {
        try {
            $excel.Quit()
        }
        catch {
            # Best effort only.
        }
    }

    Release-ComObjectSafely -ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if (-not $NoTranscript) {
        Stop-RunTranscript
    }
}
