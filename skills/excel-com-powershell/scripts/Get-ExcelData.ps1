<#
.SYNOPSIS
    Read an Excel worksheet (or all sheets) and emit the cell grid.

.DESCRIPTION
    Opens a workbook read-only, reads the used range via Value2 (fast), and
    writes tab-separated rows to stdout. Always releases COM objects.

.PARAMETER Path
    Path to the workbook (.xlsx/.xls). Relative paths resolved to absolute.

.PARAMETER Sheet
    Sheet name or 1-based index to read. Omit to read every sheet.

.EXAMPLE
    .\Get-ExcelData.ps1 -Path .\book.xlsx -Sheet 1

.EXAMPLE
    .\Get-ExcelData.ps1 -Path C:\data\book.xlsx
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [string]$Sheet
)

$ErrorActionPreference = 'Stop'

$abs = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
if (-not (Test-Path -LiteralPath $abs)) { throw "File not found: $abs" }

function Write-Sheet($ws) {
    Write-Output "--- Sheet: $($ws.Name) ---"
    $data = $ws.UsedRange.Value2
    if ($null -eq $data) { return }
    if (-not ($data -is [array])) { Write-Output ([string]$data); return }
    $rows = $data.GetLength(0)
    $cols = $data.GetLength(1)
    for ($r = 1; $r -le $rows; $r++) {
        $line = for ($c = 1; $c -le $cols; $c++) { [string]$data[$r, $c] }
        Write-Output ($line -join "`t")
    }
}

$excel = $null
$wb    = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($abs, 0, $true)  # UpdateLinks=0, ReadOnly=$true

    if ($PSBoundParameters.ContainsKey('Sheet')) {
        $key = if ($Sheet -match '^\d+$') { [int]$Sheet } else { $Sheet }
        Write-Sheet $wb.Worksheets.Item($key)
    } else {
        foreach ($ws in $wb.Worksheets) { Write-Sheet $ws }
    }
}
finally {
    if ($wb)    { $wb.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($wb, $excel)) {
        if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
