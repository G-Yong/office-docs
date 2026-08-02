<#
.SYNOPSIS
    Write tabular data into an Excel sheet in a single fast COM call.

.DESCRIPTION
    Reads rows from a CSV file (or accepts the path to one), writes them as a 2D
    block starting at A1 of the target sheet, autofits columns, and saves as
    .xlsx. Creates the workbook if it does not exist. Always releases COM
    objects.

.PARAMETER CsvPath
    Source CSV whose rows/columns are written to the sheet.

.PARAMETER OutPath
    Destination .xlsx path.

.PARAMETER Sheet
    Target sheet name (created if missing). Defaults to "Sheet1".

.EXAMPLE
    .\Set-ExcelData.ps1 -CsvPath .\people.csv -OutPath .\people.xlsx
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$CsvPath,
    [Parameter(Mandatory)][string]$OutPath,
    [string]$Sheet = 'Sheet1'
)

$ErrorActionPreference = 'Stop'
$xlOpenXMLWorkbook = 51

# Resolve a relative path against Get-Location; absolute paths pass through.
# Join-Path alone is unsafe: on PS 5.1 it concatenates an absolute ChildPath
# onto the base, producing "CWD\C:\abs\path" that GetFullPath rejects.
function Resolve-AbsPath([string]$p) {
    if ([System.IO.Path]::IsPathRooted($p)) { return $p }
    [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $p))
}

$csvAbs = Resolve-AbsPath $CsvPath
$outAbs = Resolve-AbsPath $OutPath
if (-not (Test-Path -LiteralPath $csvAbs)) { throw "CSV not found: $csvAbs" }

# Read CSV preserving column order; build a header + data matrix.
$records = Import-Csv -LiteralPath $csvAbs
if (-not $records) { throw "CSV is empty: $csvAbs" }
$headers = $records[0].PSObject.Properties.Name
$rows = $records.Count + 1
$cols = $headers.Count

$arr = New-Object 'object[,]' $rows, $cols
for ($c = 0; $c -lt $cols; $c++) { $arr[0, $c] = $headers[$c] }
for ($r = 0; $r -lt $records.Count; $r++) {
    for ($c = 0; $c -lt $cols; $c++) {
        # Parenthesize ($r + 1): PS 5.1 parses `$arr[$r + 1, $c]` as
        # `$r + (1, $c)` (int + array -> op_Addition error) on multidim arrays.
        $arr[($r + 1), $c] = $records[$r].$($headers[$c])
    }
}

$excel = $null
$wb    = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = $Sheet

    $start = $ws.Cells.Item(1, 1)
    $end   = $ws.Cells.Item($rows, $cols)
    $range = $ws.Range($start, $end)
    $range.Value2 = $arr                       # one COM round-trip
    $ws.Rows.Item(1).Font.Bold = $true
    $ws.UsedRange.EntireColumn.AutoFit() | Out-Null

    $wb.SaveAs($outAbs, $xlOpenXMLWorkbook)
    Write-Output "Wrote $($records.Count) rows -> $outAbs"
}
finally {
    try { if ($wb)    { $wb.Close($false) } } catch { }
    try { if ($excel) { $excel.Quit() } } catch { }
    foreach ($o in @($ws, $wb, $excel)) {
        if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
