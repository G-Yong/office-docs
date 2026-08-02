<#
.SYNOPSIS
    Export an Excel workbook (or a single sheet) to PDF.

.DESCRIPTION
    Opens a workbook read-only and exports it as PDF via ExportAsFixedFormat.
    Always releases COM objects.

.PARAMETER Path
    Path to the workbook (.xlsx/.xls).

.PARAMETER OutPath
    Destination .pdf path. Defaults to the source name with .pdf extension.

.PARAMETER Sheet
    Optional sheet name or 1-based index. When given, only that sheet is
    exported; otherwise the entire workbook is exported.

.EXAMPLE
    .\Convert-ExcelToPdf.ps1 -Path .\book.xlsx

.EXAMPLE
    .\Convert-ExcelToPdf.ps1 -Path .\book.xlsx -Sheet "Summary" -OutPath .\summary.pdf
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [string]$OutPath,
    [string]$Sheet
)

$ErrorActionPreference = 'Stop'
$xlTypePDF = 0

# Resolve a relative path against Get-Location; absolute paths pass through.
# Join-Path alone is unsafe: on PS 5.1 it concatenates an absolute ChildPath
# onto the base, producing "CWD\C:\abs\path" that GetFullPath rejects.
function Resolve-AbsPath([string]$p) {
    if ([System.IO.Path]::IsPathRooted($p)) { return $p }
    [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $p))
}

$abs = Resolve-AbsPath $Path
if (-not (Test-Path -LiteralPath $abs)) { throw "File not found: $abs" }

if (-not $OutPath) {
    $OutPath = [System.IO.Path]::ChangeExtension($abs, '.pdf')
}
$outAbs = Resolve-AbsPath $OutPath

$excel = $null
$wb    = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($abs, 0, $true)

    if ($PSBoundParameters.ContainsKey('Sheet')) {
        $key = if ($Sheet -match '^\d+$') { [int]$Sheet } else { $Sheet }
        $target = $wb.Worksheets.Item($key)
    } else {
        $target = $wb
    }

    $target.ExportAsFixedFormat($xlTypePDF, $outAbs)
    Write-Output "Exported PDF -> $outAbs"
}
finally {
    try { if ($wb)    { $wb.Close($false) } } catch { }
    try { if ($excel) { $excel.Quit() } } catch { }
    foreach ($o in @($target, $wb, $excel)) {
        if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
