---
name: excel-com-powershell
description: Use when reading, writing, editing, or generating Microsoft Excel workbooks (.xlsx/.xls/.csv) on Windows - drives Excel.Application via PowerShell COM to read/write cells and ranges, set formulas, manage multiple sheets, recalculate, or export to PDF
---

# Excel Automation via PowerShell COM

## Overview

Drive the installed Microsoft Excel application through COM
(`New-Object -ComObject Excel.Application`) to read and write `.xlsx`/`.xls`
workbooks with real formula recalculation and exact formatting.

**Core principle:** Bulk reads/writes go through `Range.Value2` as a 2D array —
touching cells one at a time over COM is extremely slow. And, like all Office
COM, you MUST quit and release every reference in `finally` or you leak
`EXCEL.EXE` processes.

**REQUIRED BACKGROUND:** Read `office-docs:office-com-cleanup` for the release
pattern. Read `office-docs:office-docs-overview` for when to prefer a library
(Excel COM is especially unreliable under service accounts / Session 0).

## When to Use

- Read values, formulas, or whole ranges from a workbook
- Write cells/ranges, add formulas, format, autofit
- Create/iterate multiple worksheets
- Force recalculation so cached formula values are correct
- Export a sheet/workbook to PDF
- Convert between `.xlsx`, `.xls`, `.csv`

For headless/CI or pure tabular data, prefer `openpyxl`/`ClosedXML` — see
`office-docs:office-docs-overview`.

## Critical Constants

| Name | Value | Meaning |
|------|-------|---------|
| `xlOpenXMLWorkbook` | `51` | Save as `.xlsx` |
| `xlOpenXMLWorkbookMacroEnabled` | `52` | `.xlsm` |
| `xlExcel8` | `56` | Legacy `.xls` |
| `xlCSV` | `6` | `.csv` |
| `xlTypePDF` | `0` | PDF export type for `ExportAsFixedFormat` |
| `xlCalculationManual` | `-4135` | Manual calc mode |
| `xlCalculationAutomatic` | `-4105` | Automatic calc mode |
| `xlUp` | `-4162` | Direction for `.End()` (find last row) |
| `xlToLeft` | `-4159` | Direction for `.End()` |

```powershell
$xlOpenXMLWorkbook = 51
$xlCSV = 6
```

## Setup and Teardown (always)

```powershell
$ErrorActionPreference = 'Stop'
$excel = $null
$wb    = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false      # big speedup for writes

    $path = [System.IO.Path]::GetFullPath('C:\data\book.xlsx')  # ABSOLUTE
    $wb = $excel.Workbooks.Open($path)
    $ws = $wb.Worksheets.Item(1)

    # ... work ...

    $wb.Save()
}
finally {
    if ($wb)    { $wb.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($ws, $wb, $excel)) {
        if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
```

## Quick Reference

| Goal | Code |
|------|------|
| New workbook | `$wb = $excel.Workbooks.Add()` |
| Open existing | `$wb = $excel.Workbooks.Open($absPath)` |
| Sheet by index (1-based) | `$ws = $wb.Worksheets.Item(1)` |
| Sheet by name | `$ws = $wb.Worksheets.Item("Data")` |
| Add sheet | `$ws = $wb.Worksheets.Add()` |
| Read one cell | `$ws.Cells.Item($r,$c).Value2` |
| Write one cell | `$ws.Cells.Item($r,$c).Value2 = 42` |
| Read whole used range | `$ws.UsedRange.Value2` (2D array) |
| Last used row | `$ws.Cells.Item($ws.Rows.Count,1).End(-4162).Row` |
| Set a formula | `$ws.Cells.Item($r,$c).Formula = "=SUM(A1:A10)"` |
| Recalculate | `$excel.Calculate()` |
| Autofit columns | `$ws.UsedRange.EntireColumn.AutoFit()` |
| Save as xlsx | `$wb.SaveAs($path, 51)` |
| Save as csv | `$wb.SaveAs($path, 6)` |
| Close (no save) | `$wb.Close($false)` |

> `Value2` is the fast, locale-independent accessor (no currency/date COM
> wrapping). Prefer it over `.Value` and `.Text`.

## Reading

### Whole sheet as a 2D array (fast)
```powershell
$data = $ws.UsedRange.Value2   # [object[,]] 1-based on BOTH dimensions
$rows = $data.GetLength(0)
$cols = $data.GetLength(1)
for ($r = 1; $r -le $rows; $r++) {
    $line = for ($c = 1; $c -le $cols; $c++) { $data[$r, $c] }
    Write-Output ($line -join "`t")
}
```

> When `UsedRange` is a single cell, `Value2` is a scalar, not an array. Guard
> with `if ($data -is [array])`.

### Find the last data row
```powershell
$lastRow = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row   # -4162 = xlUp
```

## Writing

### Write a 2D block in ONE call (fast)
```powershell
$data = @(
    @("Name","Score"),
    @("Ada", 99),
    @("Alan", 98)
)
$rows = $data.Count
$cols = $data[0].Count

# Build a [,] array Excel can consume in a single assignment
$arr = New-Object 'object[,]' $rows, $cols
for ($r = 0; $r -lt $rows; $r++) {
    for ($c = 0; $c -lt $cols; $c++) { $arr[$r, $c] = $data[$r][$c] }
}

$start = $ws.Cells.Item(1, 1)
$end   = $ws.Cells.Item($rows, $cols)
$range = $ws.Range($start, $end)
$range.Value2 = $arr     # single COM round-trip — fast
```

> Writing cell-by-cell in a loop is 100x+ slower. Always batch via a range.

> **CRITICAL — PowerShell `[,]` index parsing pitfall:** In `$arr[$r + 1, 0]`,
> the `,` is the array-construction operator, so PowerShell parses this as
> `$r + (1, 0)` — array concatenation — and throws
> `[System.Object[]] 不包含名为 "op_Addition" 的方法`. **Always** compute
> the row index into a separate variable first:
> ```powershell
> $ri = $r + 1                          # compute once
> $arr[$ri, 0] = $ri                    # then index with a plain integer
> ```
> This applies to ALL `object[,]` indexing in PowerShell, not just Excel scripts.

> **CRITICAL — Excel auto-interprets string values:** When writing strings via
> `Range.Value2`, Excel automatically converts values that look like numbers,
> dates, or percentages. `"002594"` becomes `2594` (drops leading zeros);
> `"+0.93%"` becomes `0.0093` (parsed as percentage). To preserve exact text:
> ```powershell
> # 1. Prefix with apostrophe when building the array (forces text storage)
> $arr[$r, $c] = "'" + $code    # "'002594" stored as text, leading zeros kept
>
> # 2. Also set NumberFormat to '@' (Text) on the target column range
> $ws.Range($ws.Cells(2, 2), $ws.Cells($n, 2)).NumberFormat = '@'
> ```
> Both steps are recommended: the apostrophe ensures the raw value is stored as
> text at write time, and the `NumberFormat = '@'` prevents Excel from
> re-interpreting the value if a user double-clicks the cell later. This
> pattern is essential for stock codes, phone numbers, ID numbers, and any
> numeric-looking string that must not lose leading zeros or formatting.

### Formulas + recalculation
```powershell
$ws.Cells.Item(4, 2).Formula = "=SUM(B2:B3)"
$excel.Calculate()                       # cached value now correct
$result = $ws.Cells.Item(4, 2).Value2
```

### Save / Save As / PDF
```powershell
$wb.Save()                                # in place
$wb.SaveAs("C:\out\book.xlsx", 51)        # 51 = xlOpenXMLWorkbook
$wb.SaveAs("C:\out\data.csv", 6)          # 6  = xlCSV (active sheet only)
$ws.ExportAsFixedFormat(0, "C:\out\sheet.pdf")  # 0 = xlTypePDF
```

## Ready-to-use scripts

- `scripts/Get-ExcelData.ps1` — read a sheet (or all sheets) to objects/CSV.
- `scripts/Set-ExcelData.ps1` — write a 2D array / CSV into a sheet in one call.
- `scripts/Convert-ExcelToPdf.ps1` — export a workbook/sheet to PDF.

All scripts implement the full open/try/finally/release lifecycle.

## Performance Checklist

1. `Visible = $false`, `ScreenUpdating = $false`, `DisplayAlerts = $false`.
2. Read with `UsedRange.Value2`; write with a single `Range.Value2 = [,]`.
3. For very large writes, set `Calculation = -4135` (manual), write, then
   `$excel.Calculate()` and restore `-4105`.
4. Never loop one cell at a time over COM.

## Common Mistakes

| Mistake | Symptom | Fix |
|---------|---------|-----|
| Cell-by-cell loop | Painfully slow | Batch via `Range.Value2 = [,]` |
| `Value2` array assumed 0-based | Off-by-one / null | `UsedRange.Value2` is **1-based** both dims |
| Single-cell `UsedRange` | Indexing a scalar fails | Guard `if ($data -is [array])` |
| Relative path | File saved in wrong place | `GetFullPath()` first |
| Not releasing COM | `EXCEL.EXE` leaks, file locked | `finally` + `ReleaseComObject` + GC |
| Running in CI/service | Flaky COM errors | Use `openpyxl`/`ClosedXML` instead |
| Reading formula expecting value | Got `=SUM(...)` string | Use `.Value2` after `$excel.Calculate()` |
| `$arr[$r+1,0]` with computed index | `op_Addition` / array-concat error | `$ri=$r+1; $arr[$ri,0]` — compute index into a variable first |
| Writing strings that look like numbers via `Value2` | Leading zeros dropped, `"+0.93%"` → `0.0093` | Prefix with `'` and set `NumberFormat = '@'` on the column |
| Non-ASCII `.ps1` saved UTF-8 **without** BOM | Parser error on lines with Chinese/CJK text under PS 5.1 | Save `.ps1` as **UTF-8 with BOM** — see `office-docs:word-com-powershell` → Script File Encoding |
