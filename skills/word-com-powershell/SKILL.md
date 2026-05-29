---
name: word-com-powershell
description: Use when reading, writing, editing, or generating Microsoft Word documents (.docx/.doc) on Windows - drives Word.Application via PowerShell COM to extract text/tables, find-and-replace, insert content, fill templates, or export to PDF
---

# Word Automation via PowerShell COM

## Overview

Drive the installed Microsoft Word application through COM
(`New-Object -ComObject Word.Application`) to read and write `.docx`/`.doc`
files with full fidelity — real styles, fields, tables, and PDF export.

**Core principle:** You are remote-controlling a real Word instance. It is
invisible but fully alive, so you MUST quit it and release every COM reference
in a `finally` block or you leak `WINWORD.EXE` processes and lock files.

**REQUIRED BACKGROUND:** Read `office-docs:office-com-cleanup` for the object
release pattern. Read `office-docs:office-docs-overview` for when COM is the
wrong choice (CI, no Office, bulk data).

> **WPS note:** On a machine with WPS Office (and no Microsoft Office),
> `New-Object -ComObject Word.Application` may launch **WPS Writer** instead —
> WPS registers under the Microsoft ProgIDs. Check `$word.Path` to confirm. WPS
> covers common automation but implements only a subset of the API, and its
> hidden process is **`wps.exe`**, not `WINWORD.EXE`. See
> `office-docs:office-docs-overview`.

## When to Use

- Extract text or tables from `.docx`/`.doc`
- Find-and-replace across a document (including in headers/footers)
- Insert paragraphs, tables, images, page breaks
- Fill a template (placeholder replacement or bookmarks)
- Convert Word → PDF using the real Word engine
- Anything needing exact Word formatting or legacy `.doc` support

For headless/CI or pure data, use a library instead — see
`office-docs:office-docs-overview`.

## Critical Constants

Word's COM API uses magic numbers. The important ones:

| Name | Value | Meaning |
|------|-------|---------|
| `wdFormatDocumentDefault` | `16` | Save as `.docx` |
| `wdFormatPDF` | `17` | Export PDF |
| `wdFormatXMLDocument` | `12` | `.docx` (Word 2007 XML) |
| `wdFormatDocument97` | `0` | Legacy `.doc` |
| `wdReplaceAll` | `2` | Replace all matches |
| `wdReplaceOne` | `1` | Replace first match |
| `wdAlertsNone` | `0` | Suppress dialogs |
| `wdStory` | `6` | Move cursor to start/end of document |
| `wdGoToPage` | `1` | GoTo page unit |

PowerShell does not know these names — use the integers, or define them once:

```powershell
$wdFormatDocumentDefault = 16
$wdFormatPDF = 17
$wdReplaceAll = 2
```

### wdBuiltinStyle constants (for `$sel.Style` / `.Range.Style`)

Style names passed as strings (e.g. `"Heading 1"`, `"Normal"`) are
**locale-sensitive**: they vary between English, Chinese, and other Office
installations, causing `E_FAIL` on non-English Word. Use the negative `WdBuiltinStyle`
integers instead — they work on every locale:

| Name | Value | Meaning |
|------|-------|---------|
| `wdStyleNormal` | `-1` | Normal (body text) |
| `wdStyleHeading1` | `-2` | Heading 1 |
| `wdStyleHeading2` | `-3` | Heading 2 |
| `wdStyleHeading3` | `-4` | Heading 3 |
| `wdStyleDefaultParagraphFont` | `-66` | Default Paragraph Font |
| `wdStyleTableGrid` | `-155` | Table Grid |

> **IMPORTANT:** `Styles.Item()` with an integer uses these `WdBuiltinStyle` enum
> values. Passing `0` or any positive integer that is not a valid enum member
> throws `E_ACCESSDENIED`. Always use the negative constants above.

```powershell
# ✅ Locale-safe style assignment
$sel.Style = $doc.Styles.Item(-2)   # Heading 1 on any locale
$sel.TypeText("My Title")
$sel.TypeParagraph()
$sel.Style = $doc.Styles.Item(-1)   # Normal
```

## Script File Encoding (CRITICAL for non-ASCII / Chinese / CJK)

**Windows PowerShell 5.1 decodes a `.ps1` file that has NO byte-order mark
(BOM) using the system ANSI codepage** (e.g. CP936/GBK on Chinese Windows,
CP932 on Japanese), **not UTF-8.** If your script contains non-ASCII literals
— Chinese titles, table headers, JSON property names like `$s.代码` — and is
saved as UTF-8 *without* a BOM, PS 5.1 mis-decodes those bytes and the script
fails to even parse, with a misleading error like:

```
所在位置 ...\script.ps1:37 字符: 55   (parser error / unexpected token)
Command exited with code 1
```

Most editors and file-writing tools save UTF-8 **without** a BOM by default, so
this bites silently.

**Rule:** Any `.ps1` containing non-ASCII characters that must run under Windows
PowerShell 5.1 MUST be saved as **UTF-8 with BOM** (or UTF-16 LE). Verify the
first three bytes are `239,187,191` (`EF BB BF`).

After writing a script with a tool that omits the BOM, re-encode it:

```powershell
$p = 'C:\path\script.ps1'
$txt = [System.IO.File]::ReadAllText($p, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText($p, $txt, [System.Text.UTF8Encoding]::new($true))  # $true = emit BOM
# verify:
([System.IO.File]::ReadAllBytes($p))[0..2]   # -> 239 187 191
```

The same applies to data files you read with `Get-Content`: keep them UTF-8
*with* BOM, or pass `-Encoding UTF8` explicitly (a BOM-less UTF-8 data file read
without `-Encoding UTF8` is also mis-decoded as ANSI).

> PowerShell 7+ (`pwsh`) defaults to UTF-8 and does not need the BOM, but the
> system `powershell.exe` (5.1) is what most Windows boxes invoke — assume 5.1
> and add the BOM.

## Setup and Teardown (always)

```powershell
$ErrorActionPreference = 'Stop'
$word = $null
$doc  = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0            # wdAlertsNone — no modal dialogs

    $path = [System.IO.Path]::GetFullPath('C:\reports\report.docx')  # ABSOLUTE
    $doc = $word.Documents.Open($path)

    # ... work ...

    $doc.Save()
}
finally {
    try { if ($doc)  { $doc.Close($false) } } catch { }   # $false = don't save again
    try { if ($word) { $word.Quit() } } catch { }
    # Release refs + GC — see office-docs:office-com-cleanup
    foreach ($o in @($doc, $word)) {
        if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
```

## Quick Reference

| Goal | Code |
|------|------|
| New blank document | `$doc = $word.Documents.Add()` |
| Open existing | `$doc = $word.Documents.Open($absPath)` |
| All text | `$doc.Content.Text` |
| Paragraph count | `$doc.Paragraphs.Count` |
| Nth paragraph text | `$doc.Paragraphs.Item($i).Range.Text` |
| Append text | `$doc.Content.InsertAfter("text")` |
| Append paragraph | `$doc.Content.InsertParagraphAfter()` |
| Table count | `$doc.Tables.Count` |
| Cell text | `$doc.Tables.Item(1).Cell($row,$col).Range.Text` |
| Save as docx | `$doc.SaveAs2($path, 16)` |
| Export PDF | `$doc.SaveAs2($pdfPath, 17)` |
| Close (no save) | `$doc.Close($false)` |

> Cell text from Word ends with a `\r\a` (bell) marker. Strip it:
> `$cell.Range.Text -replace "[\r\a]", ""`.
>
> **When writing to a cell,** `.Range.Text` only accepts `[string]`. PowerShell
> integers cause `InvalidCastException`. Always cast: `$cell.Range.Text = [string]$value`.

## Reading

### Extract all text
```powershell
$text = $doc.Content.Text
```

### Iterate paragraphs (1-based!)
```powershell
for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
    $p = $doc.Paragraphs.Item($i).Range.Text.TrimEnd("`r","`n")
    if ($p) { Write-Output $p }
}
```

### Read a table into objects
```powershell
$table = $doc.Tables.Item(1)
$rows = $table.Rows.Count
$cols = $table.Columns.Count
for ($r = 1; $r -le $rows; $r++) {
    $cells = for ($c = 1; $c -le $cols; $c++) {
        ($table.Cell($r, $c).Range.Text) -replace "[\r\a]", ""
    }
    Write-Output ($cells -join " | ")
}
```

## Writing

### Find and replace (whole document)
```powershell
$find = $doc.Content.Find
# Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
#         MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace)
$null = $find.Execute("{{NAME}}", $false, $false, $false, $false, $false,
                      $true, 1, $false, "Ada Lovelace", 2)  # 2 = wdReplaceAll
```

> `$doc.Content.Find` only covers the main story. To also replace in
> headers/footers, iterate `$doc.Sections` and run `Find` on each
> `$section.Headers` / `$section.Footers` range. See
> `scripts/Set-WordReplace.ps1`.

### Insert a heading + paragraph at the end

**On a NEW blank document** (no existing content), setting `$range.Style` on a
collapsed Range throws `HRESULT E_FAIL`. Use `$word.Selection` or font
properties instead:

```powershell
# ✅ Safe on new blank docs — use Selection with locale-safe numeric style constants
$sel = $word.Selection
$sel.Style = $doc.Styles.Item(-2)   # -2 = wdStyleHeading1 (works on any locale)
$sel.TypeText("Summary")
$sel.TypeParagraph()
$sel.Style = $doc.Styles.Item(-1)   # -1 = wdStyleNormal
$sel.TypeText("Body text here.")
$sel.TypeParagraph()
```

> String style names like `"Heading 1"` or `"Normal"` are locale-sensitive and
> will fail (E_FAIL / item not found) on non-English Office. Always prefer
> negative `WdBuiltinStyle` integer constants — see the Critical Constants section.

**On an EXISTING document** with content, `$doc.Content` works because the
Range has content to apply the style to:

```powershell
$range = $doc.Content
$range.Collapse(0)                      # 0 = wdCollapseEnd
$range.InsertAfter("Summary`r")         # insert text first
$range.Paragraphs.Item(1).Range.Style = "Heading 1"  # style the paragraph that now exists
$range.Collapse(0)
$range.InsertAfter("Body text here.`r")
```

> **Why:** On a blank doc, `Collapse(0)` reduces the Range to a zero-width point.
> Setting `.Style` on a collapsed Range throws `HRESULT E_FAIL`. `Selection`
> auto-expands when you type into it, so it works reliably in both cases.

### Insert a table from data
```powershell
$data = @(
    @("Name","Score"),
    @("Ada","99"),
    @("Alan","98")
)
$range = $doc.Content
$range.Collapse(0)
$table = $doc.Tables.Add($range, $data.Count, $data[0].Count)
$table.Borders.Enable = $true
for ($r = 0; $r -lt $data.Count; $r++) {
    for ($c = 0; $c -lt $data[$r].Count; $c++) {
        $table.Cell($r + 1, $c + 1).Range.Text = [string]$data[$r][$c]
    }
}
```

### Save / Save As
```powershell
$doc.Save()                                   # save in place
$doc.SaveAs2($docxPath, 16)                   # 16 = wdFormatDocumentDefault
$doc.SaveAs2($pdfPath,  17)                   # 17 = wdFormatPDF
```

> **Prefer `SaveAs2` over `SaveAs([ref]...)`.** The `[ref]` pattern with
> `SaveAs` can throw `ArgumentException` ("值不在预期的范围内") when the path
> contains non-ASCII characters (Chinese/CJK). `SaveAs2` accepts plain
> `[string]` arguments and works reliably with any path.

## Ready-to-use scripts

- `scripts/Get-WordText.ps1` — extract all text (and optionally tables) from a doc.
- `scripts/Set-WordReplace.ps1` — find/replace across body, headers, and footers.
- `scripts/Convert-WordToPdf.ps1` — batch-convert `.docx` → `.pdf`.
- `scripts/Invoke-WordTemplate.ps1` — fill a `{{placeholder}}` template and save.

All scripts implement the full open/try/finally/release lifecycle. Run any with
`-?` or read the header comment for parameters.

## Common Mistakes

| Mistake | Symptom | Fix |
|---------|---------|-----|
| Relative path to `Open`/`SaveAs` | File "disappears" or opens wrong file | `[System.IO.Path]::GetFullPath()` first |
| `DisplayAlerts` left on | Script hangs invisibly | `$word.DisplayAlerts = 0` |
| 0-based collection index | "Item -1 not found" / off-by-one | All COM collections are **1-based** |
| Not releasing COM refs | `WINWORD.EXE` piles up, file locked | `finally` + `ReleaseComObject` + GC |
| `Close()` without arg after edits | Hidden "save changes?" prompt | `$doc.Close($false)` or `$doc.Save()` first |
| Reading cell text raw | Trailing `\r\a` junk | `-replace "[\r\a]", ""` |
| Expecting it to work in CI | Random COM errors / no output | Use a library (see overview) |
| Non-ASCII script saved UTF-8 **without** BOM | Parser error at the line with Chinese/CJK text; `exit code 1` under PS 5.1 | Save `.ps1` as **UTF-8 with BOM** (see Script File Encoding) |
| `$word.Quit()` throws RPC failure `0x800706BE` in `finally` | Doc still saved fine, but script exits non-zero | Wrap `Quit()` in its own `try{}catch{}`; confirm no orphan `WINWORD.EXE` |
| `$range.Style = "Heading 1"` on blank new doc | `HRESULT E_FAIL` — collapsed Range has no paragraph to style | Use `$word.Selection` + `TypeText()` on new docs; or insert text first then style via `$range.Paragraphs.Item(1).Range.Style` |
| `$doc.SaveAs([ref]$path, [ref]16)` with Chinese path | `ArgumentException` ("值不在预期的范围内") — `[ref]` type coercion unstable with non-ASCII | Use `$doc.SaveAs2($path, 16)` (accepts plain `[string]`) |
| `$sel.Style = "Heading 1"` on non-English Office | `E_FAIL` / item not found — style names are locale-specific strings | Use `WdBuiltinStyle` integer: `$doc.Styles.Item(-2)` for Heading 1, `Item(-1)` for Normal |
| `$doc.Styles.Item(0)` or any non-negative integer | `E_ACCESSDENIED` — `0` / positive indices don't map to valid built-in styles | All built-in styles use **negative** enum values (see Critical Constants → wdBuiltinStyle) |
