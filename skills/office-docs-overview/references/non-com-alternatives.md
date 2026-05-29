# Non-COM Alternatives for Office Files

Use these when COM automation is not viable (no Windows, no Office installed,
headless CI, service account / Session 0, or high-volume batch jobs).

## Word `.docx`

| Tool | Language | Notes |
|------|----------|-------|
| `python-docx` | Python | Read/write paragraphs, runs, tables, styles. No `.doc`. Cannot render or recalc fields. |
| Open XML SDK (`DocumentFormat.OpenXml`) | .NET / C# / PowerShell | Full control of the OOXML package. Verbose. No rendering. |
| `docx` / `docxtemplater` | Node.js | Template-driven generation. |
| `pandoc` | CLI | Convert between formats; good for text extraction. |

`.docx` is a ZIP of XML parts. For quick read-only text extraction without any
library you can unzip and read `word/document.xml`.

## Excel `.xlsx`

| Tool | Language | Notes |
|------|----------|-------|
| `openpyxl` | Python | Read/write cells, formulas (as text), styles. No recalculation. |
| `pandas` (via `openpyxl`/`xlsxwriter`) | Python | Bulk tabular read/write. |
| `ClosedXML` / Open XML SDK | .NET | `ClosedXML` is high-level; SDK is low-level. |
| `exceljs` / `xlsx` (SheetJS) | Node.js | Cross-platform read/write. |

**Formula recalculation:** none of the library approaches recalculate formulas.
If you write a formula you get the formula string; the cached value is only
updated when a real Excel engine opens and saves the file. If you must have
recalculated values, you need COM (or LibreOffice headless `--convert-to`).

## LibreOffice (cross-platform, no Microsoft Office)

```bash
soffice --headless --convert-to pdf input.docx
soffice --headless --convert-to xlsx input.csv
```

Good for conversion and recalculation on Linux/macOS, but rendering fidelity
differs slightly from Microsoft Office.

## Decision summary

- Need exact Microsoft rendering / PDF / recalculated formulas → **COM** (these skills).
- Headless / cross-platform / bulk data only → **library**.
- Just need text out of a `.docx` → unzip + read `word/document.xml`, or `pandoc`.
