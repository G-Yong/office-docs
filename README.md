# office-docs

**English** | [中文](README.zh-CN.md)

A skills library that teaches coding agents how to **read and write Microsoft
Office files** — Word (`.docx`/`.doc`) and Excel (`.xlsx`/`.xls`/`.csv`) — on
Windows, primarily by driving the installed Office applications through
**PowerShell COM automation**.

Inspired by the structure of [obra/superpowers](https://github.com/obra/superpowers):
each skill is a self-contained `SKILL.md` with a focused trigger, and the agent
loads it on demand.

## Why COM automation?

Driving the real Word/Excel application via COM gives you exact rendering, real
formula recalculation, native PDF export, and legacy `.doc`/`.xls` support — the
output matches what a user sees. The cost is that it is **Windows-only, requires
Office installed and an interactive session, and demands careful COM cleanup**.
When those constraints don't fit (CI, headless, bulk data), the skills point you
to library-based alternatives.

## Skills

| Skill | Use it when |
|-------|-------------|
| [`office-docs-overview`](skills/office-docs-overview/SKILL.md) | Deciding how to handle an Office file; COM vs library; safety prerequisites |
| [`word-com-powershell`](skills/word-com-powershell/SKILL.md) | Reading/writing/editing/generating Word documents or exporting to PDF |
| [`excel-com-powershell`](skills/excel-com-powershell/SKILL.md) | Reading/writing cells, ranges, formulas, sheets; Excel → PDF |
| [`office-com-cleanup`](skills/office-com-cleanup/SKILL.md) | Leaked `WINWORD.EXE`/`EXCEL.EXE`, locked files, hung scripts, COM release pattern |

### Bundled scripts

Each automation skill ships ready-to-run PowerShell scripts under
`skills/<skill>/scripts/`, all implementing the full open → try/finally →
release lifecycle:

- **Word:** `Get-WordText.ps1`, `Set-WordReplace.ps1`, `Convert-WordToPdf.ps1`, `Invoke-WordTemplate.ps1`
- **Excel:** `Get-ExcelData.ps1`, `Set-ExcelData.ps1`, `Convert-ExcelToPdf.ps1`

Example:

```powershell
# Extract all text and tables from a Word document
.\skills\word-com-powershell\scripts\Get-WordText.ps1 -Path .\report.docx -IncludeTables

# Fill a {{placeholder}} template
.\skills\word-com-powershell\scripts\Invoke-WordTemplate.ps1 `
    -TemplatePath .\invoice.docx -OutPath .\out.docx `
    -Values @{ CLIENT = 'Acme'; TOTAL = '$1,200' }

# Write a CSV into a formatted .xlsx
.\skills\excel-com-powershell\scripts\Set-ExcelData.ps1 -CsvPath .\people.csv -OutPath .\people.xlsx
```

## Requirements

- Windows with Microsoft Office (Word/Excel) installed
- Windows PowerShell 5.1 or PowerShell 7+
- An **interactive** desktop session (COM automation of Office is unsupported
  under Session 0 / service accounts)

## Using these skills

### With agents that support the skills/SKILL.md convention

Point your agent at this repository (or install it as a plugin). Skills are
auto-discovered from the `skills/` directory; the agent reads a `SKILL.md` when
its `description` trigger matches the task. See [AGENTS.md](AGENTS.md).

### Claude Code (plugin)

This repo is packaged as a Claude Code plugin via `.claude-plugin/`. Register it
as a marketplace from your fork and install:

```bash
/plugin marketplace add G-Yong/office-docs
/plugin install office-docs@office-docs
```

### Manually

The skills are just Markdown. Open the relevant `SKILL.md`, follow its Quick
Reference and patterns, and run the bundled scripts directly.

## Safety notes

- Always use **absolute paths** with COM — Office resolves relative paths
  against its own working directory.
- Always **quit and release** COM objects in a `finally` block, or hidden Office
  processes leak and lock files. See `office-com-cleanup`.
- Do not run this automation as a service / in CI — use a library instead
  (alternatives are documented in the overview skill).

## License

[MIT](LICENSE)
