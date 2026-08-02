<#
.SYNOPSIS
    Extract text (and optionally tables) from a Word document via COM.

.DESCRIPTION
    Opens a .docx/.doc with the installed Word application, prints paragraph
    text to stdout, and always releases COM objects. Read-only: the document
    is opened and closed without saving.

.PARAMETER Path
    Path to the Word document. Relative paths are resolved to absolute.

.PARAMETER IncludeTables
    Also emit each table as pipe-delimited rows.

.PARAMETER OutFile
    When given, write the extracted text to this file as UTF-8 (with BOM)
    instead of stdout. Recommended for CJK content and for any tool/AI that
    will read the output back — avoids console codepage mojibake.

.EXAMPLE
    .\Get-WordText.ps1 -Path .\report.docx

.EXAMPLE
    .\Get-WordText.ps1 -Path C:\docs\report.docx -IncludeTables

.EXAMPLE
    .\Get-WordText.ps1 -Path .\报告.doc -IncludeTables -OutFile .\doc.txt
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Path,

    [switch]$IncludeTables,

    [string]$OutFile
)

$ErrorActionPreference = 'Stop'

# Resolve a relative path against Get-Location; absolute paths pass through.
# NOTE: Join-Path alone is unsafe here — on PS 5.1 it does NOT treat an
# absolute ChildPath as absolute and produces "CWD\C:\abs\path", which then
# makes [IO.Path]::GetFullPath throw NotSupportedException.
function Resolve-AbsPath([string]$p) {
    if ([System.IO.Path]::IsPathRooted($p)) { return $p }
    [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $p))
}

$abs = Resolve-AbsPath $Path
if (-not (Test-Path -LiteralPath $abs)) {
    throw "File not found: $abs"
}

# Collect output so we can either print it or dump it to a UTF-8 file.
$lines = New-Object System.Collections.Generic.List[string]
function Write-Out([string]$text) { $lines.Add($text) }

$word = $null
$doc  = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0   # wdAlertsNone

    $doc = $word.Documents.Open($abs, $false, $true)  # ConfirmConversions, ReadOnly

    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $line = $doc.Paragraphs.Item($i).Range.Text.TrimEnd("`r", "`n", "`a")
        if ($line) { Write-Out $line }
    }

    if ($IncludeTables) {
        for ($t = 1; $t -le $doc.Tables.Count; $t++) {
            $table = $doc.Tables.Item($t)
            Write-Out ""
            Write-Out "--- Table $t ---"
            for ($r = 1; $r -le $table.Rows.Count; $r++) {
                $cells = for ($c = 1; $c -le $table.Columns.Count; $c++) {
                    ($table.Cell($r, $c).Range.Text) -replace "[\r\a]", ""
                }
                Write-Out ($cells -join " | ")
            }
        }
    }
}
finally {
    try { if ($doc)  { $doc.Close($false) } } catch { }
    try { if ($word) { $word.Quit() } } catch { }
    foreach ($o in @($doc, $word)) {
        if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

if ($OutFile) {
    $outAbs = Resolve-AbsPath $OutFile
    [System.IO.File]::WriteAllLines($outAbs, $lines, [System.Text.UTF8Encoding]::new($true))
    Write-Output "Wrote $($lines.Count) lines -> $outAbs"
} else {
    foreach ($l in $lines) { Write-Output $l }
}
