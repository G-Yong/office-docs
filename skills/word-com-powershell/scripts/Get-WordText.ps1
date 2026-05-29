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

.EXAMPLE
    .\Get-WordText.ps1 -Path .\report.docx

.EXAMPLE
    .\Get-WordText.ps1 -Path C:\docs\report.docx -IncludeTables
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Path,

    [switch]$IncludeTables
)

$ErrorActionPreference = 'Stop'

$abs = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
if (-not (Test-Path -LiteralPath $abs)) {
    throw "File not found: $abs"
}

$word = $null
$doc  = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0   # wdAlertsNone

    $doc = $word.Documents.Open($abs, $false, $true)  # ConfirmConversions, ReadOnly

    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $line = $doc.Paragraphs.Item($i).Range.Text.TrimEnd("`r", "`n", "`a")
        if ($line) { Write-Output $line }
    }

    if ($IncludeTables) {
        for ($t = 1; $t -le $doc.Tables.Count; $t++) {
            $table = $doc.Tables.Item($t)
            Write-Output "`n--- Table $t ---"
            for ($r = 1; $r -le $table.Rows.Count; $r++) {
                $cells = for ($c = 1; $c -le $table.Columns.Count; $c++) {
                    ($table.Cell($r, $c).Range.Text) -replace "[\r\a]", ""
                }
                Write-Output ($cells -join " | ")
            }
        }
    }
}
finally {
    if ($doc)  { $doc.Close($false) }
    if ($word) { $word.Quit() }
    foreach ($o in @($doc, $word)) {
        if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
