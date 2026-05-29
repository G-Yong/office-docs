<#
.SYNOPSIS
    Convert one or more Word documents to PDF using the Word engine.

.DESCRIPTION
    Accepts a file or a directory. For a directory, converts every .docx/.doc
    found (optionally recursively). Reuses a single Word instance for the whole
    batch and always releases COM objects.

.PARAMETER Path
    A .docx/.doc file, or a folder containing them.

.PARAMETER Recurse
    When -Path is a folder, also search subfolders.

.PARAMETER OutDir
    Output directory for PDFs. Defaults to each source file's folder.

.EXAMPLE
    .\Convert-WordToPdf.ps1 -Path .\report.docx

.EXAMPLE
    .\Convert-WordToPdf.ps1 -Path C:\docs -Recurse -OutDir C:\pdfs
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [switch]$Recurse,
    [string]$OutDir
)

$ErrorActionPreference = 'Stop'
$wdFormatPDF = 17

$abs = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
if (-not (Test-Path -LiteralPath $abs)) { throw "Path not found: $abs" }

if (Test-Path -LiteralPath $abs -PathType Container) {
    $files = Get-ChildItem -LiteralPath $abs -Recurse:$Recurse -File |
             Where-Object { $_.Extension -in '.docx', '.doc' }
} else {
    $files = @(Get-Item -LiteralPath $abs)
}

if (-not $files) { Write-Warning "No Word documents found."; return }

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    foreach ($file in $files) {
        $targetDir = if ($OutDir) {
            [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $OutDir))
        } else { $file.DirectoryName }
        if (-not (Test-Path -LiteralPath $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }
        $pdf = Join-Path $targetDir ($file.BaseName + '.pdf')

        $doc = $null
        try {
            $doc = $word.Documents.Open($file.FullName, $false, $true)
            $doc.SaveAs([ref]$pdf, [ref]$wdFormatPDF)
            Write-Output "Converted: $($file.Name) -> $pdf"
        }
        finally {
            if ($doc) {
                $doc.Close($false)
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
            }
        }
    }
}
finally {
    if ($word) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
