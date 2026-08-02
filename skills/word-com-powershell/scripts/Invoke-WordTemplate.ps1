<#
.SYNOPSIS
    Fill a Word template by replacing {{placeholder}} tokens with values.

.DESCRIPTION
    Opens a template document, replaces every {{key}} token (in body, headers,
    and footers) with the matching value from -Values, and saves the result to
    -OutPath (the template itself is never modified). Always releases COM
    objects.

.PARAMETER TemplatePath
    Path to the template .docx containing {{key}} placeholders.

.PARAMETER OutPath
    Path to write the filled document.

.PARAMETER Values
    Hashtable mapping placeholder names (without braces) to replacement text.

.EXAMPLE
    .\Invoke-WordTemplate.ps1 -TemplatePath .\invoice.docx -OutPath .\out.docx `
        -Values @{ CLIENT = 'Acme'; TOTAL = '$1,200'; DATE = '2026-05-29' }
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TemplatePath,
    [Parameter(Mandatory)][string]$OutPath,
    [Parameter(Mandatory)][hashtable]$Values
)

$ErrorActionPreference = 'Stop'
$wdReplaceAll = 2
$wdFormatDocumentDefault = 16

# Resolve a relative path against Get-Location; absolute paths pass through.
# Join-Path alone is unsafe: on PS 5.1 it concatenates an absolute ChildPath
# onto the base, producing "CWD\C:\abs\path" that GetFullPath rejects.
function Resolve-AbsPath([string]$p) {
    if ([System.IO.Path]::IsPathRooted($p)) { return $p }
    [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $p))
}

$tpl = Resolve-AbsPath $TemplatePath
$out = Resolve-AbsPath $OutPath
if (-not (Test-Path -LiteralPath $tpl)) { throw "Template not found: $tpl" }

function Invoke-Replace($range, $find, $replace) {
    $f = $range.Find
    $f.ClearFormatting()
    $f.Replacement.ClearFormatting()
    [void]$f.Execute($find, $false, $false, $false, $false, $false,
                     $true, 1, $false, $replace, $wdReplaceAll)
}

$word = $null
$doc  = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    # Open the template, then immediately save under the output name so the
    # template file is never touched.
    $doc = $word.Documents.Open($tpl)

    foreach ($key in $Values.Keys) {
        $token = "{{${key}}}"
        $value = [string]$Values[$key]
        Invoke-Replace $doc.Content $token $value
        foreach ($section in $doc.Sections) {
            foreach ($h in $section.Headers) { Invoke-Replace $h.Range $token $value }
            foreach ($ft in $section.Footers) { Invoke-Replace $ft.Range $token $value }
        }
    }

    $doc.SaveAs2($out, $wdFormatDocumentDefault)   # SaveAs2: safe with non-ASCII paths
    Write-Output "Wrote: $out"
}
finally {
    try { if ($doc)  { $doc.Close($false) } } catch { }
    try { if ($word) { $word.Quit() } } catch { }
    foreach ($o in @($doc, $word)) {
        if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
