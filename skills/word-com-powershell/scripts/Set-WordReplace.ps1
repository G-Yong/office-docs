<#
.SYNOPSIS
    Find-and-replace text in a Word document, including headers and footers.

.DESCRIPTION
    Opens a document, replaces all occurrences of -Find with -Replace across the
    main body and every section's headers and footers, then saves. Always
    releases COM objects.

.PARAMETER Path
    Path to the Word document (modified in place unless -OutPath is given).

.PARAMETER Find
    Text to search for (literal, not wildcard).

.PARAMETER Replace
    Replacement text.

.PARAMETER OutPath
    Optional path to save a copy instead of overwriting the original.

.EXAMPLE
    .\Set-WordReplace.ps1 -Path .\contract.docx -Find "{{CLIENT}}" -Replace "Acme Corp"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$Find,
    [Parameter(Mandatory)][AllowEmptyString()][string]$Replace,
    [string]$OutPath
)

$ErrorActionPreference = 'Stop'
$wdReplaceAll = 2
$wdFormatDocumentDefault = 16

$abs = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
if (-not (Test-Path -LiteralPath $abs)) { throw "File not found: $abs" }

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

    $doc = $word.Documents.Open($abs)

    # Main body
    Invoke-Replace $doc.Content $Find $Replace

    # Headers and footers in every section
    foreach ($section in $doc.Sections) {
        foreach ($header in $section.Headers) { Invoke-Replace $header.Range $Find $Replace }
        foreach ($footer in $section.Footers) { Invoke-Replace $footer.Range $Find $Replace }
    }

    if ($OutPath) {
        $out = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $OutPath))
        $doc.SaveAs([ref]$out, [ref]$wdFormatDocumentDefault)
    } else {
        $doc.Save()
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
