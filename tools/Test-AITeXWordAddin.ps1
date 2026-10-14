[CmdletBinding()]
param(
    [string]$ExpectedTemplate = (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\AI TeX Word Add-in.dotm')
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ExpectedTemplate)) {
    throw "Installed template not found: $ExpectedTemplate"
}

$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Add()
    $document.Content.Text = 'Inline \(C_{be},C_{bc}\); display \[\frac{g_m}{2\pi C_{be}}\].'

    $converted = $word.Run('AITeX_ConvertDocumentSilent')
    $equationCount = $document.OMaths.Count

    if ([int]$converted -ne 2) {
        throw "Expected 2 converted equations, got $converted."
    }
    if ($equationCount -ne 2) {
        throw "Expected 2 Word equations, got $equationCount."
    }

    $inlineType = [int]$document.OMaths.Item(1).Type
    $displayType = [int]$document.OMaths.Item(2).Type
    $combinedXml = $document.OMaths.Item(1).Range.WordOpenXML + $document.OMaths.Item(2).Range.WordOpenXML

    if ($inlineType -ne 1) { throw "The first equation is not inline (type $inlineType)." }
    if ($displayType -ne 0) { throw "The second equation is not display (type $displayType)." }
    if ($combinedXml -notmatch '<m:sSub>') { throw 'Subscript structure was not generated.' }
    if ($combinedXml -notmatch '<m:f>') { throw 'Fraction structure was not generated.' }

    Write-Output 'PASS: converted 2 delimiters into editable Word equations.'
    Write-Output 'PASS: inline/display layout, subscripts, and fraction structure verified.'
} finally {
    if ($null -ne $document) {
        try { $document.Close(0) } catch { }
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
