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
    $document.Content.Text = 'Escaped scripts \(R\_\pi, C\_\pi\); Greek script \(C_\mu\); ordinary scripts \(R_o,C_o\); doubled delimiter \\(x_\alpha\\); operator \(\operatorname{Re}(z)\); display \[\frac{g_m}{2\pi C_{be}}\].'

    $converted = $word.Run('AITeX_ConvertDocumentSilent')
    $equationCount = $document.OMaths.Count

    if ([int]$converted -ne 6) {
        throw "Expected 6 converted equations, got $converted."
    }
    if ($equationCount -ne 6) {
        throw "Expected 6 Word equations, got $equationCount."
    }

    $escapedXml = $document.OMaths.Item(1).Range.WordOpenXML
    $greekXml = $document.OMaths.Item(2).Range.WordOpenXML
    $ordinaryXml = $document.OMaths.Item(3).Range.WordOpenXML
    $doubledXml = $document.OMaths.Item(4).Range.WordOpenXML
    $operatorText = $document.OMaths.Item(5).Range.Text
    $displayXml = $document.OMaths.Item(6).Range.WordOpenXML

    for ($index = 1; $index -le 5; $index++) {
        $inlineType = [int]$document.OMaths.Item($index).Type
        if ($inlineType -ne 1) { throw "Equation $index is not inline (type $inlineType)." }
    }
    $displayType = [int]$document.OMaths.Item(6).Type
    if ($displayType -ne 0) { throw "The sixth equation is not display (type $displayType)." }

    if (([regex]::Matches($escapedXml, '<m:sSub>')).Count -ne 2) {
        throw 'Escaped underscores did not produce two subscript structures.'
    }
    if ($escapedXml -match '<m:t>_</m:t>') {
        throw 'An escaped underscore remained as a literal underscore.'
    }
    if (([regex]::Matches($escapedXml, [string][char]0x03C0)).Count -ne 2) {
        throw 'The two pi symbols were not preserved in escaped scripts.'
    }
    if ($greekXml -notmatch '<m:sSub>' -or $greekXml -notmatch [string][char]0x03BC) {
        throw 'An unbraced Greek command did not produce a complete subscript.'
    }
    if (([regex]::Matches($ordinaryXml, '<m:sSub>')).Count -ne 2) {
        throw 'Ordinary letter subscripts were not generated.'
    }
    if ($doubledXml -notmatch '<m:sSub>' -or $doubledXml -notmatch [string][char]0x03B1) {
        throw 'A doubled delimiter or its Greek subscript was not converted correctly.'
    }
    if ($document.Content.Text -match '\\') {
        throw 'A backslash delimiter or unsupported command remained in the document.'
    }
    if ($operatorText -match 'operatorname') {
        throw '\operatorname was rendered as literal text.'
    }
    if ($displayXml -notmatch '<m:f>') { throw 'Fraction structure was not generated.' }

    Write-Output 'PASS: converted 6 delimiters into editable Word equations.'
    Write-Output 'PASS: escaped input, Greek scripts, ordinary subscripts, operator names, fraction, and layout verified.'
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
