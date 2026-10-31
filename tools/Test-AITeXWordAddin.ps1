[CmdletBinding()]
param(
    [string]$ExpectedTemplate = (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\AI TeX Word Add-in.dotm'),
    [string]$ReviewDirectory
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ExpectedTemplate)) {
    throw "Installed template not found: $ExpectedTemplate"
}

$word = $null
$document = $null
$selectionDocument = $null
$compatibilityDocument = $null
$corpusDocument = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    for ($index = 1; $index -le $word.AddIns.Count; $index++) {
        $addin = $word.AddIns.Item($index)
        if ($addin.Name -eq 'AI TeX Word Add-in.dotm') { $addin.Installed = $false }
        # As in the builder, old protected MathType projects can hang SaveAs2.
        if ($addin.Name -like 'MathType Commands*') { $addin.Installed = $false }
    }
    [void]$word.AddIns.Add($ExpectedTemplate, $true)

    $document = $word.Documents.Add()
    # Office suppresses automatic macros in a hidden COM test instance.
    # Invoke the same startup entry point explicitly before verifying it.
    $word.Run('AutoExec')
    $selectionShortcutCode = 512 + 1024 + 256 + [int][char]'E'
    $selectionBinding = $word.FindKey($selectionShortcutCode)
    if ($selectionBinding.Command -notmatch 'AITeX_ConvertSelectionMacro') {
        throw "Selection shortcut was not registered; current command: $($selectionBinding.Command)"
    }

    $document.Content.Text = 'Escaped scripts \(R\_\pi, C\_\pi\); Greek script \(C_\mu\); ordinary scripts \(R_o,C_o\); doubled delimiter \\(x_\alpha\\); operator \(\operatorname{Re}(z)\); display \[\frac{g_m}{2\pi C_{be}}\].'

    $converted = $word.Run('AITeX_ConvertDocumentSilent')
    $equationCount = $document.OMaths.Count

    if ([int]$converted -ne 6) {
        throw "Expected 6 converted equations, got $converted. $($word.Run('AITeX_GetLastReport'))"
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
    if ([int]$word.Run('AITeX_ConvertDocumentSilent') -ne 0 -or $document.OMaths.Count -ne 6) {
        throw 'A second conversion changed existing native equations.'
    }

    $selectionDocument = $word.Documents.Add()
    $selectionSource = 'Keep \(x_1\), leave \(x_2\).'
    $selectionDocument.Content.Text = $selectionSource
    $selectionDocument.Activate()
    $selectionStart = $selectionSource.IndexOf('\(')
    $selectionEnd = $selectionSource.IndexOf('\)') + 2
    $word.Selection.SetRange($selectionStart, $selectionEnd)
    $word.Run('AITeX_ConvertSelectionMacro')

    if ($selectionDocument.OMaths.Count -ne 1) {
        throw "The shortcut target macro converted $($selectionDocument.OMaths.Count) equations instead of one."
    }
    if (-not $selectionDocument.Content.Text.Contains('\(x_2\)')) {
        throw 'The selection macro changed TeX outside the selected range.'
    }

    $compatibilityDocument = $word.Documents.Add()
    $compatibilityDocument.Content.Text = 'Parameters \\(r\_{b''e},\ g_m,\ C\_{b''e},\ C\_{b''c},\ h\_{fe}\\). Matrix \\(\begin{bmatrix} V_1 \\ I_2 \end{bmatrix} = \begin{bmatrix} h\_{11} & h\_{12} \\ h\_{21} & h\_{22} \end{bmatrix} \begin{bmatrix} I_1 \\ V_2 \end{bmatrix}\\).'
    $compatibilityDocument.Activate()
    $compatibilityConverted = $word.Run('AITeX_ConvertDocumentSilent')

    if ([int]$compatibilityConverted -ne 2 -or $compatibilityDocument.OMaths.Count -ne 2) {
        throw "Spacing/matrix sample did not produce 2 equations (converted $compatibilityConverted, found $($compatibilityDocument.OMaths.Count))."
    }

    $parameterXml = $compatibilityDocument.OMaths.Item(1).Range.WordOpenXML
    $matrixXml = $compatibilityDocument.OMaths.Item(2).Range.WordOpenXML
    if ($parameterXml.Contains([string][char]0xFFFD)) {
        throw 'A TeX spacing command produced a replacement character.'
    }
    if (([regex]::Matches($parameterXml, '<m:sSub>')).Count -ne 5) {
        throw 'The parameter sample did not produce five subscripts.'
    }
    if ($matrixXml.Contains('AITEXEMPTYCOL') -or
        $matrixXml.Contains('\begin') -or $matrixXml.Contains('\end')) {
        throw 'Matrix conversion left an internal marker or raw environment text.'
    }

    [xml]$matrixDocumentXml = $matrixXml
    $matrixNamespaces = New-Object Xml.XmlNamespaceManager($matrixDocumentXml.NameTable)
    $matrixNamespaces.AddNamespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
    $matrixNodes = @($matrixDocumentXml.SelectNodes('//m:m', $matrixNamespaces))
    if ($matrixNodes.Count -ne 3) {
        throw "Expected 3 native matrices, got $($matrixNodes.Count)."
    }
    $expectedColumns = @(1, 2, 1)
    for ($matrixIndex = 0; $matrixIndex -lt $matrixNodes.Count; $matrixIndex++) {
        $rows = @($matrixNodes[$matrixIndex].SelectNodes('./m:mr', $matrixNamespaces))
        if ($rows.Count -ne 2) {
            throw "Matrix $($matrixIndex + 1) has $($rows.Count) rows instead of 2."
        }
        foreach ($row in $rows) {
            $cellCount = @($row.SelectNodes('./m:e', $matrixNamespaces)).Count
            if ($cellCount -ne $expectedColumns[$matrixIndex]) {
                throw "Matrix $($matrixIndex + 1) has $cellCount columns instead of $($expectedColumns[$matrixIndex])."
            }
        }
    }

    $cases = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\tests\formulas.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $validCases = @($cases | Where-Object { -not $_.error })
    $corpusDocument = $word.Documents.Add()
    $corpusSource = (($validCases | ForEach-Object { $_.id + "`r\[" + $_.tex.Replace("`n", "`r") + "\]`r" }) -join "`r")
    $invalidSource = 'Keep \(\ThisCommandDoesNotExist{x}\), broken \[\frac{a}{b\], dollars $100 and $200.'
    $corpusDocument.Content.Text = $corpusSource + $invalidSource
    $beforeConversion = $corpusDocument.Content.Text
    $corpusDocument.Activate()
    $convertedCorpus = $word.Run('AITeX_ConvertDocumentSilent')
    if ([int]$convertedCorpus -ne $validCases.Count -or $corpusDocument.OMaths.Count -ne $validCases.Count) {
        throw "Corpus insertion failed: $convertedCorpus / $($validCases.Count). $($word.Run('AITeX_GetLastReport'))"
    }
    if (-not $corpusDocument.Content.Text.Contains($invalidSource)) { throw 'Invalid formulas or dollar text changed.' }
    if ($word.Run('AITeX_GetLastReport') -notmatch 'ThisCommandDoesNotExist') { throw 'Parser diagnostics are missing.' }
    for ($caseIndex = 0; $caseIndex -lt $validCases.Count; $caseIndex++) {
        $case = $validCases[$caseIndex]
        [xml]$native = $corpusDocument.OMaths.Item($caseIndex + 1).Range.WordOpenXML
        $ns = New-Object Xml.XmlNamespaceManager($native.NameTable)
        $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
        $math = $native.SelectSingleNode('//m:oMath', $ns)
        if ($case.xpath -and $math.SelectNodes($case.xpath, $ns).Count -ne $case.count) {
            throw "Word changed the expected native structure: $($case.id)."
        }
        foreach ($check in $case.checks) {
            if ($math.SelectNodes($check.xpath, $ns).Count -ne $check.count) {
                throw "Word changed $($case.id): $($check.xpath)."
            }
        }
        if ($case.text) {
            $nativeText = ($math.SelectNodes('.//m:t', $ns) | ForEach-Object { $_.InnerText }) -join ''
            $pattern = ($case.text.ToCharArray() | ForEach-Object { [regex]::Escape([string]$_) }) -join '.*'
            if ($nativeText -cnotmatch $pattern) { throw "Word lost or reordered text in $($case.id)." }
        }
    }
    Write-Output 'Checking that one Undo restores the source...'
    [void]$corpusDocument.Undo(1)
    if ($corpusDocument.OMaths.Count -ne 0 -or $corpusDocument.Content.Text -cne $beforeConversion) {
        $afterUndo = $corpusDocument.Content.Text
        $firstDifference = 0
        while ($firstDifference -lt [Math]::Min($afterUndo.Length, $beforeConversion.Length) -and
            $afterUndo[$firstDifference] -ceq $beforeConversion[$firstDifference]) { $firstDifference++ }
        throw "One Undo did not restore the original. Equations: $($corpusDocument.OMaths.Count); lengths: $($afterUndo.Length)/$($beforeConversion.Length); first difference: $firstDifference; after: $($afterUndo.Substring($firstDifference, [Math]::Min(100, $afterUndo.Length - $firstDifference)))."
    }
    if ($ReviewDirectory) {
        [void]$corpusDocument.Redo(1)
        if ($corpusDocument.OMaths.Count -ne $validCases.Count) { throw 'Redo did not restore native equations.' }
        Write-Output 'Saving the Word/PDF formula coverage artifacts...'
        New-Item -ItemType Directory -Path $ReviewDirectory -Force | Out-Null
        # Pass plain typed strings to COM, rather than pipeline-wrapped values.
        [string]$reviewDocx = Join-Path $ReviewDirectory 'formula-coverage.docx'
        [string]$reviewPdf = Join-Path $ReviewDirectory 'formula-coverage.pdf'
        $corpusDocument.SaveAs2($reviewDocx, 16)
        $corpusDocument.ExportAsFixedFormat($reviewPdf, 17)
    }
    Write-Output "PASS: $($validCases.Count) formula families verified inside Word; failures preserved; one Undo restores the original."
    Write-Output 'PASS: converted 6 delimiters into editable Word equations.'
    Write-Output 'PASS: escaped input, Greek scripts, ordinary subscripts, operator names, fraction, and layout verified.'
    Write-Output 'PASS: Ctrl+Alt+Shift+E is registered and converts only the selected range.'
    Write-Output 'PASS: TeX spacing and bmatrix environments produce clean native Word equations.'
} finally {
    if ($null -ne $corpusDocument) {
        try { $corpusDocument.Close(0) } catch { }
    }
    if ($null -ne $compatibilityDocument) {
        try { $compatibilityDocument.Close(0) } catch { }
    }
    if ($null -ne $selectionDocument) {
        try { $selectionDocument.Close(0) } catch { }
    }
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
