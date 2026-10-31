[CmdletBinding()]
param([string]$ExpectedTemplate = (Join-Path $PSScriptRoot '..\dist\AI TeX Word Add-in.dotm'))
$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
$failures = @()
$properties = @('SpaceBefore', 'SpaceAfter', 'SpaceBeforeAuto', 'SpaceAfterAuto', 'LineSpacingRule', 'LineSpacing',
    'LeftIndent', 'RightIndent', 'FirstLineIndent', 'Alignment', 'KeepWithNext', 'KeepTogether', 'WidowControl')

function Get-ParagraphSnapshot($doc) {
    $snapshot = @()
    foreach ($paragraph in $doc.Paragraphs) {
        $item = [ordered]@{ Style = $paragraph.Style.NameLocal }
        foreach ($property in $properties) { $item[$property] = $paragraph.Format.$property }
        $snapshot += [pscustomobject]$item
    }
    return $snapshot
}

function Assert-ParagraphSnapshot($doc, $expected) {
    $actual = @(Get-ParagraphSnapshot $doc)
    if ($actual.Count -ne $expected.Count) { throw "Paragraph count: $($actual.Count) / $($expected.Count)." }
    for ($i = 0; $i -lt $expected.Count; $i++) {
        foreach ($property in @('Style') + $properties) {
            if ($actual[$i].$property -ne $expected[$i].$property) {
                throw "Paragraph $($i + 1), ${property}: $($expected[$i].$property) -> $($actual[$i].$property)."
            }
        }
    }
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    foreach ($addin in $word.AddIns) {
        if ($addin.Name -eq 'AI TeX Word Add-in.dotm' -or $addin.Name -like 'MathType Commands*') { $addin.Installed = $false }
    }
    [void]$word.AddIns.Add([IO.Path]::GetFullPath($ExpectedTemplate), $true)
    foreach ($mode in @('direct', 'normal-style', 'custom-style', 'mixed', 'selection')) {
        $document = $word.Documents.Add()
        try {
            $document.Content.Text = "Unchanged before.`rInline \(C_d\), \(L_1\), \(M_1\).`r\[\frac{a}{b}\]`rUnchanged after."
            # Reproduce the user's 2 pt override over Word's 8 pt Normal style.
            $document.Styles.Item(-1).ParagraphFormat.SpaceAfter = 8
            $document.Content.ParagraphFormat.SpaceAfter = 2
            if ($mode -eq 'normal-style') {
                $document.Styles.Item(-1).ParagraphFormat.SpaceAfter = 2
                $document.Content.ParagraphFormat.Reset()
            }
            if ($mode -eq 'custom-style') {
                $style = $document.Styles.Add('AITeX Test Body', 1)
                $style.ParagraphFormat.SpaceAfter = 2
                $document.Content.Style = $style
                $document.Content.ParagraphFormat.Reset()
            }
            for ($i = 1; $i -le $document.Paragraphs.Count; $i++) {
                $format = $document.Paragraphs.Item($i).Format
                $format.SpaceBefore = $i
                $format.LeftIndent = $i * 3
                $format.RightIndent = $i * 2
                $format.FirstLineIndent = 6
                $format.LineSpacingRule = 4
                $format.LineSpacing = 18
                $format.KeepWithNext = ($i % 2 -eq 0)
                if ($mode -eq 'mixed') { $format.SpaceAfter = $i * 2 }
            }
            $before = @(Get-ParagraphSnapshot $document)
            $normalBefore = $document.Styles.Item(-1).ParagraphFormat.SpaceAfter
            $source = $document.Content.Text
            if ($mode -eq 'selection') {
                $selected = $document.Paragraphs.Item(2).Range
                $word.Selection.SetRange($selected.Start, $selected.End - 1)
                $word.Run('AITeX_ConvertSelectionMacro')
                if ($document.OMaths.Count -ne 3) { throw 'Selection conversion did not create three equations.' }
            } else {
                $converted = $word.Run('AITeX_ConvertDocumentSilent')
                if ([int]$converted -ne 4) { throw "Conversion failed: $($word.Run('AITeX_GetLastReport'))" }
            }
            Assert-ParagraphSnapshot $document $before
            if ($document.Styles.Item(-1).ParagraphFormat.SpaceAfter -ne $normalBefore) { throw 'The Normal style changed.' }
            [void]$document.Undo(1)
            if ($document.Content.Text -cne $source -or $document.OMaths.Count -ne 0) { throw 'Undo failed.' }
            Assert-ParagraphSnapshot $document $before
        } catch {
            $failures += "${mode}: $($_.Exception.Message)"
            [IO.File]::WriteAllText((Join-Path $PSScriptRoot "..\dist\probe-format-$mode.xml"), $document.Content.WordOpenXML)
            Write-Output ("${mode}: " + (ConvertTo-Json -InputObject @(Get-ParagraphSnapshot $document) -Depth 5 -Compress))
        }
        finally { $document.Close(0); $document = $null }
    }
    if ($failures.Count) { throw ($failures -join "`n") }
    Write-Output 'PASS: 5 formatting cases; spacing, styles, indents, line spacing, untouched paragraphs and Undo preserved.'
} finally {
    if ($document) { $document.Close(0) }
    if ($word) {
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
