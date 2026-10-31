[CmdletBinding()]
param(
    [string]$ExpectedTemplate = (Join-Path $PSScriptRoot '..\dist\AI TeX Word Add-in.dotm'),
    [string]$ReviewDirectory
)
$ErrorActionPreference = 'Stop'
$cases = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\tests\layout.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$word = $null
$document = $null
$failures = @()
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    for ($index = 1; $index -le $word.AddIns.Count; $index++) {
        $addin = $word.AddIns.Item($index)
        if ($addin.Name -eq 'AI TeX Word Add-in.dotm' -or $addin.Name -like 'MathType Commands*') {
            $addin.Installed = $false
        }
    }
    [void]$word.AddIns.Add([IO.Path]::GetFullPath($ExpectedTemplate), $true)
    foreach ($case in $cases) {
        $document = $word.Documents.Add()
        try {
            $document.Content.Text = $case.source
            $source = $document.Content.Text
            $converted = $word.Run('AITeX_ConvertDocumentSilent')
            $expectedConverted = if ($case.convertedCount) { $case.convertedCount } else { $case.types.Count }
            if ([int]$converted -ne $expectedConverted -or $document.OMaths.Count -ne $case.types.Count) {
                throw "Equation count: $converted. $($word.Run('AITeX_GetLastReport'))"
            }
            [xml]$native = $document.Content.WordOpenXML
            $ns = New-Object Xml.XmlNamespaceManager($native.NameTable)
            $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
            $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            $layout = ''
            foreach ($paragraph in $native.SelectNodes('//w:body/w:p', $ns)) {
                foreach ($node in $paragraph.SelectNodes('.//w:t[not(ancestor::m:oMath)] | .//m:oMath | .//w:br', $ns)) {
                    switch ($node.LocalName) {
                        't' { $layout += $node.InnerText }
                        'oMath' { $layout += '[M]' }
                        'br' { $layout += [char]11 }
                    }
                }
                $layout += "`r"
            }
            if ($layout -cne $case.layout) {
                throw ('Paragraph/break mismatch. Expected ' + (ConvertTo-Json $case.layout -Compress) + '; got ' + (ConvertTo-Json $layout -Compress))
            }
            for ($index = 0; $index -lt $case.types.Count; $index++) {
                if ([int]$document.OMaths.Item($index + 1).Type -ne $case.types[$index]) {
                    throw "Equation $($index + 1) has the wrong inline/display type."
                }
            }
            if ($case.mathText) {
                $mathText = ($native.SelectNodes('//m:oMath//m:t', $ns) | ForEach-Object { $_.InnerText }) -join ''
                if ($mathText -cne $case.mathText) { throw 'Adjacent inline formulas lost their contents.' }
            }
            [void]$document.Undo(1)
            if ($document.OMaths.Count -ne 0 -or $document.Content.Text -cne $source) {
                throw 'One Undo did not restore the original source and breaks.'
            }
            if ($ReviewDirectory -and $case.id -eq 'reported-chinese-prose') {
                [void]$document.Redo(1)
                New-Item -ItemType Directory -Path $ReviewDirectory -Force | Out-Null
                [string]$reviewDocx = Join-Path $ReviewDirectory 'inline-layout.docx'
                [string]$reviewPdf = Join-Path $ReviewDirectory 'inline-layout.pdf'
                $document.SaveAs2($reviewDocx, 16)
                $document.ExportAsFixedFormat($reviewPdf, 17)
            }
        } catch { $failures += "$($case.id): $($_.Exception.Message)" }
        finally { $document.Close(0); $document = $null }
    }
    if ($failures.Count) { throw ($failures -join "`n") }
    Write-Output "PASS: $($cases.Count) Word layout cases; inline prose, display blocks, original breaks and Undo verified."
} finally {
    if ($document) { $document.Close(0) }
    if ($word) {
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
