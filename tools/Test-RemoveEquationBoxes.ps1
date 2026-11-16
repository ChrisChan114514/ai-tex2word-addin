[CmdletBinding()]
param([string]$ExpectedTemplate = (Join-Path $PSScriptRoot '..\dist\AI TeX Word Add-in.dotm'))

$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
$mathNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
$wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

function New-MathRun([string]$text) { return "<m:r><m:t>$text</m:t></m:r>" }
function New-Box([string]$contents) { return "<m:borderBox><m:e>$contents</m:e></m:borderBox>" }

function Set-Fixture($doc, [string]$body) {
    [xml]$package = $doc.Content.WordOpenXML
    $ns = New-Object Xml.XmlNamespaceManager($package.NameTable)
    $ns.AddNamespace('w', $wordNamespace)
    $package.SelectSingleNode('//w:body', $ns).InnerXml =
        "<w:p xmlns:w='$wordNamespace' xmlns:m='$mathNamespace'>$body</w:p>"
    $doc.Content.InsertXML($package.OuterXml)
}

function Get-MathXml($range) {
    [xml]$xml = $range.WordOpenXML
    $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('m', $mathNamespace)
    # Exclude related header/footer parts from the selected story's snapshot.
    $ns.AddNamespace('pkg', 'http://schemas.microsoft.com/office/2006/xmlPackage')
    return ,@($xml.SelectNodes("/pkg:package/pkg:part[@pkg:name='/word/document.xml']//m:oMath", $ns))
}

function Get-Semantics($math) {
    # Compare the actual mathematical tree and run formatting, ignoring only
    # visible frame wrappers and Word's regenerated function control formatting.
    [xml]$xml = $math.OuterXml
    $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('m', $mathNamespace)
    foreach ($box in @($xml.SelectNodes('//m:borderBox', $ns))) {
        $hidden = $box.SelectNodes('m:borderBoxPr/m:hideTop | m:borderBoxPr/m:hideBot | m:borderBoxPr/m:hideLeft | m:borderBoxPr/m:hideRight', $ns)
        if ($hidden.Count -eq 4) { continue }
        $parent = $box.ParentNode
        foreach ($child in @($box.SelectSingleNode('m:e', $ns).ChildNodes)) {
            [void]$parent.InsertBefore($child, $box)
        }
        [void]$parent.RemoveChild($box)
    }
    foreach ($node in @($xml.SelectNodes('//m:ctrlPr', $ns))) { [void]$node.ParentNode.RemoveChild($node) }
    return $xml.DocumentElement.OuterXml
}

function Assert-Removed([int]$expected) {
    $removed = [int]$word.Run('AITeX_RemoveDocumentBoxesSilent')
    if ($removed -ne $expected) { throw "Removed $removed / $expected boxes. $($word.Run('AITeX_GetLastReport'))" }
    if ($word.Run('AITeX_GetLastReport')) { throw $word.Run('AITeX_GetLastReport') }
}

function Get-StoryBoxCount($doc) {
    $count = 0
    foreach ($firstStory in $doc.StoryRanges) {
        $story = $firstStory
        while ($null -ne $story) {
            foreach ($math in $story.OMaths) {
                foreach ($fn in $math.Functions) { if ($fn.Type -eq 4) { $count++ } }
            }
            $story = $story.NextStoryRange
        }
    }
    return $count
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    foreach ($addin in $word.AddIns) {
        if ($addin.Name -eq 'AI TeX Word Add-in.dotm' -or $addin.Name -like 'MathType Commands*') { $addin.Installed = $false }
    }
    [void]$word.AddIns.Add([IO.Path]::GetFullPath($ExpectedTemplate), $true)

    $document = $word.Documents.Add()
    $fraction = '<m:f><m:num>' + (New-Box (New-MathRun 'a')) + '</m:num><m:den>' + (New-MathRun 'b') + '</m:den></m:f>'
    $script = '<m:sSup><m:e>' + (New-Box (New-MathRun 'x')) + '</m:e><m:sup>' + (New-MathRun '2') + '</m:sup></m:sSup>'
    $cancel = '<m:borderBox><m:borderBoxPr><m:hideTop/><m:hideBot/><m:hideLeft/><m:hideRight/><m:strikeBLTR/></m:borderBoxPr><m:e>' + (New-MathRun 'z') + '</m:e></m:borderBox>'
    $invisible = '<m:box><m:e>' + (New-MathRun 'q') + '</m:e></m:box>'
    $body = '<w:r><w:t xml:space="preserve">Before </w:t></w:r><m:oMath>' +
        (New-Box (New-Box $fraction)) + (New-MathRun '+') + (New-Box $script) +
        (New-MathRun '+') + $cancel + $invisible + '</m:oMath><w:r><w:t xml:space="preserve"> after.</w:t></w:r>'
    Set-Fixture $document $body
    $document.Content.ParagraphFormat.SpaceAfter = 2
    $document.Content.ParagraphFormat.LeftIndent = 13
    $beforeXml = (Get-MathXml $document.Content)[0]
    $beforeSemantics = Get-Semantics $beforeXml
    $beforeText = $document.Content.Text
    $beforeType = $document.OMaths.Item(1).Type
    Assert-Removed 5
    $afterXml = (Get-MathXml $document.Content)[0]
    if ((Get-Semantics $afterXml) -cne $beforeSemantics) { throw 'Mathematical structure or formatting changed.' }
    if ($document.OMaths.Count -ne 1 -or $document.OMaths.Item(1).Type -ne $beforeType) { throw 'Equation count/type changed.' }
    if ($document.Content.ParagraphFormat.SpaceAfter -ne 2 -or $document.Content.ParagraphFormat.LeftIndent -ne 13) { throw 'Paragraph formatting changed.' }
    if (-not $document.Content.Text.StartsWith('Before ') -or -not $document.Content.Text.EndsWith(" after.`r")) { throw 'Surrounding prose changed.' }
    Assert-Removed 0
    [void]$document.Undo(1)
    if ($document.Content.Text -cne $beforeText -or (Get-MathXml $document.Content)[0].OuterXml -cne $beforeXml.OuterXml) { throw 'One Undo did not restore all nested boxes.' }
    $document.Close(0)
    $document = $null

    # Exercise every story and a second section's unlinked header/footer.
    $document = $word.Documents.Add()
    Set-Fixture $document ('<w:r><w:t>Text</w:t></w:r><m:oMath>' + (New-Box (New-MathRun 'v')) + '</m:oMath><w:r><w:t>End</w:t></w:r>')
    $source = $document.OMaths.Item(1).Range
    $targets = @($document.Sections.Item(1).Headers.Item(1).Range, $document.Sections.Item(1).Footers.Item(1).Range)
    $anchor = $document.Range($document.Content.End - 1, $document.Content.End - 1)
    $targets += $document.Footnotes.Add($anchor).Range
    $targets += $document.Endnotes.Add($anchor).Range
    $targets += $document.Comments.Add($document.Range(0, 1), 'note').Range
    $shape = $document.Shapes.AddTextbox(1, 20, 20, 200, 50)
    $targets += $shape.TextFrame.TextRange
    $shape = $document.Shapes.AddTextbox(1, 20, 100, 200, 50)
    $targets += $shape.TextFrame.TextRange
    [void]$document.Shapes.AddShape(1, 250, 20, 50, 50)
    $table = $document.Tables.Add($document.Range($document.Content.End - 1, $document.Content.End - 1), 1, 1)
    $cell = $table.Cell(1, 1).Range
    $cell.End = $cell.End - 1
    $targets += $cell
    $section = $document.Sections.Add()
    $section.Headers.Item(1).LinkToPrevious = $false
    $section.Footers.Item(1).LinkToPrevious = $false
    $targets += $section.Headers.Item(1).Range
    $targets += $section.Footers.Item(1).Range
    foreach ($target in $targets) { $target.FormattedText = $source.FormattedText }
    $headerShapes = @($section.Headers.Item(1).Shapes.AddTextbox(1, 20, 20, 200, 50),
        $section.Footers.Item(1).Shapes.AddTextbox(1, 20, 20, 200, 50))
    foreach ($shape in $headerShapes) { $shape.TextFrame.TextRange.FormattedText = $source.FormattedText }
    if ((Get-StoryBoxCount $document) -ne (1 + $targets.Count)) { throw 'Story fixture has an unexpected box count.' }
    Assert-Removed (1 + $targets.Count + $headerShapes.Count)
    foreach ($range in @($source) + $targets + @($headerShapes | ForEach-Object { $_.TextFrame.TextRange })) {
        if ($range.OMaths.Count -ne 1 -or $range.WordOpenXML -match '<m:borderBox[ >]') { throw 'A story lost its equation or kept its box.' }
    }
    [void]$document.Undo(1)
    # Reacquire ranges: previously held COM ranges need not expand after Undo.
    if ((Get-StoryBoxCount $document) -ne (1 + $targets.Count)) { throw 'Undo did not restore all story boxes.' }
    foreach ($shape in $headerShapes) {
        if ($shape.TextFrame.TextRange.WordOpenXML -notmatch '<m:borderBox[ >]') { throw 'Undo did not restore a header/footer textbox box.' }
    }
    $document.Close(0)
    $document = $null

    $document = $word.Documents.Add()
    $document.Content.Text = 'Inline \(\boxed{E=mc^2}\), plain \(x_i\). Display \[\boxed{\frac{a}{b}}\]'
    if ([int]$word.Run('AITeX_ConvertDocumentSilent') -ne 3) { throw 'Conversion integration failed.' }
    $types = @(foreach ($math in $document.OMaths) { [int]$math.Type })
    Assert-Removed 2
    if ($document.Content.WordOpenXML -match '<m:borderBox[ >]' -or $document.OMaths.Count -ne 3) { throw 'Converted boxed formulas were not unwrapped.' }
    if ((@(foreach ($math in $document.OMaths) { [int]$math.Type }) -join ',') -ne ($types -join ',')) { throw 'Inline/display types changed.' }
    $document.Protect(3)
    if ([int]$word.Run('AITeX_RemoveDocumentBoxesSilent') -ne 0 -or -not $word.Run('AITeX_GetLastReport')) { throw 'Protected document handling failed.' }
    $document.Unprotect()
    $document.Content.Text = 'No equations; leave this text alone.'
    $sourceText = $document.Content.Text
    Assert-Removed 0
    if ($document.Content.Text -cne $sourceText) { throw 'A document without equations changed.' }
    Write-Output 'PASS: nested/sibling boxes, fraction/script contents, cancellation/grouping, formatting, all stories, conversion integration, no-op, protection and single Undo.'
} finally {
    if ($document) { $document.Close(0) }
    if ($word) {
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
