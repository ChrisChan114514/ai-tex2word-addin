[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$RequestPath,
    [Parameter(Mandatory = $true)][string]$ResponsePath,
    [string]$PandocPath
)

$ErrorActionPreference = 'Stop'
$utf8 = New-Object Text.UTF8Encoding($false)
$response = New-Object Xml.XmlDocument
$root = $response.AppendChild($response.CreateElement('results'))
$scratch = $null

function Add-ErrorResult($id, [string]$message) {
    $item = $root.AppendChild($response.CreateElement('result'))
    $item.SetAttribute('id', [string]$id)
    $item.SetAttribute('error', $message)
}

function Repair-StretchyGroups($document, $math, $ns) {
    # texmath can encode a lower brace/bracket as a limit containing one glyph.
    # In Word that glyph stays tiny. Lower/upper grouping marks must use the
    # semantic groupChr structure so they stretch across the complete operand.
    # This adapts OMML structures; it never rewrites or reparses TeX commands.
    $mathUri = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    $nodes = $math.SelectNodes('.//m:limLow | .//m:limUpp', $ns)
    for ($index = $nodes.Count - 1; $index -ge 0; $index--) {
        $node = $nodes.Item($index)
        $limit = $node.SelectSingleNode('m:lim', $ns)
        if ($null -eq $limit -or $limit.ChildNodes.Count -ne 1) { continue }
        $symbol = $limit.SelectSingleNode('m:r/m:t', $ns)
        if ($null -eq $symbol -or $symbol.InnerText.Length -ne 1) { continue }
        $code = [int][char]$symbol.InnerText
        $bottom = $node.LocalName -eq 'limLow'
        if (($bottom -and $code -notin @(0x23DF, 0x23B5)) -or
            (-not $bottom -and $code -notin @(0x23DE, 0x23B4))) { continue }
        $group = $document.CreateElement('m', 'groupChr', $mathUri)
        $properties = $group.AppendChild($document.CreateElement('m', 'groupChrPr', $mathUri))
        $values = [ordered]@{
            chr = $symbol.InnerText
            pos = $(if ($bottom) { 'bot' } else { 'top' })
            vertJc = $(if ($bottom) { 'top' } else { 'bot' })
        }
        foreach ($name in $values.Keys) {
            $property = $properties.AppendChild($document.CreateElement('m', $name, $mathUri))
            [void]$property.SetAttribute('val', $mathUri, $values[$name])
        }
        [void]$group.AppendChild($node.SelectSingleNode('m:e', $ns).CloneNode($true))
        [void]$node.ParentNode.ReplaceChild($group, $node)
    }
}

try {
    if (-not $PandocPath) {
        $config = Join-Path $PSScriptRoot 'engine.json'
        if (Test-Path -LiteralPath $config) {
            $PandocPath = (Get-Content -LiteralPath $config -Raw -Encoding UTF8 | ConvertFrom-Json).pandoc
        } else {
            $PandocPath = (Get-Command pandoc -ErrorAction Stop).Source
        }
    }
    if (-not (Test-Path -LiteralPath $PandocPath -PathType Leaf)) {
        throw 'Pandoc was not found. Reinstall the add-in with -PandocPath.'
    }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $request = New-Object Xml.XmlDocument
    $request.XmlResolver = $null
    $request.Load($RequestPath)
    $scratch = Join-Path ([IO.Path]::GetTempPath()) ('AITeX-engine-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($scratch)
    $inputFile = Join-Path $scratch 'input.json'
    $outputFile = Join-Path $scratch 'output.docx'
    $cache = @{}
    foreach ($formula in $request.SelectNodes('/request/formula')) {
        try {
            $tex = $formula.InnerText.Trim()
            if (-not $tex) { throw 'Empty formula.' }
            # Preserve TeX syntax. Only repair the existing AI/Markdown convention.
            $tex = $tex.Replace('\_', '_').Replace([char]13, [char]10).Replace([char]11, [char]10)
            $key = $formula.GetAttribute('display') + ':' + $tex
            if (-not $cache.ContainsKey($key)) {
                $kind = if ($formula.GetAttribute('display') -eq '1') { 'DisplayMath' } else { 'InlineMath' }
                # A Math AST avoids interpreting formula content as Markdown, HTML,
                # document-level LaTeX, filters, includes, or shell commands.
                $ast = @{
                    'pandoc-api-version' = @(1, 23, 1)
                    meta = @{}
                    blocks = @(@{ t = 'Para'; c = @(@{ t = 'Math'; c = @(@{t = $kind}, $tex) }) })
                }
                [IO.File]::WriteAllText($inputFile, (ConvertTo-Json -InputObject $ast -Depth 20 -Compress), $utf8)
                if (Test-Path -LiteralPath $outputFile) { Remove-Item -LiteralPath $outputFile }
                $info = New-Object Diagnostics.ProcessStartInfo
                $info.FileName = $PandocPath
                $info.Arguments = '--from=json --to=docx --sandbox --fail-if-warnings --output="' + $outputFile + '" "' + $inputFile + '"'
                $info.UseShellExecute = $false
                $info.CreateNoWindow = $true
                $info.RedirectStandardError = $true
                $info.StandardErrorEncoding = $utf8
                $process = New-Object Diagnostics.Process
                $process.StartInfo = $info
                try {
                    [void]$process.Start()
                    $errorTask = $process.StandardError.ReadToEndAsync()
                    if (-not $process.WaitForExit(15000)) {
                        $process.Kill()
                        throw 'The formula exceeded the 15 second parsing limit.'
                    }
                    $details = $errorTask.GetAwaiter().GetResult()
                    if ($process.ExitCode -ne 0 -or $details.Trim()) {
                        throw ('TeX parser: ' + $details.Trim())
                    }
                } finally { $process.Dispose() }
                $archive = [IO.Compression.ZipFile]::OpenRead($outputFile)
                try {
                    $reader = New-Object IO.StreamReader($archive.GetEntry('word/document.xml').Open())
                    try { [xml]$document = $reader.ReadToEnd() } finally { $reader.Dispose() }
                } finally { $archive.Dispose() }
                $ns = New-Object Xml.XmlNamespaceManager($document.NameTable)
                $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
                $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                $math = $document.SelectNodes('//m:oMath', $ns)
                if ($math.Count -ne 1 -or $document.SelectNodes('//w:body//w:t', $ns).Count -ne 0) {
                    throw 'The parser did not produce exactly one native equation.'
                }
                Repair-StretchyGroups $document $math.Item(0) $ns
                $cache[$key] = $math.Item(0).OuterXml
            }
            $item = $root.AppendChild($response.CreateElement('result'))
            $item.SetAttribute('id', $formula.GetAttribute('id'))
            $fragment = New-Object Xml.XmlDocument
            $fragment.LoadXml($cache[$key])
            [void]$item.AppendChild($response.ImportNode($fragment.DocumentElement, $true))
        } catch {
            Add-ErrorResult $formula.GetAttribute('id') $_.Exception.Message
        }
    }
} catch {
    $root.SetAttribute('error', $_.Exception.Message)
} finally {
    $response.Save($ResponsePath)
    if ($scratch -and (Test-Path -LiteralPath $scratch)) {
        $resolvedScratch = [IO.Path]::GetFullPath($scratch)
        $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
        if ([IO.Path]::GetDirectoryName($resolvedScratch) -eq $tempRoot -and
            [IO.Path]::GetFileName($resolvedScratch) -like 'AITeX-engine-*') {
            Remove-Item -LiteralPath $resolvedScratch -Recurse -Force
        }
    }
}
