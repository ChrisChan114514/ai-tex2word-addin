[CmdletBinding()]
param([string]$EnginePath = (Join-Path $PSScriptRoot '..\src\Convert-TexToOmml.ps1'))
$ErrorActionPreference = 'Stop'
$cases = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\tests\formulas.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$scratch = Join-Path ([IO.Path]::GetTempPath()) ('AITeX-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($scratch)
try {
    $request = New-Object Xml.XmlDocument
    [void]$request.LoadXml('<request/>')
    foreach ($case in $cases) {
        $formula = $request.DocumentElement.AppendChild($request.CreateElement('formula'))
        $formula.SetAttribute('id', $case.id)
        $formula.SetAttribute('display', '1')
        $formula.InnerText = $case.tex
    }
    $inputPath = Join-Path $scratch 'request.xml'
    $outputPath = Join-Path $scratch 'response.xml'
    $request.Save($inputPath)
    & $EnginePath -RequestPath $inputPath -ResponsePath $outputPath
    [xml]$response = Get-Content -LiteralPath $outputPath -Raw -Encoding UTF8
    if ($response.results.HasAttribute('error')) { throw $response.results.error }
    $ns = New-Object Xml.XmlNamespaceManager($response.NameTable)
    $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
    $failures = @()
    foreach ($case in $cases) {
        try {
            $result = $response.SelectSingleNode("/results/result[@id='$($case.id)']")
            if ($case.error) {
                if (-not $result.HasAttribute('error')) { throw 'Expected rejection.' }
            } else {
                if ($result.HasAttribute('error')) { throw $result.error }
                $math = $result.SelectSingleNode('m:oMath', $ns)
                if ($null -eq $math) { throw 'No native OMML.' }
                if ($case.xpath -and $math.SelectNodes($case.xpath, $ns).Count -ne $case.count) {
                    throw "Expected $($case.count) nodes for $($case.xpath); got $($math.SelectNodes($case.xpath, $ns).Count)."
                }
                foreach ($check in $case.checks) {
                    if ($math.SelectNodes($check.xpath, $ns).Count -ne $check.count) {
                        throw "Expected $($check.count) nodes for $($check.xpath)."
                    }
                }
                $text = ($math.SelectNodes('.//m:t', $ns) | ForEach-Object { $_.InnerText }) -join ''
                if ($case.text) {
                    $pattern = ($case.text.ToCharArray() | ForEach-Object { [regex]::Escape([string]$_) }) -join '.*'
                    if ($text -cnotmatch $pattern) { throw "Text missing or reordered: $($case.text)" }
                }
            }
        } catch { $failures += "$($case.id): $($_.Exception.Message)" }
    }
    if ($failures.Count) { throw ($failures -join "`n") }
    Write-Output "PASS: $($cases.Count) TeX engine fixtures (native structures, Unicode, nesting, macros and rejection)."
} finally {
    $resolvedScratch = [IO.Path]::GetFullPath($scratch)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
    if ([IO.Path]::GetDirectoryName($resolvedScratch) -eq $tempRoot -and
        [IO.Path]::GetFileName($resolvedScratch) -like 'AITeX-test-*') {
        Remove-Item -LiteralPath $resolvedScratch -Recurse -Force
    }
}
