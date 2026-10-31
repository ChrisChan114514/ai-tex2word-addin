[CmdletBinding()]
param(
    [string]$InstallDirectory = (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'),
    [string]$PandocPath,
    [switch]$BuildOnly,
    [switch]$SkipTest
)

$ErrorActionPreference = 'Stop'

if (-not $PandocPath) { $PandocPath = (Get-Command pandoc -ErrorAction Stop).Source }
$PandocPath = (Resolve-Path -LiteralPath $PandocPath).Path
$pandocVersion = & $PandocPath --version
if ($LASTEXITCODE -ne 0) { throw 'Pandoc could not start.' }
Write-Output $pandocVersion[0]

$projectRoot = Split-Path $PSScriptRoot -Parent
$sourceModule = Join-Path $projectRoot 'src\AITeXConverter.bas'
$ribbonSource = Join-Path $projectRoot 'src\customUI14.xml'
$distDirectory = Join-Path $projectRoot 'dist'
$templateName = 'AI TeX Word Add-in.dotm'
$distTemplate = Join-Path $distDirectory $templateName
$installedTemplate = Join-Path $InstallDirectory $templateName
$engineSource = Join-Path $projectRoot 'src\Convert-TexToOmml.ps1'
$engineDirectory = Join-Path $InstallDirectory 'AITeXEngine'

if (-not (Test-Path -LiteralPath $sourceModule)) {
    throw "Missing VBA source: $sourceModule"
}
if (-not (Test-Path -LiteralPath $ribbonSource)) {
    throw "Missing Ribbon source: $ribbonSource"
}
if (-not (Test-Path -LiteralPath $engineSource)) {
    throw "Missing TeX engine: $engineSource"
}

New-Item -ItemType Directory -Path $distDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $InstallDirectory -Force | Out-Null

$temporaryTemplate = Join-Path $distDirectory 'build.dotm'

if (Test-Path -LiteralPath $temporaryTemplate) {
    Remove-Item -LiteralPath $temporaryTemplate -Force
}

$builder = Join-Path $PSScriptRoot 'Build-VbaTemplate.ps1'
$powerShell = Join-Path $PSHOME 'powershell.exe'
& $powerShell -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass `
    -File $builder `
    -SourceModule $sourceModule `
    -OutputTemplate $temporaryTemplate

if ($LASTEXITCODE -ne 0) {
    throw "The isolated VBA build failed with exit code $LASTEXITCODE."
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [IO.Compression.ZipFile]::Open($temporaryTemplate, [IO.Compression.ZipArchiveMode]::Update)
try {
    $oldRibbon = $archive.GetEntry('customUI/customUI14.xml')
    if ($null -ne $oldRibbon) { $oldRibbon.Delete() }

    $ribbonEntry = $archive.CreateEntry('customUI/customUI14.xml')
    $ribbonWriter = New-Object IO.StreamWriter($ribbonEntry.Open(), (New-Object Text.UTF8Encoding($false)))
    try {
        $ribbonWriter.Write([IO.File]::ReadAllText($ribbonSource, [Text.Encoding]::UTF8))
    } finally {
        $ribbonWriter.Dispose()
    }

    $relsEntry = $archive.GetEntry('_rels/.rels')
    if ($null -eq $relsEntry) { throw 'The generated template has no root relationships file.' }

    $reader = New-Object IO.StreamReader($relsEntry.Open())
    try { $relsText = $reader.ReadToEnd() } finally { $reader.Dispose() }

    if ($relsText -notmatch 'customUI/customUI14\.xml') {
        $relationship = '<Relationship Id="rIdAITeXCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
        $relsText = $relsText.Replace('</Relationships>', $relationship + '</Relationships>')
    }

    $relsEntry.Delete()
    $newRelsEntry = $archive.CreateEntry('_rels/.rels')
    $relsWriter = New-Object IO.StreamWriter($newRelsEntry.Open(), (New-Object Text.UTF8Encoding($false)))
    try { $relsWriter.Write($relsText) } finally { $relsWriter.Dispose() }
} finally {
    $archive.Dispose()
}

Move-Item -LiteralPath $temporaryTemplate -Destination $distTemplate -Force

$distEngine = Join-Path $distDirectory 'AITeXEngine'
New-Item -ItemType Directory -Path $distEngine -Force | Out-Null
Copy-Item -LiteralPath $engineSource -Destination $distEngine -Force
@{ pandoc = $PandocPath } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $distEngine 'engine.json') -Encoding UTF8
Write-Output "Built:     $distTemplate"

if ($BuildOnly) {
    if (-not $SkipTest) {
        & (Join-Path $PSScriptRoot 'Test-AITeXWordAddin.ps1') -ExpectedTemplate $distTemplate
        & (Join-Path $PSScriptRoot 'Test-FormulaLayout.ps1') -ExpectedTemplate $distTemplate
        & (Join-Path $PSScriptRoot 'Test-FormulaFormatting.ps1') -ExpectedTemplate $distTemplate
    }
    return
}

if (Test-Path -LiteralPath $installedTemplate) {
    $backupName = $installedTemplate + '.backup-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
    Copy-Item -LiteralPath $installedTemplate -Destination $backupName -Force
}

try {
    Copy-Item -LiteralPath $distTemplate -Destination $installedTemplate -Force
} catch [IO.IOException] {
    throw "The build is ready in $distDirectory, but Word has locked the installed template. Save and close Word, then run this installer again. $($_.Exception.Message)"
}
Unblock-File -LiteralPath $installedTemplate -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $engineDirectory -Force | Out-Null
Copy-Item -LiteralPath $engineSource -Destination $engineDirectory -Force
@{ pandoc = $PandocPath } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $engineDirectory 'engine.json') -Encoding UTF8
Write-Output "Installed: $installedTemplate"

if (-not $SkipTest) {
    & (Join-Path $PSScriptRoot 'Test-AITeXWordAddin.ps1') -ExpectedTemplate $installedTemplate
    & (Join-Path $PSScriptRoot 'Test-FormulaLayout.ps1') -ExpectedTemplate $installedTemplate
    & (Join-Path $PSScriptRoot 'Test-FormulaFormatting.ps1') -ExpectedTemplate $installedTemplate
}

$openWord = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
if ($openWord.Count -gt 0) {
    Write-Warning 'Word is open. Restart it once to show the AI formula tab.'
}
