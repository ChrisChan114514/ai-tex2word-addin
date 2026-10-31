[CmdletBinding()]
param(
    [string]$InstallDirectory = (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP')
)

$ErrorActionPreference = 'Stop'
$installedTemplate = Join-Path $InstallDirectory 'AI TeX Word Add-in.dotm'

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw 'Word is running. Save your documents, close Word, and run the uninstaller again.'
}

if (Test-Path -LiteralPath $installedTemplate) {
    Remove-Item -LiteralPath $installedTemplate -Force
    Write-Output "Removed: $installedTemplate"
} else {
    Write-Output 'The AI TeX Word Add-in is not installed.'
}

# Remove only companion files owned by this add-in, leaving other files alone.
$engineDirectory = Join-Path $InstallDirectory 'AITeXEngine'
foreach ($name in @('Convert-TexToOmml.ps1', 'engine.json')) {
    $path = Join-Path $engineDirectory $name
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
}
if ((Test-Path -LiteralPath $engineDirectory) -and
    @(Get-ChildItem -LiteralPath $engineDirectory -Force).Count -eq 0) {
    Remove-Item -LiteralPath $engineDirectory
}
