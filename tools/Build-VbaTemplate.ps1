[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceModule,

    [Parameter(Mandatory = $true)]
    [string]$OutputTemplate
)

$ErrorActionPreference = 'Stop'
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security'
$word = $null
$document = $null

$hadSecurityKey = Test-Path -LiteralPath $securityKey
if (-not $hadSecurityKey) {
    New-Item -Path $securityKey -Force | Out-Null
}

$accessProperty = Get-ItemProperty -LiteralPath $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue
$hadAccessProperty = $null -ne $accessProperty
$oldAccessValue = if ($hadAccessProperty) { [int]$accessProperty.AccessVBOM } else { 0 }

try {
    Set-ItemProperty -LiteralPath $securityKey -Name AccessVBOM -Type DWord -Value 1

    $moduleText = [IO.File]::ReadAllText($SourceModule, [Text.Encoding]::UTF8)
    $moduleText = [regex]::Replace(
        $moduleText,
        '(?m)^Attribute VB_Name = "AITeXConverter"\r?\n',
        ''
    )

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    for ($index = 1; $index -le $word.AddIns.Count; $index++) {
        $addIn = $word.AddIns.Item($index)
        if ($addIn.Name -like 'MathType Commands*') {
            $addIn.Installed = $false
        }
    }

    $document = $word.Documents.Add()
    $component = $document.VBProject.VBComponents.Add(1)
    $component.Name = 'AITeXConverter'
    $component.CodeModule.AddFromString($moduleText)
    $document.SaveAs2($OutputTemplate, 15)
    $document.Close(0)
    $document = $null
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

    if ($hadAccessProperty) {
        Set-ItemProperty -LiteralPath $securityKey -Name AccessVBOM -Type DWord -Value $oldAccessValue
    } else {
        Remove-ItemProperty -LiteralPath $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue
    }
}

if (-not (Test-Path -LiteralPath $OutputTemplate)) {
    throw "Word did not create the template: $OutputTemplate"
}
