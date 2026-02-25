param(
    [string]$SourcePath = (Join-Path $PSScriptRoot "..\installer\wix\Product.wxs")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$wxs = [System.IO.Path]::GetFullPath($SourcePath)
$distDir = Join-Path $repoRoot "dist"
$wixObjDir = Join-Path $distDir "wix"
$wixObj = Join-Path $wixObjDir "Product.wixobj"
$msiPath = Join-Path $distDir "Doco.msi"

$candle = (Get-Command "candle.exe" -ErrorAction SilentlyContinue)?.Source
$light = (Get-Command "light.exe" -ErrorAction SilentlyContinue)?.Source
if (-not $candle -or -not $light) {
    throw "WiX v3 tools not found. Install WiX Toolset and ensure candle.exe/light.exe are available."
}

New-Item -ItemType Directory -Path $wixObjDir -Force | Out-Null

Push-Location $repoRoot
try {
    Write-Host "Compiling WiX source"
    & $candle "-out" $wixObj $wxs

    Write-Host "Linking MSI"
    & $light "-ext" "WixUIExtension" "-out" $msiPath $wixObj

    if (-not (Test-Path $msiPath)) {
        throw "Expected MSI not produced: $msiPath"
    }

    $sizeMb = [Math]::Round((Get-Item $msiPath).Length / 1MB, 2)
    Write-Host ("MSI created: {0} ({1} MB)" -f $msiPath, $sizeMb)
}
finally {
    Pop-Location
}
