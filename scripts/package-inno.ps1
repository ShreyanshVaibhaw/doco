param(
    [string]$ScriptPath = (Join-Path $PSScriptRoot "..\installer\inno\doco.iss")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$innoScript = [System.IO.Path]::GetFullPath($ScriptPath)
$distDir = Join-Path $repoRoot "dist"

$iscc = (Get-Command "iscc.exe" -ErrorAction SilentlyContinue)?.Source
if (-not $iscc) {
    $default = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    if (Test-Path $default) {
        $iscc = $default
    }
}
if (-not $iscc) {
    throw "Inno Setup compiler not found. Install Inno Setup 6 and ensure ISCC.exe is available."
}

New-Item -ItemType Directory -Path $distDir -Force | Out-Null

Push-Location $repoRoot
try {
    Write-Host "Building Inno Setup installer via $iscc"
    & $iscc $innoScript

    $artifact = Join-Path $distDir "Doco-Setup.exe"
    if (-not (Test-Path $artifact)) {
        throw "Expected installer not produced: $artifact"
    }

    $sizeMb = [Math]::Round((Get-Item $artifact).Length / 1MB, 2)
    Write-Host ("Inno installer created: {0} ({1} MB)" -f $artifact, $sizeMb)
}
finally {
    Pop-Location
}
