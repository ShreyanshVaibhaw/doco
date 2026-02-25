param(
    [switch]$FetchPdfium,
    [switch]$SkipInno,
    [switch]$SkipWix
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$distDir = Join-Path $repoRoot "dist"
$exePath = Join-Path $repoRoot "target\release\doco.exe"
$pdfiumPath = Join-Path $repoRoot "pdfium.dll"
$portableMarker = Join-Path $repoRoot "doco.ini"
$zipPath = Join-Path $distDir "doco-windows-x64.zip"
$checksumPath = Join-Path $distDir "checksums.txt"

Push-Location $repoRoot
try {
    New-Item -ItemType Directory -Path $distDir -Force | Out-Null

    & (Join-Path $PSScriptRoot "build-release.ps1") -FetchPdfium:$FetchPdfium

    if (-not (Test-Path $portableMarker)) {
        New-Item -ItemType File -Path $portableMarker -Force | Out-Null
    }

    $zipInputs = @($exePath, $portableMarker)
    if (Test-Path $pdfiumPath) {
        $zipInputs += $pdfiumPath
    }

    Compress-Archive -Path $zipInputs -DestinationPath $zipPath -Force

    if (-not $SkipInno) {
        & (Join-Path $PSScriptRoot "package-inno.ps1")
    }

    if (-not $SkipWix) {
        & (Join-Path $PSScriptRoot "package-wix.ps1")
    }

    $artifacts = Get-ChildItem -Path $distDir -File |
        Where-Object { $_.Extension -in @(".zip", ".exe", ".msi") } |
        Sort-Object Name

    $lines = @()
    foreach ($file in $artifacts) {
        $hash = (Get-FileHash -Algorithm SHA256 -Path $file.FullName).Hash.ToLowerInvariant()
        $lines += "$hash  $($file.Name)"
    }
    Set-Content -Path $checksumPath -Value $lines -NoNewline:$false -Encoding Ascii

    Write-Host "Release artifacts:"
    foreach ($file in $artifacts) {
        $sizeMb = [Math]::Round($file.Length / 1MB, 2)
        Write-Host ("- {0} ({1} MB)" -f $file.Name, $sizeMb)
    }
    Write-Host "Checksums written to $checksumPath"
}
finally {
    Pop-Location
}
