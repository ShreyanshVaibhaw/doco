param(
    [switch]$FetchPdfium,
    [switch]$FailOnSizeLimit
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$exePath = Join-Path $repoRoot "target\release\doco.exe"
$pdfiumPath = Join-Path $repoRoot "pdfium.dll"
$sizeTargetBytes = 5MB

Push-Location $repoRoot
try {
    if ($FetchPdfium -and -not (Test-Path $pdfiumPath)) {
        & (Join-Path $PSScriptRoot "fetch-pdfium.ps1") -Destination $pdfiumPath
    }

    Write-Host "Running cargo build --release"
    cargo build --release

    if (-not (Test-Path $exePath)) {
        throw "Release executable not found at $exePath"
    }

    $exe = Get-Item $exePath
    $exeSizeMb = [Math]::Round($exe.Length / 1MB, 2)
    Write-Host ("doco.exe size: {0} MB" -f $exeSizeMb)

    if ($exe.Length -gt $sizeTargetBytes) {
        $message = "doco.exe exceeds 5 MB target (actual: $exeSizeMb MB)."
        if ($FailOnSizeLimit) {
            throw $message
        }
        Write-Warning $message
    }

    if (-not (Test-Path $pdfiumPath)) {
        Write-Warning "pdfium.dll is missing in repository root. Installers can build, but PDF support will be incomplete."
    }
    else {
        $pdfium = Get-Item $pdfiumPath
        Write-Host ("pdfium.dll size: {0} MB" -f ([Math]::Round($pdfium.Length / 1MB, 2)))
    }
}
finally {
    Pop-Location
}
