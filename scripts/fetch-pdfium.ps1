param(
    [string]$Url = "https://github.com/bblanchon/pdfium-binaries/releases/latest/download/pdfium-win-x64.tgz",
    [string]$Destination = (Join-Path $PSScriptRoot "..\pdfium.dll")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$archiveRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("doco-pdfium-" + [System.Guid]::NewGuid().ToString("N"))
$archivePath = Join-Path $archiveRoot "pdfium-win-x64.tgz"

New-Item -ItemType Directory -Path $archiveRoot -Force | Out-Null

try {
    Write-Host "Downloading pdfium from $Url"
    Invoke-WebRequest -Uri $Url -OutFile $archivePath

    Write-Host "Extracting archive"
    tar -xf $archivePath -C $archiveRoot

    $dll = Get-ChildItem -Path $archiveRoot -Recurse -Filter "pdfium.dll" |
        Select-Object -First 1
    if (-not $dll) {
        throw "pdfium.dll not found after extraction."
    }

    $destinationPath = [System.IO.Path]::GetFullPath($Destination)
    $destinationDir = Split-Path -Parent $destinationPath
    if ($destinationDir) {
        New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
    }

    Copy-Item -Path $dll.FullName -Destination $destinationPath -Force
    Write-Host "pdfium.dll saved to $destinationPath"
}
finally {
    if (Test-Path $archiveRoot) {
        Remove-Item -Path $archiveRoot -Recurse -Force
    }
}
