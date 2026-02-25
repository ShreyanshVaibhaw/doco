# Doco

Doco is a native Windows document editor and viewer written in Rust with Win32 + Direct2D rendering.

## Screenshots

Runtime screenshots should live in `screenshots/` and be embedded here:

- Main editor shell: `screenshots/main-window.png`
- Command palette: `screenshots/command-palette.png`
- Settings panel: `screenshots/settings-panel.png`

## Features

- Native Windows shell with tabs, sidebar, toolbar, command palette, status bar, and toast notifications
- Document support for `.docx`, `.pdf` (feature-gated), `.txt`, and `.md`
- Editing tools: formatting, find/replace, table editing, image insert/resize, clipboard support
- Export paths: DOCX, PDF, TXT, Markdown, and HTML
- Portable mode support via `doco.ini`
- Windows integration for drag-drop, print dialog integration, file associations, and Explorer context menu scaffolding

## Downloads

- Latest releases: `https://github.com/ShreyanshVaibhaw/doco/releases/latest`
- Release assets include:
- `doco-windows-x64.zip` (portable package)
- `Doco-Setup.exe` (Inno consumer installer)
- `Doco.msi` (WiX MSI for enterprise deployment)
- `checksums.txt` (SHA-256 hashes)

## Build From Source

Requirements:

- Windows 10 or Windows 11
- Rust stable toolchain
- Visual Studio Build Tools (MSVC target)
- Optional: `pdfium.dll` in repo root for PDF runtime support

Development:

```powershell
cargo check
cargo test
cargo run
```

Release:

```powershell
.\scripts\build-release.ps1 -FetchPdfium
```

Create all distribution artifacts (zip + installers + checksums):

```powershell
.\scripts\release.ps1 -FetchPdfium
```

## Packaging

- Inno Setup script: `installer/inno/doco.iss`
- WiX MSI script: `installer/wix/Product.wxs`
- Release automation: `.github/workflows/release.yml`
- Helper scripts:
- `scripts/fetch-pdfium.ps1`
- `scripts/build-release.ps1`
- `scripts/package-inno.ps1`
- `scripts/package-wix.ps1`
- `scripts/release.ps1`

## Portable Mode

If `doco.ini` exists next to `doco.exe`, Doco runs in portable mode:

- Settings are read/written next to the executable (`settings.json`)
- Recovery files are written next to the executable (`recovery/`)
- Theme overrides are read from local portable theme paths
- No `%APPDATA%` requirement for settings/theme/recovery paths

## Contributing

Contribution process is documented in `CONTRIBUTING.md`.

## License

Planned dual-license model:

- MIT
- Apache-2.0

