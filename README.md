# Doco

Doco is a native Windows document editor and viewer built in Rust with Win32 + Direct2D.

## Features

- Modern Win11-style shell: tabs, sidebar, command palette, status bar
- Multi-format support: `.docx`, `.pdf` (feature-gated), `.txt`, `.md`
- Editing systems: find/replace, formatting model, tables, image blocks
- Export targets: DOCX, PDF, TXT, Markdown, HTML
- Settings system with schema migration and debounced auto-save
- Windows integration scaffolding: drag/drop, jump list updates, print hooks

## Build From Source

Requirements:

- Windows 10/11
- Rust stable toolchain
- Visual Studio Build Tools (MSVC)

Commands:

```powershell
cargo check
cargo run
```

Release build:

```powershell
cargo build --release
```

## Packaging

- Inno Setup script: `installer/inno/doco.iss`
- WiX template: `installer/wix/Product.wxs`
- CI workflow: `.github/workflows/release.yml`

## Portable Mode

If `doco.ini` exists next to `doco.exe`, Doco runs in portable mode:

- Settings stored next to the executable (`settings.json`)
- Recovery data stored next to the executable (`recovery/`)
- No mandatory `%APPDATA%` usage

## Contributing

1. Fork the repository.
2. Create a feature branch.
3. Run `cargo check` before opening a PR.

## License

Suggested dual license:

- MIT
- Apache-2.0

