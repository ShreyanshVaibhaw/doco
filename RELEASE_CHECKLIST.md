# Doco Release Checklist

## Phase 21.1 Deliverables

- [x] Release build profile configured in `Cargo.toml`
- [x] Release build script available (`scripts/build-release.ps1`)
- [x] Packaging scripts for Inno + WiX + checksum generation
- [x] Inno installer spec updated for shortcuts, file associations, context menu, optional PATH
- [x] WiX MSI spec updated for enterprise deployment path
- [x] Portable mode marker support (`doco.ini`) wired in settings path resolution
- [x] GitHub Actions release workflow added/updated
- [x] README sections expanded for features/download/build/contributing/license

## Final Validation Checklist

- [ ] All features working end-to-end
- [ ] No panics on any file input
- [ ] Memory usage within targets
- [ ] Startup time within targets
- [ ] All keyboard shortcuts working
- [ ] Themes all look correct
- [ ] Print working
- [ ] File associations working
- [ ] Installer tested on clean Windows 10 and 11
- [ ] README includes final runtime screenshots

## Current Evidence

- Automated checks executed in prompt flow:
- `cargo check` (twice)
- `cargo test` (twice)
- `cargo check --features pdf` (twice)
- Release build executed in prompt flow:
- `cargo build --release` (twice)
- `scripts/release.ps1 -SkipInno -SkipWix` (twice)
- Observed release binary size: `doco.exe` = `4.16 MB` (within `<5 MB` target)
- Local installer compilers are not installed on this machine (`ISCC`, `candle`, `light` missing), so installer binaries were not generated locally in this run.
