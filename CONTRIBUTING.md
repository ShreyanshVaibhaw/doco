# Contributing to Doco

## Setup

1. Install Rust stable (`rustup default stable`).
2. Install Visual Studio Build Tools with MSVC support.
3. Clone the repository and run:

```powershell
cargo check
cargo test
```

## Development Workflow

1. Create a feature branch from `main`.
2. Keep changes scoped to one prompt/phase when possible.
3. Run validation before opening a pull request:

```powershell
cargo check
cargo test
cargo check --features pdf
```

## Pull Requests

1. Add a short summary of behavior changes.
2. Include manual test notes for UI and Windows integration changes.
3. Update `RELEASE_CHECKLIST.md` when release criteria are affected.

## Code Style

- Prefer small, composable functions.
- Keep platform-specific behavior behind modules in `src/window/` and `src/settings/`.
- Add focused tests for logic-only changes.
