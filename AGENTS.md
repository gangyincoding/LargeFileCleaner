# Repository Guidelines

## Project Structure & Module Organization

- Core modules (root): `disk_analyzer_gui_stable.py` (Tkinter GUI, entry point),
  `disk_scanner_simple.py` (scan engine), `duplicate_finder.py` (duplicate logic),
  `file_safety.py` (safety checks), exporters `export_excel.py` and `export_csv.py`.
- Packaging: `DiskCleaner_v1.4.spec` (PyInstaller). Build artifacts in `build/` and `dist/`.
- Docs: `docs/` (development reports), user guides at repo root, plus `README.md` and `CHANGELOG.md`.
- Misc: `DiskCleaner_GUI_Stable.bat` (Windows launcher), `index.html` (GitHub Pages project page).

## Build, Test, and Development Commands

- Setup (Windows, PowerShell):
  - `python -m venv .venv` && `.venv\Scripts\activate`
  - `pip install -r requirements.txt`
- Run GUI (dev): `python disk_analyzer_gui_stable.py` or `./DiskCleaner_GUI_Stable.bat`
- Package (PyInstaller): `pyinstaller DiskCleaner_v1.4.spec` (output in `dist/`)
- Optional tools: `pytest -q` (if tests exist), `pylint disk_* export_* file_safety.py`

## Coding Style & Naming Conventions

- Python 3.6+; follow PEP 8; 4‑space indentation; UTF‑8 source and file I/O.
- snake_case for functions/variables, PascalCase for classes; module names in snake_case.
- Prefer `pathlib.Path` for filesystem paths; keep Windows compatibility in mind.
- Keep functions small and single‑purpose; avoid introducing mandatory new deps.

## Testing Guidelines

- Use `pytest`; place tests under `tests/` and name `test_*.py`.
- Prioritize pure logic: hashing, grouping, size formatting, safety checks.
- For GUI, add smoke tests or manual steps; verify scanning a small sample and all export formats.
- If you add tests, ensure they pass on Python 3.6+.

## Commit & Pull Request Guidelines

- Commits: use conventional prefixes — `feat:`, `fix:`, `docs:`, `chore:`, `refactor:`, `test:`.
  Example: `feat(duplicate): add MD5 two‑phase check`.
- PRs include: summary/purpose, linked issues, validation steps, screenshots/GIFs for UI changes,
  and notes on safety defaults. Update `CHANGELOG.md` for user‑visible changes.

## Security & Configuration Tips

- Default deletions to `send2trash` when available; never hard‑code system paths.
- Respect `file_safety.py` protections; avoid scanning `C:\` root by default.
- Log destructive operations to `delete_log.txt`; require explicit user confirmation in UI flows.

