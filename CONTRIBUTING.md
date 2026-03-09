# Contributing

Thanks for contributing to PDF Toolkit.

## Local setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install -e .[dev]
```

## Run the app

```powershell
.\run_pdf_toolkit.bat
```

Or run directly:

```powershell
python -m pdf_toolkit
```

## Test before opening a PR

```powershell
python -m pytest
powershell -ExecutionPolicy Bypass -File scripts\build_gui.ps1
```

If you touch installer or release flow code, also run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\package_release.ps1
```

## Pull request expectations

- Keep changes focused and easy to review.
- Add or update tests when behavior changes.
- Preserve the Windows-first packaging flow.
- Document user-facing changes in `README.md` or `docs/` when needed.
- Update `CHANGELOG.md` for notable user-facing changes.

## Packaging notes

- The primary public release target is the Windows installer.
- The portable ZIP remains a first-class fallback artifact.
- OCR is optional for the first public release and should not break core workflows when unavailable.
- If you add bundled tools later, place them under `vendor\`.
