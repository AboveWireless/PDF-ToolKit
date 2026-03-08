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

## Pull request expectations

- Keep changes focused and easy to review.
- Add or update tests when behavior changes.
- Preserve the Windows-first packaging flow.
- Document user-facing changes in `README.md` or `docs/` when needed.

## Packaging notes

- The primary release target is a portable Windows build.
- OCR is optional for the first public release and should not break core workflows when unavailable.
- If you add bundled tools later, place them under `vendor\`.
