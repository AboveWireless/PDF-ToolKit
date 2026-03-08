# Run From Source

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install -e .[dev]
```

## Start the app

```powershell
.\run_pdf_toolkit.bat
```

Or:

```powershell
python -m pdf_toolkit
```

## Build the packaged app

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_gui.ps1
```
