# Enable OCR Later

OCR is optional for the first public release. The rest of PDF Toolkit works without it.

## OCR dependencies

- `ocrmypdf`
- `tesseract`
- Ghostscript (`gswin64c.exe`)

## Options

- Install the tools system-wide and make sure they are on `PATH`.
- Bundle them under `vendor\` when building your own packaged app.

## Bundled layout

```text
vendor\
  bin\
    ocrmypdf.exe
    tesseract.exe
    gswin64c.exe
  tessdata\
    eng.traineddata
```

After bundling, rebuild the packaged app:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_gui.ps1
```
