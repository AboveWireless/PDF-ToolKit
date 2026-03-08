$pythonExe = "python"
if (Test-Path ".\.venv\Scripts\python.exe") {
  $pythonExe = ".\.venv\Scripts\python.exe"
}

& $pythonExe -m PyInstaller `
  --noconfirm `
  --clean `
  pdf-toolkit-gui.spec
