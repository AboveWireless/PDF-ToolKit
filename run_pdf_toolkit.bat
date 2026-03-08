@echo off
setlocal EnableExtensions DisableDelayedExpansion

cd /d "%~dp0" || (
    echo Failed to switch to the PDF Toolkit directory.
    pause
    exit /b 1
)

set "APP_NAME=PDF Toolkit"
set "PYTHON_CMD="
set "EXTRA_PATHS="

call :resolve_python
if not defined PYTHON_CMD (
    echo.
    echo %APP_NAME% requires Python 3.11 or newer.
    echo Install Python 3.11+ or create a local .venv in this folder.
    pause
    exit /b 1
)

call :append_path_if_dir "%CD%\.venv\Scripts"
call :append_path_if_dir "C:\Program Files\Tesseract-OCR"
set "LAST_MATCH="
for /d %%D in ("%ProgramFiles%\gs\*\bin") do set "LAST_MATCH=%%~fD"
if defined LAST_MATCH call :append_path_if_dir "%LAST_MATCH%"
set "PROGRAMFILES_X86=%ProgramFiles(x86)%"
if defined PROGRAMFILES_X86 (
    set "LAST_MATCH="
    for /d %%D in ("%PROGRAMFILES_X86%\gs\*\bin") do set "LAST_MATCH=%%~fD"
    if defined LAST_MATCH call :append_path_if_dir "%LAST_MATCH%"
)
set "LAST_MATCH="
for /d %%D in ("%LOCALAPPDATA%\Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\*\Library\bin") do set "LAST_MATCH=%%~fD"
if defined LAST_MATCH call :append_path_if_dir "%LAST_MATCH%"

if defined EXTRA_PATHS (
    set "PATH=%EXTRA_PATHS%;%PATH%"
)

if defined PDF_TOOLKIT_VERBOSE (
    echo Launcher Python: %PYTHON_CMD%
    if defined EXTRA_PATHS echo Launcher PATH additions: %EXTRA_PATHS%
)

call :require_python_version
if errorlevel 1 exit /b 1

call :ensure_install
if errorlevel 1 exit /b 1

if defined PDF_TOOLKIT_DRY_RUN (
    echo %APP_NAME% launcher is configured.
    exit /b 0
)

call %PYTHON_CMD% -m pdf_toolkit
if errorlevel 1 (
    echo.
    echo %APP_NAME% failed to start.
    echo Set PDF_TOOLKIT_VERBOSE=1 before running for launcher diagnostics.
    pause
    exit /b 1
)

exit /b 0

:resolve_python
if defined PDF_TOOLKIT_PYTHON (
    if exist "%PDF_TOOLKIT_PYTHON%" (
        set "PYTHON_CMD="%PDF_TOOLKIT_PYTHON%""
        goto :eof
    )
)

if exist "%CD%\.venv\Scripts\python.exe" (
    set "PYTHON_CMD="%CD%\.venv\Scripts\python.exe""
    goto :eof
)

where py >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_CMD=py -3.11"
    goto :eof
)

where python >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_CMD=python"
)
goto :eof

:append_path_if_dir
if not "%~1"=="" (
    if exist "%~1\" (
        if defined EXTRA_PATHS (
            set "EXTRA_PATHS=%EXTRA_PATHS%;%~1"
        ) else (
            set "EXTRA_PATHS=%~1"
        )
    )
)
goto :eof

:require_python_version
call %PYTHON_CMD% -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 11) else 1)" >nul 2>nul
if errorlevel 1 (
    echo.
    echo %APP_NAME% requires Python 3.11 or newer.
    echo Launcher resolved: %PYTHON_CMD%
    pause
    exit /b 1
)
exit /b 0

:ensure_install
call %PYTHON_CMD% -c "import PySide6, pdf_toolkit" >nul 2>nul
if errorlevel 1 (
    if defined PDF_TOOLKIT_SKIP_INSTALL (
        echo.
        echo Required Python packages are missing and auto-install is disabled.
        echo Clear PDF_TOOLKIT_SKIP_INSTALL or run: %PYTHON_CMD% -m pip install -e .
        pause
        exit /b 1
    )

    echo Preparing %APP_NAME% environment...
    call %PYTHON_CMD% -m pip install --disable-pip-version-check --editable .
    if errorlevel 1 (
        echo.
        echo %APP_NAME% failed to install dependencies.
        echo Launcher resolved: %PYTHON_CMD%
        pause
        exit /b 1
    )
)
exit /b 0
