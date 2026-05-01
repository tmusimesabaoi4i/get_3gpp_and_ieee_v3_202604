@echo off
setlocal
REM ============================================================
REM standarddocapp: Build distributable .exe for Windows
REM
REM Usage:
REM   1) (Optional) Activate venv (..\.venv\Scripts\activate)
REM   2) Run this script (double-click, cmd, or PowerShell)
REM
REM Output:
REM   dist\StandardDocApp.exe  (standalone, no Python required)
REM ============================================================

cd /d "%~dp0"

set "REPO_ROOT=%~dp0.."
set "ICON_PATH=%CD%\src\standarddocapp\assets\app.ico"
set "SPEC_PATH=%CD%\StandardDocApp.spec"
set "MAIN_PATH=%CD%\src\standarddocapp\__main__.py"

echo ============================================================
echo StandardDocApp build
echo Current directory: %CD%
echo Repo root:         %REPO_ROOT%
echo ============================================================

echo [0/5] Checking required files...

if not exist "%SPEC_PATH%" (
    echo ERROR: StandardDocApp.spec not found.
    echo Expected: %SPEC_PATH%
    pause
    exit /b 1
)

if not exist "%MAIN_PATH%" (
    echo ERROR: src\standarddocapp\__main__.py not found.
    echo Expected: %MAIN_PATH%
    pause
    exit /b 1
)

if not exist "%ICON_PATH%" (
    echo ERROR: app.ico not found.
    echo Expected: %ICON_PATH%
    echo Place a multi-size .ico ^(16/24/32/48/64/128/256^) at that path
    echo and re-run this script. See standarddocapp\src\standarddocapp\assets\README.md
    pause
    exit /b 1
)

echo OK: spec found.
echo OK: __main__.py found.
echo OK: icon found.

echo.
echo [1/5] Installing required packages...
python -m pip install --upgrade pip setuptools wheel
if errorlevel 1 (
    echo ERROR: pip upgrade failed.
    pause
    exit /b 1
)

python -m pip install -e "%REPO_ROOT%\stdharvest" -e "%REPO_ROOT%\stdsearch" -e . pyinstaller pillow
if errorlevel 1 (
    echo ERROR: Package install failed.
    pause
    exit /b 1
)

echo.
echo [2/5] Validating app.ico ^(sizes 16/24/32/48/64/128/256^)...
python "%REPO_ROOT%\tools\check_icon.py"
if errorlevel 1 (
    echo ERROR: Icon validation failed.
    pause
    exit /b 1
)

echo.
echo [3/5] Removing old build / dist...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

echo.
echo [4/5] Running PyInstaller...
python -m PyInstaller --noconfirm --clean StandardDocApp.spec
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo [5/5] Checking output exe...
if not exist "%CD%\dist\StandardDocApp.exe" (
    echo ERROR: dist\StandardDocApp.exe was not created.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo Build complete:
echo   %CD%\dist\StandardDocApp.exe
echo.
echo Please double-click the exe and confirm:
echo   - GUI launches
echo   - Window icon (title bar) is the new app.ico
echo   - Taskbar icon is the new app.ico
echo   - File icon in Explorer is the new app.ico
echo   - No black console window appears
echo.
echo If the Explorer icon still looks old, run:
echo   ie4uinit.exe -show
echo (Windows icon cache; it does NOT mean the embed failed.)
echo ============================================================
pause
endlocal
