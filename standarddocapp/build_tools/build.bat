@echo off
REM Wrapper that runs build.ps1 without being blocked by the
REM Windows PowerShell execution policy. Double-click friendly.
REM
REM Usage:
REM   build.bat                ... full build (creates .venv-build)
REM   build.bat -SkipDeps      ... skip pip installs
REM   build.bat -Clean         ... wipe dist/ and build/ first

setlocal
set "SCRIPT_DIR=%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%build.ps1" %*
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo Build failed with exit code %RC%.
)
exit /b %RC%
