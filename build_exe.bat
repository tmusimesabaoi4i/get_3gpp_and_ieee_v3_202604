@echo off
REM ============================================================
REM Repo-root convenience wrapper.
REM Forwards to standarddocapp\build_exe.bat which performs the
REM real PyInstaller build and writes
REM   standarddocapp\dist\StandardDocApp.exe
REM ============================================================

setlocal

if not exist "%~dp0standarddocapp\build_exe.bat" (
    echo ERROR: standarddocapp\build_exe.bat not found.
    echo Are you running this from the repo root?
    pause
    exit /b 1
)

call "%~dp0standarddocapp\build_exe.bat" %*
exit /b %ERRORLEVEL%
