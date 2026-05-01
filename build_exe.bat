@echo off
REM ============================================================
REM Repo-root convenience wrapper.
REM Forwards to standarddocapp\build_exe.bat which performs the
REM real PyInstaller build and writes
REM   standarddocapp\dist\StandardDocApp.exe
REM ============================================================

setlocal
call "%~dp0standarddocapp\build_exe.bat" %*
exit /b %ERRORLEVEL%
