@echo off
REM One-click build entry point at the repo root.
REM Forwards to standarddocapp\build_tools\build.bat which runs
REM standarddocapp\build_tools\build.ps1 with -ExecutionPolicy Bypass.
REM
REM Usage (from a regular cmd.exe / Explorer double-click):
REM   build_exe.bat
REM   build_exe.bat -SkipDeps
REM   build_exe.bat -Clean

setlocal
set "REPO_ROOT=%~dp0"
call "%REPO_ROOT%standarddocapp\build_tools\build.bat" %*
exit /b %ERRORLEVEL%
