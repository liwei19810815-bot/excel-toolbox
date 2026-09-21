@echo off
REM ============================================================================
REM  Excel Toolbox - single entry point (install / uninstall / exit).
REM
REM  The script detects whether the add-in is already installed and asks what
REM  to do. IT can bypass the prompt with -Install or -Uninstall.
REM
REM  This .bat is deliberately ASCII-only. Chinese text inside a .bat depends on
REM  the console code page; on a machine with a different code page it turns
REM  into garbage before the user can read anything useful. All user-facing
REM  Chinese lives in the PowerShell script, which handles Unicode properly.
REM ============================================================================
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Toolbox.ps1" %*
echo.
pause
endlocal
