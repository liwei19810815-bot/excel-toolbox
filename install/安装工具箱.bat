@echo off
REM ============================================================================
REM  Excel Toolbox - one-click installer entry point.
REM
REM  This .bat is deliberately ASCII-only. Chinese text inside a .bat depends on
REM  the console code page; on a machine with a different code page it turns
REM  into garbage before anyone can read it. All user-facing Chinese lives in
REM  the PowerShell script, which handles Unicode properly.
REM ============================================================================
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Toolbox.ps1" %*
echo.
pause
endlocal
