@echo off
REM ASCII-only on purpose; see 安装工具箱.bat for why.
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Toolbox.ps1" -Uninstall %*
echo.
pause
endlocal
