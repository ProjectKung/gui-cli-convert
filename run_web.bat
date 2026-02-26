@echo off
setlocal
cd /d "%~dp0"

if /I "%~1"=="--tray" (
  start "" wscript.exe "%~dp0run_web.vbs"
  exit /b 0
)

title Switch Converter Launcher (Console)
chcp 65001 >nul
echo ============================================================
echo              SWITCH CONVERTER WEB LAUNCHER
echo ============================================================
echo  URL: http://127.0.0.1:5000
echo  Mode: Console (show logs)
echo  Tip : Use "run_web.vbs" for tray mode (no console)
echo  Auto-update: ON (GitHub main)
echo.
echo [INFO] Preparing environment and starting web app...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File ".\setup_and_run.ps1" -AutoInstallPython
set "EXIT_CODE=%ERRORLEVEL%"
echo.
if not "%EXIT_CODE%"=="0" (
  echo [ERROR] Program exited with code %EXIT_CODE%.
  echo [TIP] Please read the message above to see what failed.
  echo.
  pause
  exit /b %EXIT_CODE%
)
endlocal
