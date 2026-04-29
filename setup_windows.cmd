@echo off
setlocal

set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%bootstrap_windows.ps1" -AutoLaunchBrowser

if errorlevel 1 (
  echo.
  echo Setup failed. Please check the error output above.
  pause
  exit /b 1
)

echo.
echo Setup done.
pause
