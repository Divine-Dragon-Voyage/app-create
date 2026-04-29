@echo off
setlocal

set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%deploy_windows.ps1" -Mode installOrUpdate -AutoLaunchBrowser

if errorlevel 1 (
  echo.
  echo Install failed. Please check the error output above.
  pause
  exit /b 1
)

echo.
echo Install finished.
pause
