@echo off
setlocal

for %%I in ("%~dp0..") do set "ROOT_DIR=%%~fI"
powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT_DIR%\deploy_windows.ps1" -Mode update -AutoLaunchBrowser

if errorlevel 1 (
  echo.
  echo Update failed. Please check the error output above.
  pause
  exit /b 1
)

echo.
echo Update finished.
pause
