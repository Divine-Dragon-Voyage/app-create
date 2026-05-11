@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%AppCreateLauncher.ps1"
set "EXIT_CODE=%errorlevel%"

if not "%EXIT_CODE%"=="0" (
  echo.
  echo Launcher failed with exit code %EXIT_CODE%.
  echo Press any key to close...
  pause >nul
)

exit /b %EXIT_CODE%
