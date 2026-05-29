@echo off
setlocal

set SCRIPT_DIR=%~dp0
for %%I in ("%SCRIPT_DIR%..") do set "ROOT_DIR=%%~fI"
set DATA_DIR=C:\app-create-data

if not exist "%DATA_DIR%" (
  mkdir "%DATA_DIR%"
)

set APP_CREATE_CONFIG_DIR=%DATA_DIR%
cd /d "%ROOT_DIR%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT_DIR%\user_ops\launcher_windows.ps1" -ProjectRoot "%ROOT_DIR%" -DataDir "%DATA_DIR%"

if errorlevel 1 (
  echo.
  echo Task failed or canceled. Please check the error output above.
  echo Press any key to close...
  pause >nul
  exit /b 1
)

echo.
echo Task finished.
echo Press any key to close...
pause >nul
exit /b 0
