@echo off
setlocal

set SCRIPT_DIR=%~dp0
set DATA_DIR=C:\app-create-data
set EXCEL_FILE=%DATA_DIR%\apps.xlsx

if not exist "%DATA_DIR%" (
  mkdir "%DATA_DIR%"
)

if not exist "%EXCEL_FILE%" (
  echo Excel not found: %EXCEL_FILE%
  echo Please put your Excel file there and run again.
  pause
  exit /b 1
)

cd /d "%SCRIPT_DIR%"
npm run start -- "%EXCEL_FILE%"

echo.
echo Task finished.
pause
