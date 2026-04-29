@echo off
setlocal

set SCRIPT_DIR=%~dp0
for %%I in ("%SCRIPT_DIR%..") do set "ROOT_DIR=%%~fI"
set DATA_DIR=C:\app-create-data
set EXCEL_FILE=%DATA_DIR%\apps.xlsx
set URL_CONFIG_FILE=%DATA_DIR%\developer_url.txt

if not exist "%DATA_DIR%" (
  mkdir "%DATA_DIR%"
)

if not exist "%URL_CONFIG_FILE%" (
  if exist "%ROOT_DIR%\developer_url.txt" (
    copy /y "%ROOT_DIR%\developer_url.txt" "%URL_CONFIG_FILE%" >nul
  )
)

if not exist "%URL_CONFIG_FILE%" (
  (
    echo # Paste your Play Console developer URL below ^(single line^).
    echo # Example:
    echo # https://play.google.com/console/u/0/developers/1234567890123456789/app-list
    echo.
    echo https://play.google.com/console/u/0/developers/REPLACE_WITH_YOUR_DEVELOPER_ID/app-list
  ) > "%URL_CONFIG_FILE%"
  echo Developer URL config created: %URL_CONFIG_FILE%
  echo Please open this file, paste your Play Console developer link, then run again.
  pause
  exit /b 1
)

findstr /C:"REPLACE_WITH_YOUR_DEVELOPER_ID" "%URL_CONFIG_FILE%" >nul
if not errorlevel 1 (
  echo Please edit this file first and paste your real developer URL:
  echo   %URL_CONFIG_FILE%
  pause
  exit /b 1
)

if not exist "%EXCEL_FILE%" (
  echo Excel not found: %EXCEL_FILE%
  echo Please put your Excel file there and run again.
  pause
  exit /b 1
)

set APP_CREATE_CONFIG_DIR=%DATA_DIR%
cd /d "%ROOT_DIR%"
npm run start -- "%EXCEL_FILE%"

echo.
echo Task finished.
pause
