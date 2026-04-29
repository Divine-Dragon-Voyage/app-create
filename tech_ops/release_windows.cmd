@echo off
setlocal

for %%I in ("%~dp0..") do set "ROOT_DIR=%%~fI"
cd /d "%ROOT_DIR%"
npm run release:zip:win

if errorlevel 1 (
  echo.
  echo Release package failed.
  pause
  exit /b 1
)

echo.
echo Release package finished.
pause
