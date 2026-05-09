@echo off
setlocal

for %%I in ("%~dp0..") do set "ROOT_DIR=%%~fI"
cd /d "%ROOT_DIR%"

echo Select release mode:
echo   [1] Build zip package
echo   [2] Build EXE installer (requires Inno Setup 6)
set /p MODE=Enter 1 or 2 [2]: 
if "%MODE%"=="" set "MODE=2"

if "%MODE%"=="1" (
  npm run release:zip:win
) else (
  npm run release:exe:win
)

if errorlevel 1 (
  echo.
  echo Release build failed.
  pause
  exit /b 1
)

echo.
echo Release build finished.
pause
