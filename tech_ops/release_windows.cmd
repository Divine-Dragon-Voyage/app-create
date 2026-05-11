@echo off
setlocal

for %%I in ("%~dp0..") do set "ROOT_DIR=%%~fI"
cd /d "%ROOT_DIR%"

echo Select release mode:
echo   [1] Build zip package
echo   [2] Build EXE installer (requires Inno Setup 6)
echo   [3] Build launcher EXE (small, online update)
set /p MODE=Enter 1, 2 or 3 [3]: 
if "%MODE%"=="" set "MODE=3"

if "%MODE%"=="1" (
  npm run release:zip:win
) else if "%MODE%"=="2" (
  npm run release:exe:win
) else (
  npm run release:launcher:exe:win
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
