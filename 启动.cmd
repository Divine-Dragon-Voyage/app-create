@echo off
setlocal

for %%I in ("%~dp0.") do set "ROOT_DIR=%%~fI"
set "RUN_CMD=%ROOT_DIR%\user_ops\run_windows.cmd"

if not exist "%RUN_CMD%" (
  echo [ERROR] run entry not found: %RUN_CMD%
  echo Please make sure the zip is fully extracted.
  pause
  exit /b 1
)

set "APP_CREATE_RUN_CMD=%RUN_CMD%"

echo [STEP] Creating/refreshing desktop shortcut...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop';" ^
  "$run=[System.IO.Path]::GetFullPath($env:APP_CREATE_RUN_CMD);" ^
  "$workDir=[System.IO.Path]::GetDirectoryName($run);" ^
  "$desktopList=@([Environment]::GetFolderPath('Desktop'),[Environment]::GetFolderPath('CommonDesktopDirectory')) | Where-Object { $_ } | Select-Object -Unique;" ^
  "$shell=New-Object -ComObject WScript.Shell;" ^
  "foreach($desktop in $desktopList){$shortcutPath=Join-Path $desktop 'App Create.lnk';$shortcut=$shell.CreateShortcut($shortcutPath);$shortcut.TargetPath=$run;$shortcut.WorkingDirectory=$workDir;$shortcut.Description='Double-click to start App Create automation';$shortcut.IconLocation=($env:SystemRoot + '\\System32\\imageres.dll,2');$shortcut.Save()}" >nul

if errorlevel 1 (
  echo [WARN] Failed to create desktop shortcut, continue anyway.
)

echo [STEP] Starting App Create...
call "%RUN_CMD%"
set "EXIT_CODE=%errorlevel%"

if not "%EXIT_CODE%"=="0" (
  echo.
  echo [ERROR] run_windows.cmd exited with code %EXIT_CODE%.
  echo Press any key to close...
  pause >nul
)

exit /b %EXIT_CODE%
