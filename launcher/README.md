# App Create Launcher

This launcher provides a lightweight distribution mode:

1. Launcher checks remote package metadata on every start.
2. If a new package is detected, user chooses whether to update now.
3. If user confirms, launcher runs `deploy_windows.ps1` to update `C:\app-create`.
4. Launcher then starts `C:\app-create\user_ops\run_windows.cmd`.

## Files
- `AppCreateLauncher.ps1`: update-check + prompt + launch logic
- `AppCreateLauncher.cmd`: double-click entry
- `release_url.txt`: package URL config

## Build launcher installer
```powershell
npm run release:launcher:exe:win -- -InnoCompilerPath "D:\Inno Setup 6\ISCC.exe" -PackageUrl "https://your-domain.com/app-create-latest.zip"
```
