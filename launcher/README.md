# App Create Launcher

This launcher provides a lightweight distribution mode:

1. Launcher checks remote package metadata.
2. If package changed (or app missing), it runs `deploy_windows.ps1` to update `C:\app-create`.
3. Launcher then starts `C:\app-create\user_ops\run_windows.cmd`.

## Files
- `AppCreateLauncher.ps1`: update-check + launch logic
- `AppCreateLauncher.cmd`: double-click entry
- `release_url.txt`: package URL config

## Build launcher installer
```powershell
npm run release:launcher:exe:win -- -InnoCompilerPath "D:\Inno Setup 6\ISCC.exe" -PackageUrl "https://your-domain.com/app-create-latest.zip"
```
