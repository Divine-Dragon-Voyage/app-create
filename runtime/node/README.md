# Embedded Node Runtime

This directory is used for offline Node runtime packages shipped with the project.

Required files (recommended):

- `node-v20.20.2-win-x64.zip`
- `node-v20.20.2-win-x86.zip`

You can prepare both archives with:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tech_ops\prepare_embedded_node.ps1 -NodeVersion 20.20.2
```

Bootstrap behavior:

1. Prefer system `node` in `PATH` when version is `>=18`.
2. If unavailable/incompatible, try embedded `x64` first, then `x86`.
3. If embedded archives are missing/unusable, fallback to downloading MSI installer.
