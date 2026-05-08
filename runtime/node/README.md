# Embedded Node Runtime

This directory is used for offline Node runtime packages shipped with the project.

Required embedded file (recommended):

- `node-v20.20.2-win-x64.zip`

Default preparation (x64 only):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tech_ops\prepare_embedded_node.ps1 -NodeVersion 20.20.2
```

Optional preparation (include x86 archive too):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tech_ops\prepare_embedded_node.ps1 -NodeVersion 20.20.2 -IncludeX86
```

Bootstrap behavior:

1. Prefer system `node` in `PATH` when version is `>=18`.
2. If unavailable/incompatible on 64-bit OS, try embedded `x64`.
3. If embedded runtime is missing/unusable, download/install MSI online.
4. `x86` runtime is installed online on demand (not required in the package by default).
