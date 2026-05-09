# Embedded Node Runtime

This directory is used for offline Node runtime packages shipped with the project.

Required embedded files:

- `node-v20.20.2-win-x64.zip`
- `node-v20.20.2-win-x86.zip`

Default preparation (x64 + x86):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tech_ops\prepare_embedded_node.ps1 -NodeVersion 20.20.2
```

Optional preparation (same behavior, kept for backward compatibility):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tech_ops\prepare_embedded_node.ps1 -NodeVersion 20.20.2 -IncludeX86
```

Bootstrap behavior:

1. Prefer system `node` in `PATH` when version is `>=18`.
2. If unavailable/incompatible on 64-bit OS, try embedded `x64` then `x86`.
3. If unavailable/incompatible on 32-bit OS, try embedded `x86`.
4. If embedded runtime is missing/unusable, download/install MSI online.
