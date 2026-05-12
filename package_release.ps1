param(
    [string]$OutputDir = ".\dist",
    [string]$ReleaseName = "app-create",
    [switch]$IncludeSampleExcel
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Ensure-Directory {
    param([string]$PathValue)
    if (-not (Test-Path $PathValue)) {
        New-Item -Path $PathValue -ItemType Directory -Force | Out-Null
    }
}

function Main {
    $projectDir = Split-Path -Parent $PSCommandPath
    Ensure-Directory -PathValue $OutputDir
    $outputDirAbs = (Resolve-Path -Path $OutputDir).Path
    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("app-create-pack-" + [Guid]::NewGuid().ToString("N"))
    $stagingDir = Join-Path $tempRoot "staging"

    Ensure-Directory -PathValue $tempRoot
    Ensure-Directory -PathValue $stagingDir

    try {
        $files = @(
            "README.md",
            "WORKFLOW.md",
            "LAUNCHER_WORKFLOW.md",
            "USER_WORKFLOW.md",
            "TECH_WORKFLOW.md",
            "用户使用文档.md",
            "用户使用文档.txt",
            "开发者维护文档.md",
            "package.json",
            "package-lock.json",
            "create_app.js",
            "bootstrap_windows.ps1",
            "deploy_windows.ps1",
            "developer_url.txt",
            "release_url.txt",
            "浣跨敤璇存槑.md",
            "tech_ops\README.md",
            "tech_ops\app_create_launcher_installer.iss",
            "tech_ops\build_launcher_installer.ps1",
            "tech_ops\build_installer.ps1",
            "tech_ops\prepare_embedded_node.ps1",
            "tech_ops\release_windows.cmd",
            "tech_ops\release_mac_linux.sh",
            "launcher\AppCreateLauncher.ps1",
            "launcher\AppCreateLauncher.cmd",
            "launcher\README.md",
            "launcher\release_url.txt",
            "user_ops\README.md",
            "user_ops\install_windows.cmd",
            "user_ops\update_windows.cmd",
            "user_ops\run_windows.cmd",
            "user_ops\launcher_windows.ps1"
        )

        if ($IncludeSampleExcel) {
            $files += "apps_test_data.xlsx"
        }

        foreach ($name in $files) {
            $src = Join-Path $projectDir $name
            if (Test-Path $src) {
                $dest = Join-Path $stagingDir $name
                $destParent = Split-Path -Parent $dest
                if ($destParent -and -not (Test-Path $destParent)) {
                    New-Item -Path $destParent -ItemType Directory -Force | Out-Null
                }
                Copy-Item -Path $src -Destination $dest -Force
            }
        }

        $folders = @(
            "runtime"
        )

        foreach ($folder in $folders) {
            $srcDir = Join-Path $projectDir $folder
            if (-not (Test-Path $srcDir)) {
                continue
            }

            $destDir = Join-Path $stagingDir $folder
            Copy-Item -Path $srcDir -Destination $destDir -Recurse -Force
        }

        $dateTag = Get-Date -Format "yyyyMMdd-HHmmss"
        $versionZip = Join-Path $outputDirAbs ($ReleaseName + "-" + $dateTag + ".zip")
        $latestZip = Join-Path $outputDirAbs ($ReleaseName + "-latest.zip")

        if (Test-Path $versionZip) {
            Remove-Item -Path $versionZip -Force
        }
        if (Test-Path $latestZip) {
            Remove-Item -Path $latestZip -Force
        }

        Compress-Archive -Path (Join-Path $stagingDir "*") -DestinationPath $versionZip -Force
        Copy-Item -Path $versionZip -Destination $latestZip -Force

        Write-Host "[OK] Release package created:"
        Write-Host "     $versionZip"
        Write-Host "[OK] Latest package refreshed:"
        Write-Host "     $latestZip"
    }
    finally {
        if (Test-Path $tempRoot) {
            Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Main
