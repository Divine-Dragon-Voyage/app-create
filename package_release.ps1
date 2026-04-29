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
            "package.json",
            "package-lock.json",
            "create_app.js",
            "bootstrap_windows.ps1",
            "deploy_windows.ps1",
            "release_url.txt",
            "setup_windows.cmd",
            "install_windows.cmd",
            "update_windows.cmd",
            "run_windows.cmd"
        )

        if ($IncludeSampleExcel) {
            $files += "apps_test_data.xlsx"
        }

        foreach ($name in $files) {
            $src = Join-Path $projectDir $name
            if (Test-Path $src) {
                Copy-Item -Path $src -Destination (Join-Path $stagingDir $name) -Force
            }
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
    } finally {
        if (Test-Path $tempRoot) {
            Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Main
