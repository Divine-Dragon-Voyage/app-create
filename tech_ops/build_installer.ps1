param(
    [string]$OutputDir = ".\\dist",
    [string]$ReleaseName = "app-create-user",
    [string]$InstallerBaseName = "AppCreateSetup",
    [string]$InnoCompilerPath = "",
    [switch]$IncludeSampleExcel,
    [switch]$KeepOldInstallers
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Ensure-Directory {
    param([string]$PathValue)
    if (-not (Test-Path $PathValue)) {
        New-Item -Path $PathValue -ItemType Directory -Force | Out-Null
    }
}

function Resolve-IsccPath {
    param([string]$HintPath)

    if ($HintPath -and (Test-Path $HintPath)) {
        return (Resolve-Path $HintPath).Path
    }

    $fromPath = Get-Command "ISCC.exe" -ErrorAction SilentlyContinue
    if ($fromPath) {
        return $fromPath.Source
    }

    $candidates = @(
        (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\\ISCC.exe"),
        (Join-Path $env:ProgramFiles "Inno Setup 6\\ISCC.exe"),
        "C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe",
        "C:\\Program Files\\Inno Setup 6\\ISCC.exe"
    ) | Where-Object { $_ }

    foreach ($path in $candidates) {
        if (Test-Path $path) {
            return (Resolve-Path $path).Path
        }
    }

    throw "ISCC.exe (Inno Setup compiler) was not found. Please install Inno Setup 6 first."
}

function Remove-OldInstallers {
    param(
        [string]$TargetDir,
        [string]$BaseName
    )

    if (-not (Test-Path -LiteralPath $TargetDir)) {
        return
    }

    $patterns = @(
        "$BaseName*.exe"
    )

    foreach ($pattern in $patterns) {
        $oldFiles = Get-ChildItem -LiteralPath $TargetDir -Filter $pattern -File -ErrorAction SilentlyContinue
        foreach ($file in $oldFiles) {
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
            Write-Host "[STEP] Removed old installer: $($file.Name)"
        }
    }
}

function Main {
    $projectDir = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
    $outputDirResolved = Resolve-Path -LiteralPath $OutputDir -ErrorAction SilentlyContinue
    if ($outputDirResolved) {
        $outputDirAbs = $outputDirResolved.Path
    }
    else {
        $outputDirAbs = [System.IO.Path]::GetFullPath((Join-Path $projectDir $OutputDir))
    }
    Ensure-Directory -PathValue $outputDirAbs

    $isccPath = Resolve-IsccPath -HintPath $InnoCompilerPath
    $issPath = Join-Path $projectDir "tech_ops\\app_create_installer.iss"
    if (-not (Test-Path $issPath)) {
        throw "Inno Setup script not found: $issPath"
    }

    $packageScript = Join-Path $projectDir "package_release.ps1"
    $packageArgs = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$packageScript`"",
        "-OutputDir", "`"$outputDirAbs`"",
        "-ReleaseName", "`"$ReleaseName`""
    )
    if ($IncludeSampleExcel) {
        $packageArgs += "-IncludeSampleExcel"
    }

    Write-Host "[STEP] Building release zip for installer..."
    $zipProcess = Start-Process -FilePath "powershell.exe" -ArgumentList $packageArgs -Wait -PassThru
    if ($zipProcess.ExitCode -ne 0) {
        throw "package_release.ps1 failed with exit code $($zipProcess.ExitCode)."
    }

    $latestZip = Join-Path $outputDirAbs ("$ReleaseName-latest.zip")
    if (-not (Test-Path $latestZip)) {
        throw "Package output not found: $latestZip"
    }

    $tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("app-create-installer-" + [Guid]::NewGuid().ToString("N"))
    $stagingDir = Join-Path $tmpRoot "staging"
    Ensure-Directory -PathValue $stagingDir

    try {
        Write-Host "[STEP] Extracting package..."
        Expand-Archive -LiteralPath $latestZip -DestinationPath $stagingDir -Force

        $appVersion = "1.0.0"
        $packageJsonPath = Join-Path $projectDir "package.json"
        $packageJsonRaw = Get-Content -LiteralPath $packageJsonPath -Raw -ErrorAction SilentlyContinue
        $versionMatch = [regex]::Match($packageJsonRaw, '"version"\s*:\s*"([^"]+)"')
        if ($versionMatch.Success) {
            $appVersion = $versionMatch.Groups[1].Value
        }
        else {
            Write-Host "[WARN] Failed to read package.json version, fallback to 1.0.0"
        }

        $dateTag = Get-Date -Format "yyyyMMdd-HHmmss"
        $outputBase = "{0}-{1}-{2}" -f $InstallerBaseName, $appVersion, $dateTag

        $env:APP_CREATE_INSTALLER_SOURCE = $stagingDir
        $env:APP_CREATE_INSTALLER_OUTPUT = $outputDirAbs

        if (-not $KeepOldInstallers) {
            Remove-OldInstallers -TargetDir $outputDirAbs -BaseName $InstallerBaseName
        }

        Write-Host "[STEP] Compiling EXE installer with Inno Setup..."
        $isccArgs = @(
            "/DMyAppVersion=$appVersion",
            "/DOutputBaseFilename=$outputBase",
            "$issPath"
        )

        & $isccPath @isccArgs
        if ($LASTEXITCODE -ne 0) {
            throw "ISCC compile failed with exit code $LASTEXITCODE."
        }

        $exePath = Join-Path $outputDirAbs ("$outputBase.exe")
        if (-not (Test-Path $exePath)) {
            throw "Installer build finished but output file is missing: $exePath"
        }

        Write-Host "[OK] EXE installer created:"
        Write-Host "     $exePath"
    }
    finally {
        if (Test-Path $tmpRoot) {
            Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        Remove-Item Env:APP_CREATE_INSTALLER_SOURCE -ErrorAction SilentlyContinue
        Remove-Item Env:APP_CREATE_INSTALLER_OUTPUT -ErrorAction SilentlyContinue
    }
}

Main
