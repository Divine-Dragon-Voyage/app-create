param(
    [string]$OutputDir = ".\\dist",
    [string]$InstallerBaseName = "AppCreateLauncherSetup",
    [string]$InnoCompilerPath = "",
    [string]$PackageUrl = ""
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Ensure-Directory {
    param([string]$PathValue)
    if (-not (Test-Path -LiteralPath $PathValue)) {
        New-Item -Path $PathValue -ItemType Directory -Force | Out-Null
    }
}

function Resolve-IsccPath {
    param([string]$HintPath)

    if ($HintPath -and (Test-Path -LiteralPath $HintPath)) {
        return (Resolve-Path -LiteralPath $HintPath).Path
    }

    $fromPath = Get-Command "ISCC.exe" -ErrorAction SilentlyContinue
    if ($fromPath) {
        return $fromPath.Source
    }

    $candidates = @(
        "D:\\Inno Setup 6\\ISCC.exe",
        (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\\ISCC.exe"),
        (Join-Path $env:ProgramFiles "Inno Setup 6\\ISCC.exe")
    ) | Where-Object { $_ }

    foreach ($path in $candidates) {
        if (Test-Path -LiteralPath $path) {
            return (Resolve-Path -LiteralPath $path).Path
        }
    }

    throw "ISCC.exe (Inno Setup compiler) was not found. Please install Inno Setup 6 first."
}

function Get-FirstConfigLine {
    param([string]$FilePath)

    if (-not (Test-Path -LiteralPath $FilePath)) {
        return ""
    }

    $line = Get-Content -LiteralPath $FilePath |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and -not $_.StartsWith("#") } |
        Select-Object -First 1

    if (-not $line) {
        return ""
    }

    if ($line -match "REPLACE_WITH") {
        return ""
    }

    return $line
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
    $issPath = Join-Path $projectDir "tech_ops\\app_create_launcher_installer.iss"
    if (-not (Test-Path -LiteralPath $issPath)) {
        throw "Inno Setup script not found: $issPath"
    }

    $tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("app-create-launcher-installer-" + [Guid]::NewGuid().ToString("N"))
    $stagingDir = Join-Path $tmpRoot "staging"
    Ensure-Directory -PathValue $stagingDir

    try {
        $launcherDir = Join-Path $projectDir "launcher"
        $requiredFiles = @(
            (Join-Path $launcherDir "AppCreateLauncher.ps1"),
            (Join-Path $launcherDir "AppCreateLauncher.cmd"),
            (Join-Path $projectDir "deploy_windows.ps1")
        )

        foreach ($src in $requiredFiles) {
            if (-not (Test-Path -LiteralPath $src)) {
                throw "Required launcher file not found: $src"
            }
            Copy-Item -LiteralPath $src -Destination (Join-Path $stagingDir (Split-Path -Leaf $src)) -Force
        }

        $resolvedUrl = ""
        if ($PackageUrl) {
            $resolvedUrl = $PackageUrl.Trim()
        }
        if (-not $resolvedUrl) {
            $resolvedUrl = Get-FirstConfigLine -FilePath (Join-Path $projectDir "release_url.txt")
        }
        if (-not $resolvedUrl) {
            $resolvedUrl = Get-FirstConfigLine -FilePath (Join-Path $launcherDir "release_url.txt")
        }

        $releaseFile = Join-Path $stagingDir "release_url.txt"
        if ($resolvedUrl) {
            Set-Content -LiteralPath $releaseFile -Value $resolvedUrl -Encoding ASCII
        }
        else {
            Copy-Item -LiteralPath (Join-Path $launcherDir "release_url.txt") -Destination $releaseFile -Force
        }

        $appVersion = "1.0.0"
        $packageJsonPath = Join-Path $projectDir "package.json"
        $packageJsonRaw = Get-Content -LiteralPath $packageJsonPath -Raw -ErrorAction SilentlyContinue
        $versionMatch = [regex]::Match($packageJsonRaw, '"version"\s*:\s*"([^"]+)"')
        if ($versionMatch.Success) {
            $appVersion = $versionMatch.Groups[1].Value
        }

        $dateTag = Get-Date -Format "yyyyMMdd-HHmmss"
        $outputBase = "{0}-{1}-{2}" -f $InstallerBaseName, $appVersion, $dateTag

        $env:APP_CREATE_LAUNCHER_SOURCE = $stagingDir
        $env:APP_CREATE_LAUNCHER_OUTPUT = $outputDirAbs

        Write-Host "[STEP] Compiling launcher installer EXE..."
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
        if (-not (Test-Path -LiteralPath $exePath)) {
            throw "Launcher installer build finished but output file is missing: $exePath"
        }

        Write-Host "[OK] Launcher installer created:"
        Write-Host "     $exePath"
    }
    finally {
        if (Test-Path -LiteralPath $tmpRoot) {
            Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        Remove-Item Env:APP_CREATE_LAUNCHER_SOURCE -ErrorAction SilentlyContinue
        Remove-Item Env:APP_CREATE_LAUNCHER_OUTPUT -ErrorAction SilentlyContinue
    }
}

Main
