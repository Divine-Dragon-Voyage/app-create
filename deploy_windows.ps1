param(
    [ValidateSet("install", "update", "installOrUpdate")]
    [string]$Mode = "installOrUpdate",
    [string]$PackageUrl = "",
    [string]$InstallDir = "C:\app-create",
    [string]$DataDir = "C:\app-create-data",
    [int]$KeepBackups = 2,
    [switch]$SkipSetup,
    [switch]$AutoLaunchBrowser,
    [string]$BrowserPath,
    [int]$CdpPort = 9222,
    [int]$CdpWaitSeconds = 30
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Write-Step {
    param([string]$Message)
    Write-Host "[STEP] $Message" -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Message)
    Write-Host "[OK]   $Message" -ForegroundColor Green
}

function Write-WarnLog {
    param([string]$Message)
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][string]$PathValue)
    if (-not (Test-Path $PathValue)) {
        New-Item -Path $PathValue -ItemType Directory -Force | Out-Null
    }
}

function Resolve-PackageUrl {
    param(
        [string]$InputUrl,
        [string]$ScriptDir
    )

    if ($InputUrl) {
        return $InputUrl.Trim()
    }

    $urlFile = Join-Path $ScriptDir "release_url.txt"
    if (Test-Path $urlFile) {
        $line = Get-Content $urlFile |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith("#") } |
            Select-Object -First 1

        if ($line) {
            return $line
        }
    }

    return "https://github.com/taiwuyang1/app-create/archive/refs/heads/main.zip"
}

function Test-ProjectRoot {
    param([Parameter(Mandatory = $true)][string]$DirectoryPath)

    $mustFiles = @(
        "package.json",
        "create_app.js",
        "bootstrap_windows.ps1"
    )

    foreach ($name in $mustFiles) {
        if (-not (Test-Path (Join-Path $DirectoryPath $name))) {
            return $false
        }
    }
    return $true
}

function Find-ExtractedProjectRoot {
    param([Parameter(Mandatory = $true)][string]$ExtractDir)

    if (Test-ProjectRoot -DirectoryPath $ExtractDir) {
        return $ExtractDir
    }

    $oneLevelDirs = Get-ChildItem -Path $ExtractDir -Directory -ErrorAction SilentlyContinue
    foreach ($dir in $oneLevelDirs) {
        if (Test-ProjectRoot -DirectoryPath $dir.FullName) {
            return $dir.FullName
        }
    }

    $deepMatch = Get-ChildItem -Path $ExtractDir -Directory -Recurse -ErrorAction SilentlyContinue |
        Where-Object { Test-ProjectRoot -DirectoryPath $_.FullName } |
        Select-Object -First 1

    if ($deepMatch) {
        return $deepMatch.FullName
    }

    throw "Could not locate project root from extracted package."
}

function Copy-ExcelToDataDir {
    param(
        [string]$SourceDir,
        [string]$TargetDataDir
    )

    if (-not (Test-Path $SourceDir)) {
        return
    }

    Ensure-Directory -PathValue $TargetDataDir

    $excelFiles = Get-ChildItem -Path $SourceDir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in ".xlsx", ".xls" }

    foreach ($file in $excelFiles) {
        $dest = Join-Path $TargetDataDir $file.Name
        if (-not (Test-Path $dest)) {
            Copy-Item -Path $file.FullName -Destination $dest -Force
        }
    }

    $developerUrlFile = Join-Path $SourceDir "developer_url.txt"
    $targetDeveloperUrlFile = Join-Path $TargetDataDir "developer_url.txt"
    if ((Test-Path $developerUrlFile) -and (-not (Test-Path $targetDeveloperUrlFile))) {
        Copy-Item -Path $developerUrlFile -Destination $targetDeveloperUrlFile -Force
    }
}

function Backup-InstallDirectory {
    param(
        [string]$InstallPath,
        [int]$KeepCount
    )

    if (-not (Test-Path $InstallPath)) {
        return
    }

    $baseName = Split-Path -Leaf $InstallPath
    $parentDir = Split-Path -Parent $InstallPath
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupDir = Join-Path $parentDir ($baseName + "_backup_" + $stamp)

    Write-Step "Backing up existing files to: $backupDir"
    New-Item -Path $backupDir -ItemType Directory -Force | Out-Null

    Copy-Item -Path (Join-Path $InstallPath "*") -Destination $backupDir -Recurse -Force

    if ($KeepCount -lt 1) {
        return
    }

    $oldBackups = Get-ChildItem -Path $parentDir -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like ($baseName + "_backup_*") } |
        Sort-Object Name -Descending

    $removeList = $oldBackups | Select-Object -Skip $KeepCount
    foreach ($dir in $removeList) {
        Remove-Item -Path $dir.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Save-CurrentPackageUrl {
    param(
        [string]$InstallPath,
        [string]$Url
    )

    $urlFile = Join-Path $InstallPath "release_url.txt"
    Set-Content -Path $urlFile -Value $Url -Encoding ASCII
}

function Ensure-DataExcelInInstallDir {
    param(
        [string]$InstallPath,
        [string]$TargetDataDir
    )

    $defaultExcelPath = Join-Path $InstallPath "apps.xlsx"
    if (Test-Path $defaultExcelPath) {
        return
    }

    $dataApps = Join-Path $TargetDataDir "apps.xlsx"
    if (Test-Path $dataApps) {
        Copy-Item -Path $dataApps -Destination $defaultExcelPath -Force
        return
    }

    $anyExcel = Get-ChildItem -Path $TargetDataDir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in ".xlsx", ".xls" } |
        Select-Object -First 1
    if ($anyExcel) {
        Copy-Item -Path $anyExcel.FullName -Destination $defaultExcelPath -Force
    }
}

function Ensure-DesktopShortcut {
    param(
        [string]$InstallPath
    )

    $runScriptPath = Join-Path $InstallPath "user_ops\run_windows.cmd"
    if (-not (Test-Path $runScriptPath)) {
        Write-WarnLog "Skip desktop shortcut: run script not found at $runScriptPath"
        return
    }

    $shortcutFileName = "App Create.lnk"
    $shortcutTargets = New-Object System.Collections.Generic.List[string]

    $userDesktop = [Environment]::GetFolderPath("Desktop")
    if ($userDesktop) {
        $shortcutTargets.Add((Join-Path $userDesktop $shortcutFileName))
    }

    $publicDesktop = Join-Path $env:PUBLIC "Desktop"
    if ($env:PUBLIC -and (Test-Path $publicDesktop)) {
        $shortcutTargets.Add((Join-Path $publicDesktop $shortcutFileName))
    }

    $uniqueTargets = $shortcutTargets | Select-Object -Unique
    if (-not $uniqueTargets) {
        Write-WarnLog "Skip desktop shortcut: desktop path not found."
        return
    }

    $shell = New-Object -ComObject WScript.Shell
    foreach ($shortcutPath in $uniqueTargets) {
        $shortcutDir = Split-Path -Parent $shortcutPath
        Ensure-Directory -PathValue $shortcutDir

        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $runScriptPath
        $shortcut.WorkingDirectory = Join-Path $InstallPath "user_ops"
        $shortcut.Description = "Double-click to start App Create automation"
        $shortcut.IconLocation = "%SystemRoot%\System32\imageres.dll,2"
        $shortcut.Save()

        Write-Ok "Desktop shortcut ready: $shortcutPath"
    }
}

function Run-SetupScript {
    param(
        [string]$InstallPath
    )

    $setupScript = Join-Path $InstallPath "bootstrap_windows.ps1"
    if (-not (Test-Path $setupScript)) {
        throw "bootstrap_windows.ps1 not found in install directory."
    }

    Write-Step "Running bootstrap setup..."
    $args = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$setupScript`"",
        "-CdpPort", $CdpPort,
        "-CdpWaitSeconds", $CdpWaitSeconds
    )

    if ($AutoLaunchBrowser) {
        $args += "-AutoLaunchBrowser"
    }
    if ($BrowserPath) {
        $args += @("-BrowserPath", "`"$BrowserPath`"")
    }

    $process = Start-Process -FilePath "powershell.exe" -ArgumentList $args -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw "bootstrap_windows.ps1 failed with exit code $($process.ExitCode)."
    }
}

function Main {
    if ($env:OS -ne "Windows_NT") {
        throw "This script is for Windows only."
    }

    if ($KeepBackups -lt 0) {
        throw "KeepBackups must be >= 0."
    }

    # Older Windows often needs TLS 1.2 explicitly for HTTPS.
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $scriptDir = Split-Path -Parent $PSCommandPath
    $resolvedUrl = Resolve-PackageUrl -InputUrl $PackageUrl -ScriptDir $scriptDir

    if ($Mode -eq "install" -and (Test-Path $InstallDir)) {
        Write-WarnLog "Install directory already exists, continuing as update."
    }

    Write-Step "Package URL: $resolvedUrl"
    Write-Step "Install directory: $InstallDir"
    Write-Step "Data directory: $DataDir"

    $tempRoot = Join-Path $env:TEMP ("app-create-deploy-" + [Guid]::NewGuid().ToString("N"))
    $zipPath = Join-Path $tempRoot "package.zip"
    $extractDir = Join-Path $tempRoot "extract"

    Ensure-Directory -PathValue $tempRoot
    Ensure-Directory -PathValue $extractDir
    Ensure-Directory -PathValue $DataDir

    try {
        Write-Step "Downloading package..."
        Invoke-WebRequest -Uri $resolvedUrl -OutFile $zipPath -UseBasicParsing

        Write-Step "Extracting package..."
        Expand-Archive -LiteralPath $zipPath -DestinationPath $extractDir -Force

        $sourceRoot = Find-ExtractedProjectRoot -ExtractDir $extractDir
        Write-Ok "Detected project root: $sourceRoot"

        Copy-ExcelToDataDir -SourceDir $InstallDir -TargetDataDir $DataDir
        Backup-InstallDirectory -InstallPath $InstallDir -KeepCount $KeepBackups

        Ensure-Directory -PathValue $InstallDir

        Write-Step "Copying new version files..."
        Copy-Item -Path (Join-Path $sourceRoot "*") -Destination $InstallDir -Recurse -Force

        Save-CurrentPackageUrl -InstallPath $InstallDir -Url $resolvedUrl
        Ensure-DataExcelInInstallDir -InstallPath $InstallDir -TargetDataDir $DataDir

        if (-not $SkipSetup) {
            Run-SetupScript -InstallPath $InstallDir
        } else {
            Write-WarnLog "Skipped setup step."
        }

        Ensure-DesktopShortcut -InstallPath $InstallDir

        Write-Host ""
        Write-Ok "Deploy finished."
        Write-Host "Next actions:"
        Write-Host "  1) Double-click desktop shortcut: App Create"
        Write-Host "  2) Or run: $InstallDir\user_ops\run_windows.cmd"
    } finally {
        if (Test-Path $tempRoot) {
            Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Main
