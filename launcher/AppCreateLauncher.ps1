param(
    [string]$InstallDir = "C:\app-create",
    [string]$DataDir = "C:\app-create-data",
    [string]$PackageUrl = "",
    [int]$UpdateCheckIntervalMinutes = 15,
    [switch]$ForceUpdate,
    [switch]$SkipSetup,
    [switch]$AutoLaunchBrowser
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
    if (-not (Test-Path -LiteralPath $PathValue)) {
        New-Item -Path $PathValue -ItemType Directory -Force | Out-Null
    }
}

function Resolve-PackageUrl {
    param(
        [string]$InputUrl,
        [string]$BaseDir
    )

    if ($InputUrl) {
        return $InputUrl.Trim()
    }

    $urlFile = Join-Path $BaseDir "release_url.txt"
    if (Test-Path -LiteralPath $urlFile) {
        $line = Get-Content -LiteralPath $urlFile |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith("#") } |
            Select-Object -First 1

        if ($line -and ($line -notmatch "REPLACE_WITH")) {
            return $line
        }
    }

    throw "Package URL is not configured. Please edit launcher/release_url.txt first."
}

function Get-RemoteFingerprint {
    param([Parameter(Mandatory = $true)][string]$Url)

    $response = Invoke-WebRequest -Uri $Url -Method Head -MaximumRedirection 10 -UseBasicParsing -TimeoutSec 25
    $finalUri = ""
    try {
        $finalUri = $response.BaseResponse.ResponseUri.AbsoluteUri
    }
    catch {
        $finalUri = $Url
    }

    $etag = [string]$response.Headers["ETag"]
    $lastModified = [string]$response.Headers["Last-Modified"]
    $contentLength = [string]$response.Headers["Content-Length"]

    return "etag=$etag|last=$lastModified|len=$contentLength|uri=$finalUri"
}

function Load-LauncherState {
    param([Parameter(Mandatory = $true)][string]$StateFile)

    if (-not (Test-Path -LiteralPath $StateFile)) {
        return [pscustomobject]@{}
    }

    try {
        $raw = Get-Content -LiteralPath $StateFile -Raw
        if (-not $raw) {
            return [pscustomobject]@{}
        }

        $obj = $raw | ConvertFrom-Json
        if ($null -eq $obj) {
            return [pscustomobject]@{}
        }

        return $obj
    }
    catch {
        Write-WarnLog "State file is invalid, will reset: $StateFile"
        return [pscustomobject]@{}
    }
}

function Save-LauncherState {
    param(
        [Parameter(Mandatory = $true)][string]$StateFile,
        [Parameter(Mandatory = $true)][object]$State
    )

    $json = $State | ConvertTo-Json -Depth 5
    $json | Set-Content -LiteralPath $StateFile -Encoding UTF8
}

function Prompt-UpdateDecision {
    param([string]$InstallPath)

    if (-not [Environment]::UserInteractive) {
        Write-WarnLog "No interactive desktop session detected. Will update automatically."
        return "update"
    }

    $message = @"
检测到可用新版本，是否立即更新？

是(Y)：立即更新并启动
否(N)：跳过本次更新，直接启动当前版本
取消(C)：退出，不启动
"@

    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        $result = [System.Windows.Forms.MessageBox]::Show(
            $message,
            "App Create 更新提示",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question,
            [System.Windows.Forms.MessageBoxDefaultButton]::Yes
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { return "update" }
        if ($result -eq [System.Windows.Forms.DialogResult]::No) { return "skip" }
        return "cancel"
    }
    catch {
        Write-WarnLog "Cannot show update prompt. Fallback to auto update. Error: $($_.Exception.Message)"
        return "update"
    }
}

function New-LauncherStateSnapshot {
    param(
        [object]$CurrentState,
        [string]$PackageUrlValue,
        [string]$RemoteFingerprintValue,
        [bool]$Updated
    )

    $nowUtc = (Get-Date).ToUniversalTime().ToString("o")
    $currentLastUpdate = [string]$CurrentState.lastUpdateUtc

    return [pscustomobject]@{
        packageUrl = $PackageUrlValue
        remoteFingerprint = if ($RemoteFingerprintValue) { $RemoteFingerprintValue } else { [string]$CurrentState.remoteFingerprint }
        lastCheckUtc = $nowUtc
        lastUpdateUtc = if ($Updated) { $nowUtc } else { $currentLastUpdate }
    }
}

function Invoke-Deploy {
    param(
        [string]$LauncherDir,
        [string]$ResolvedUrl,
        [string]$InstallPath,
        [string]$DataPath,
        [switch]$SkipSetupSwitch,
        [switch]$AutoLaunchSwitch
    )

    $deployScript = Join-Path $LauncherDir "deploy_windows.ps1"
    if (-not (Test-Path -LiteralPath $deployScript)) {
        throw "deploy_windows.ps1 not found in launcher directory: $deployScript"
    }

    $args = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$deployScript`"",
        "-Mode", "installOrUpdate",
        "-PackageUrl", "`"$ResolvedUrl`"",
        "-InstallDir", "`"$InstallPath`"",
        "-DataDir", "`"$DataPath`""
    )

    $launcherEntry = Join-Path $LauncherDir "AppCreateLauncher.cmd"
    if (Test-Path -LiteralPath $launcherEntry) {
        $args += @("-ShortcutTargetPath", "`"$launcherEntry`"")
    }

    if ($SkipSetupSwitch) {
        $args += "-SkipSetup"
    }
    if ($AutoLaunchSwitch) {
        $args += "-AutoLaunchBrowser"
    }

    Write-Step "Running online deploy..."
    $process = Start-Process -FilePath "powershell.exe" -ArgumentList $args -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw "Deploy failed with exit code $($process.ExitCode)."
    }
}

function Start-AppRunner {
    param([string]$RunScriptPath)

    if (-not (Test-Path -LiteralPath $RunScriptPath)) {
        throw "run_windows.cmd not found after deploy: $RunScriptPath"
    }

    Write-Step "Starting App Create..."
    $proc = Start-Process -FilePath $RunScriptPath -WorkingDirectory (Split-Path -Parent $RunScriptPath) -Wait -PassThru
    return $proc.ExitCode
}

function Main {
    if ($env:OS -ne "Windows_NT") {
        throw "Launcher is for Windows only."
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if ($PSBoundParameters.ContainsKey("UpdateCheckIntervalMinutes")) {
        Write-WarnLog "UpdateCheckIntervalMinutes is ignored. Launcher now checks update on every start."
    }

    $launcherDir = Split-Path -Parent $PSCommandPath
    Ensure-Directory -PathValue $DataDir

    $resolvedUrl = Resolve-PackageUrl -InputUrl $PackageUrl -BaseDir $launcherDir
    $runScript = Join-Path $InstallDir "user_ops\run_windows.cmd"
    $stateFile = Join-Path $DataDir "launcher_state.json"

    $state = Load-LauncherState -StateFile $stateFile
    $remoteFingerprint = $null
    $installReady = (Test-Path -LiteralPath $runScript)
    $needsUpdate = $ForceUpdate -or (-not $installReady)
    $updatedThisRun = $false

    if (-not $needsUpdate) {
        try {
            Write-Step "Checking remote package metadata..."
            $remoteFingerprint = Get-RemoteFingerprint -Url $resolvedUrl
            $isNewPackage = ($remoteFingerprint -ne [string]$state.remoteFingerprint -or $resolvedUrl -ne [string]$state.packageUrl)

            if (-not $isNewPackage) {
                Write-Ok "No update detected, will run local install."
            } else {
                Write-Step "New package detected."
                $decision = Prompt-UpdateDecision -InstallPath $InstallDir
                if ($decision -eq "update") {
                    $needsUpdate = $true
                    Write-Step "User chose to update now."
                }
                elseif ($decision -eq "skip") {
                    Write-WarnLog "User skipped update this time. Will run local install."
                }
                else {
                    Write-WarnLog "User canceled launch."
                    $state = New-LauncherStateSnapshot -CurrentState $state -PackageUrlValue $resolvedUrl -RemoteFingerprintValue $remoteFingerprint -Updated $false
                    Save-LauncherState -StateFile $stateFile -State $state
                    exit 2
                }
            }
        }
        catch {
            Write-WarnLog "Update check failed: $($_.Exception.Message)"
            Write-WarnLog "Will continue with local install."
        }
    }

    if ($needsUpdate) {
        if (-not $remoteFingerprint) {
            try {
                Write-Step "Fetching remote package metadata before deploy..."
                $remoteFingerprint = Get-RemoteFingerprint -Url $resolvedUrl
            }
            catch {
                Write-WarnLog "Remote metadata unavailable, continue deploy anyway."
            }
        }

        Invoke-Deploy -LauncherDir $launcherDir -ResolvedUrl $resolvedUrl -InstallPath $InstallDir -DataPath $DataDir -SkipSetupSwitch:$SkipSetup -AutoLaunchSwitch:$AutoLaunchBrowser
        $updatedThisRun = $true

        if (-not $remoteFingerprint) {
            try {
                $remoteFingerprint = Get-RemoteFingerprint -Url $resolvedUrl
            }
            catch {
                Write-WarnLog "Could not refresh remote metadata after deploy."
            }
        }
    }

    $state = New-LauncherStateSnapshot -CurrentState $state -PackageUrlValue $resolvedUrl -RemoteFingerprintValue $remoteFingerprint -Updated $updatedThisRun
    Save-LauncherState -StateFile $stateFile -State $state

    $exitCode = Start-AppRunner -RunScriptPath $runScript
    exit $exitCode
}

Main
