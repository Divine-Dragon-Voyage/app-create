param(
    [string]$NodeVersion = "20.20.2",
    [switch]$SkipCdpCheck,
    [switch]$AutoLaunchBrowser,
    [string]$BrowserPath,
    [string]$BrowserUserDataDir = "C:\chrome-cdp-app-create",
    [int]$CdpPort = 9222,
    [int]$CdpWaitSeconds = 30,
    [string]$BrowserExtraArgs = "",
    [switch]$RunApp,
    [string]$ExcelFile = ""
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# Runtime commands are resolved in Ensure-Node to either system Node or embedded Node.
$script:ResolvedNodeCommand = "node"
$script:ResolvedNpmCommand = "npm"
$script:ResolvedNodeSource = "system-path"

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

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Ensure-Admin {
    if (Test-IsAdmin) {
        return
    }

    Write-WarnLog "Administrator permissions are required. Relaunching with UAC prompt..."
    $argList = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$PSCommandPath`"",
        "-NodeVersion", $NodeVersion,
        "-CdpPort", $CdpPort,
        "-CdpWaitSeconds", $CdpWaitSeconds,
        "-BrowserUserDataDir", "`"$BrowserUserDataDir`""
    )
    if ($SkipCdpCheck) {
        $argList += "-SkipCdpCheck"
    }
    if ($AutoLaunchBrowser) {
        $argList += "-AutoLaunchBrowser"
    }
    if ($BrowserPath) {
        $argList += @("-BrowserPath", "`"$BrowserPath`"")
    }
    if ($BrowserExtraArgs) {
        $argList += @("-BrowserExtraArgs", "`"$BrowserExtraArgs`"")
    }
    if ($RunApp) {
        $argList += "-RunApp"
    }
    if ($ExcelFile) {
        $argList += @("-ExcelFile", "`"$ExcelFile`"")
    }

    $process = Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList $argList -Wait -PassThru
    exit $process.ExitCode
}

function Refresh-Path {
    $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$machinePath;$userPath"
}

function Get-CommandVersion {
    param(
        [Parameter(Mandatory = $true)][string]$Command,
        [string]$PrefixToTrim = ""
    )

    if (-not (Get-Command $Command -ErrorAction SilentlyContinue)) {
        return $null
    }

    try {
        $raw = (& $Command -v 2>$null | Select-Object -First 1)
        if (-not $raw) { return $null }
        $value = $raw.Trim()
        if ($PrefixToTrim -and $value.StartsWith($PrefixToTrim)) {
            $value = $value.Substring($PrefixToTrim.Length)
        }
        return $value
    } catch {
        return $null
    }
}

function Get-NodeVersionFromExecutable {
    param([string]$NodeExecutable)

    if (-not $NodeExecutable -or -not (Test-Path $NodeExecutable)) {
        return $null
    }

    try {
        $raw = (& $NodeExecutable -v 2>$null | Select-Object -First 1)
        if (-not $raw) {
            return $null
        }
        return $raw.Trim().TrimStart("v")
    } catch {
        return $null
    }
}

function Get-NpmVersionFromCommand {
    param([string]$NpmCommand)

    if (-not $NpmCommand) {
        return $null
    }

    try {
        $raw = (& $NpmCommand -v 2>$null | Select-Object -First 1)
        if (-not $raw) {
            return $null
        }
        return $raw.Trim()
    } catch {
        return $null
    }
}

function Get-NodeMajorVersion {
    param([string]$Version)
    if (-not $Version) { return $null }
    $parts = $Version.Split(".")
    if (-not $parts.Length) { return $null }
    [int]$major = 0
    if ([int]::TryParse($parts[0], [ref]$major)) {
        return $major
    }
    return $null
}

function Set-ResolvedNodeCommands {
    param(
        [Parameter(Mandatory = $true)][string]$NodeCommand,
        [Parameter(Mandatory = $true)][string]$NpmCommand,
        [Parameter(Mandatory = $true)][string]$Source
    )

    $script:ResolvedNodeCommand = $NodeCommand
    $script:ResolvedNpmCommand = $NpmCommand
    $script:ResolvedNodeSource = $Source
}

function Resolve-EmbeddedNodeArchiveCandidates {
    param(
        [Parameter(Mandatory = $true)][string]$EmbeddedDir,
        [Parameter(Mandatory = $true)][string]$PreferredVersion
    )

    # Policy: only ship embedded x64 runtime; x86 is downloaded on demand.
    if (-not [Environment]::Is64BitOperatingSystem) {
        return @()
    }

    $archOrder = @("x64")
    $candidates = New-Object System.Collections.Generic.List[object]

    foreach ($arch in $archOrder) {
        $exactFile = Join-Path $EmbeddedDir "node-v$PreferredVersion-win-$arch.zip"
        if (Test-Path $exactFile) {
            $candidates.Add([pscustomobject]@{
                Arch = $arch
                Path = $exactFile
                Reason = "exact-version"
            })
        }

        $fallback = Get-ChildItem -Path $EmbeddedDir -Filter "node-v*-win-$arch.zip" -File -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending
        foreach ($item in $fallback) {
            if ($item.FullName -eq $exactFile) {
                continue
            }
            $candidates.Add([pscustomobject]@{
                Arch = $arch
                Path = $item.FullName
                Reason = "fallback-version"
            })
        }
    }

    return $candidates
}

function Install-EmbeddedNodeRuntime {
    param(
        [Parameter(Mandatory = $true)][string]$ProjectDir,
        [Parameter(Mandatory = $true)][string]$PreferredVersion
    )

    $embeddedDir = Join-Path $ProjectDir "runtime\node"
    if (-not (Test-Path $embeddedDir)) {
        Write-WarnLog "Embedded runtime directory not found: $embeddedDir"
        return $null
    }

    $candidates = Resolve-EmbeddedNodeArchiveCandidates -EmbeddedDir $embeddedDir -PreferredVersion $PreferredVersion
    if (-not $candidates -or $candidates.Count -eq 0) {
        if (-not [Environment]::Is64BitOperatingSystem) {
            Write-WarnLog "32-bit OS detected. Embedded x64 runtime is skipped; x86 will be installed online when needed."
            return $null
        }
        Write-WarnLog "No embedded Node runtime archive found under: $embeddedDir"
        return $null
    }

    foreach ($candidate in $candidates) {
        $archivePath = $candidate.Path
        $arch = $candidate.Arch
        $reason = $candidate.Reason

        try {
            Write-Step "Trying embedded Node archive ($arch/$reason): $(Split-Path -Leaf $archivePath)"
            # Clean old runtime directory to avoid stale files.
            $installRoot = Join-Path $ProjectDir ".runtime\node\embedded-$arch"
            if (Test-Path $installRoot) {
                Remove-Item -LiteralPath $installRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
            Ensure-Directory -PathValue $installRoot

            Expand-Archive -LiteralPath $archivePath -DestinationPath $installRoot -Force

            $nodeExe = Get-ChildItem -Path $installRoot -Filter "node.exe" -File -Recurse -ErrorAction SilentlyContinue |
                Select-Object -First 1
            if (-not $nodeExe) {
                Write-WarnLog "Archive does not contain node.exe: $archivePath"
                continue
            }

            $nodeDir = Split-Path -Parent $nodeExe.FullName
            $npmCmd = Join-Path $nodeDir "npm.cmd"
            if (-not (Test-Path $npmCmd)) {
                Write-WarnLog "Archive does not contain npm.cmd near node.exe: $archivePath"
                continue
            }

            $version = Get-NodeVersionFromExecutable -NodeExecutable $nodeExe.FullName
            $major = Get-NodeMajorVersion -Version $version
            if (-not $major -or $major -lt 18) {
                Write-WarnLog "Embedded Node version is too old (v$version), requires >= 18."
                continue
            }

            return [pscustomobject]@{
                NodeCommand = $nodeExe.FullName
                NpmCommand = $npmCmd
                NodeVersion = $version
                Source = "embedded-zip-$arch"
                BinDir = $nodeDir
                ArchivePath = $archivePath
            }
        } catch {
            Write-WarnLog "Failed to use archive $(Split-Path -Leaf $archivePath): $($_.Exception.Message)"
        }
    }

    return $null
}

function Try-Install-NodeJsMsi {
    param(
        [Parameter(Mandatory = $true)][string]$Version,
        [Parameter(Mandatory = $true)][string]$Arch
    )

    $msiName = "node-v$Version-$Arch.msi"
    $downloadUrl = "https://nodejs.org/dist/v$Version/$msiName"
    $localMsi = Join-Path $env:TEMP ("app-create-" + $msiName)

    try {
        Write-Step "Downloading Node.js MSI ($Arch): $downloadUrl"
        Invoke-WebRequest -Uri $downloadUrl -OutFile $localMsi -UseBasicParsing

        Write-Step "Installing Node.js MSI silently ($Arch)..."
        $msiArgs = "/i `"$localMsi`" /qn /norestart"
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs -Wait -PassThru
        if ($process.ExitCode -ne 0) {
            throw "msiexec exited with code $($process.ExitCode)."
        }

        return $true
    } finally {
        Remove-Item -Path $localMsi -Force -ErrorAction SilentlyContinue
    }
}

function Install-NodeJs {
    param([string]$Version)

    # 64-bit OS: try x64 first; fallback to x86 if needed.
    # 32-bit OS: install x86 directly.
    $archCandidates = if ([Environment]::Is64BitOperatingSystem) { @("x64", "x86") } else { @("x86") }
    $lastError = $null

    foreach ($arch in $archCandidates) {
        try {
            if (Try-Install-NodeJsMsi -Version $Version -Arch $arch) {
                Write-Ok "Node.js installed from MSI ($arch)."
                return
            }
        } catch {
            $lastError = $_
            Write-WarnLog "Node.js MSI install failed for ${arch}: $($_.Exception.Message)"
        }
    }

    if ($lastError) {
        throw "Node.js MSI install failed for all architectures. Last error: $($lastError.Exception.Message)"
    }
    throw "Node.js MSI install failed for all architectures."
}

function Save-RuntimeResolution {
    param([string]$ProjectDir)

    $runtimeDir = Join-Path $ProjectDir ".runtime"
    Ensure-Directory -PathValue $runtimeDir
    $runtimeFile = Join-Path $runtimeDir "node-runtime.json"

    $nodeVersion = if ($script:ResolvedNodeCommand -eq "node") {
        Get-CommandVersion -Command "node" -PrefixToTrim "v"
    } else {
        Get-NodeVersionFromExecutable -NodeExecutable $script:ResolvedNodeCommand
    }
    $npmVersion = Get-NpmVersionFromCommand -NpmCommand $script:ResolvedNpmCommand

    $payload = [ordered]@{
        generatedAt = (Get-Date).ToString("s")
        nodeSource = $script:ResolvedNodeSource
        nodeCommand = $script:ResolvedNodeCommand
        npmCommand = $script:ResolvedNpmCommand
        nodeVersion = $nodeVersion
        npmVersion = $npmVersion
    }

    $payload | ConvertTo-Json -Depth 4 | Set-Content -Path $runtimeFile -Encoding UTF8
    Write-Ok "Runtime info saved: $runtimeFile"
}

function Ensure-Node {
    param([string]$ProjectDir)

    $pathNodeVersion = Get-CommandVersion -Command "node" -PrefixToTrim "v"
    $pathNodeMajor = Get-NodeMajorVersion -Version $pathNodeVersion

    if ($pathNodeMajor -and $pathNodeMajor -ge 18) {
        Set-ResolvedNodeCommands -NodeCommand "node" -NpmCommand "npm" -Source "system-path"
        Write-Ok "Node.js already installed in PATH: v$pathNodeVersion"
        return
    }

    if ($pathNodeVersion) {
        Write-WarnLog "Detected Node.js v$pathNodeVersion (< 18). Will try embedded runtime first."
    } else {
        Write-WarnLog "Node.js not found in PATH. Will try embedded runtime first."
    }

    $embedded = Install-EmbeddedNodeRuntime -ProjectDir $ProjectDir -PreferredVersion $NodeVersion
    if ($embedded) {
        Set-ResolvedNodeCommands -NodeCommand $embedded.NodeCommand -NpmCommand $embedded.NpmCommand -Source $embedded.Source
        # Append PATH only for current process.
        $env:Path = "$($embedded.BinDir);$env:Path"
        Write-Ok "Using embedded Node runtime: v$($embedded.NodeVersion) [$($embedded.Source)]"
        return
    }

    if ($pathNodeVersion) {
        Write-WarnLog "Embedded runtime unavailable. Will install Node.js from official MSI."
    } else {
        Write-WarnLog "Embedded runtime unavailable. Will install Node.js from official MSI."
    }

    Install-NodeJs -Version $NodeVersion
    Refresh-Path

    $newVersion = Get-CommandVersion -Command "node" -PrefixToTrim "v"
    if (-not $newVersion) {
        throw "Node.js install finished but 'node' is still not available in PATH. Open a new terminal and rerun."
    }

    Set-ResolvedNodeCommands -NodeCommand "node" -NpmCommand "npm" -Source "system-msi"
    Write-Ok "Node.js ready from system install: v$newVersion"
}

function Ensure-Npm {
    $npmVersion = Get-NpmVersionFromCommand -NpmCommand $script:ResolvedNpmCommand
    if (-not $npmVersion) {
        throw "npm not found after Node.js setup. Source: $($script:ResolvedNodeSource)"
    }
    Write-Ok "npm ready: v$npmVersion (source: $($script:ResolvedNodeSource))"
}

function Ensure-ProjectDependencies {
    param([string]$ProjectDir)

    $playwrightPkg = Join-Path $ProjectDir "node_modules\playwright\package.json"
    $xlsxPkg = Join-Path $ProjectDir "node_modules\xlsx\package.json"

    if ((Test-Path $playwrightPkg) -and (Test-Path $xlsxPkg)) {
        Write-Ok "Project dependencies already installed."
        return
    }

    Write-Step "Installing npm dependencies..."
    Push-Location $ProjectDir
    try {
        if (Test-Path (Join-Path $ProjectDir "package-lock.json")) {
            & $script:ResolvedNpmCommand ci
        } else {
            & $script:ResolvedNpmCommand install
        }
        if ($LASTEXITCODE -ne 0) {
            throw "npm install step failed with exit code $LASTEXITCODE."
        }
    } finally {
        Pop-Location
    }
    Write-Ok "Project dependencies installed."
}

function Get-CdpEndpointUrl {
    return "http://127.0.0.1:$CdpPort/json/version"
}

function Test-CdpEndpoint {
    param([string]$Url = (Get-CdpEndpointUrl))

    try {
        $response = Invoke-RestMethod -Uri $Url -TimeoutSec 3 -ErrorAction Stop
        return [bool]$response.webSocketDebuggerUrl
    } catch {
        return $false
    }
}

function Resolve-ShortcutTargetPath {
    param([string]$ShortcutPath)

    try {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($ShortcutPath)
        $target = $shortcut.TargetPath
        if ($target -and (Test-Path $target) -and $target.ToLower().EndsWith(".exe")) {
            return $target
        }
    } catch { }

    return $null
}

function Resolve-BrowserPath {
    if ($BrowserPath) {
        if (-not (Test-Path $BrowserPath)) {
            throw "BrowserPath not found: $BrowserPath"
        }
        return (Resolve-Path $BrowserPath).Path
    }

    $standardCandidates = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe"
    )

    foreach ($candidate in $standardCandidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }

    $shortcutRoots = @(
        "$env:USERPROFILE\Desktop",
        "$env:PUBLIC\Desktop",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs",
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
    )

    foreach ($root in $shortcutRoots) {
        if (-not (Test-Path $root)) {
            continue
        }
        $shortcuts = Get-ChildItem -Path $root -Filter "*.lnk" -Recurse -ErrorAction SilentlyContinue
        foreach ($lnk in $shortcuts) {
            if ($lnk.Name -match "HubVPS|Chrome|Edge|Browser") {
                $target = Resolve-ShortcutTargetPath -ShortcutPath $lnk.FullName
                if ($target) {
                    return $target
                }
            }
        }
    }

    return $null
}

function Launch-BrowserWithCdp {
    $resolvedBrowserPath = Resolve-BrowserPath
    if (-not $resolvedBrowserPath) {
        Write-WarnLog "Could not auto-detect browser path."
        Write-Host "      Use -BrowserPath `"C:\Path\To\Browser.exe`" and rerun."
        return $false
    }

    Write-Step "Launching browser with CDP: $resolvedBrowserPath"
    $args = @(
        "--remote-debugging-address=127.0.0.1",
        "--remote-debugging-port=$CdpPort",
        "--user-data-dir=`"$BrowserUserDataDir`"",
        "--no-first-run",
        "--no-default-browser-check"
    )
    if ($BrowserExtraArgs) {
        $args += $BrowserExtraArgs
    }

    Start-Process -FilePath $resolvedBrowserPath -ArgumentList ($args -join " ") | Out-Null
    return $true
}

function Wait-ForCdpEndpoint {
    param([int]$TimeoutSeconds = 30)

    $url = Get-CdpEndpointUrl
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if (Test-CdpEndpoint -Url $url) {
            return $true
        }
        Start-Sleep -Seconds 1
    }
    return $false
}

function Ensure-CdpEndpoint {
    if ($SkipCdpCheck) {
        Write-WarnLog "Skipping CDP endpoint check."
        return
    }

    $endpointUrl = Get-CdpEndpointUrl
    Write-Step "Checking CDP endpoint: $endpointUrl"
    if (Test-CdpEndpoint) {
        Write-Ok "CDP endpoint is reachable."
        return
    }

    if (-not $AutoLaunchBrowser) {
        Write-WarnLog "CDP endpoint not reachable."
        Write-Host "      Start HubVPS/Chrome with --remote-debugging-port=$CdpPort, or rerun with -AutoLaunchBrowser."
        return
    }

    $launched = Launch-BrowserWithCdp
    if (-not $launched) {
        return
    }

    Write-Step "Waiting for CDP endpoint to become available (timeout: ${CdpWaitSeconds}s)..."
    if (Wait-ForCdpEndpoint -TimeoutSeconds $CdpWaitSeconds) {
        Write-Ok "CDP endpoint is reachable after auto-launch."
    } else {
        Write-WarnLog "Browser launched, but CDP endpoint is still not reachable."
        Write-Host "      Verify browser supports CDP and remote debugging flags."
    }
}

function Invoke-AppCreation {
    param(
        [Parameter(Mandatory = $true)][string]$ProjectDir,
        [string]$ExcelFilePath
    )

    $entryScript = Join-Path $ProjectDir "create_app.js"
    if (-not (Test-Path $entryScript)) {
        throw "Entry script not found: $entryScript"
    }

    $args = @($entryScript)
    if ($ExcelFilePath) {
        $args += $ExcelFilePath
    }

    Write-Step "Running app task with Node source: $($script:ResolvedNodeSource)"
    Push-Location $ProjectDir
    try {
        & $script:ResolvedNodeCommand @args
        if ($LASTEXITCODE -ne 0) {
            throw "create_app.js exited with code $LASTEXITCODE."
        }
    } finally {
        Pop-Location
    }
}

function Main {
    if ($env:OS -ne "Windows_NT") {
        throw "This script is for Windows VPS only."
    }
    if ($CdpPort -lt 1 -or $CdpPort -gt 65535) {
        throw "CdpPort must be between 1 and 65535."
    }
    if ($CdpWaitSeconds -lt 1) {
        throw "CdpWaitSeconds must be >= 1."
    }

    Ensure-Admin

    $projectDir = Split-Path -Parent $PSCommandPath
    Write-Step "Project directory: $projectDir"

    # Some older Windows images need explicit TLS 1.2 for HTTPS downloads.
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Ensure-Node -ProjectDir $projectDir
    Ensure-Npm
    Ensure-ProjectDependencies -ProjectDir $projectDir
    Ensure-CdpEndpoint
    Save-RuntimeResolution -ProjectDir $projectDir

    if ($RunApp) {
        Invoke-AppCreation -ProjectDir $projectDir -ExcelFilePath $ExcelFile
        return
    }

    Write-Host ""
    Write-Ok "Bootstrap completed."
    Write-Host "Run app creation with:"
    Write-Host ('  powershell -NoProfile -ExecutionPolicy Bypass -File "{0}\bootstrap_windows.ps1" -RunApp -ExcelFile ".\apps_test_data.xlsx"' -f $projectDir)
}

Main

