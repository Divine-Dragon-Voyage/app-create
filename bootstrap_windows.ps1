param(
    [string]$NodeVersion = "20.20.2",
    [switch]$SkipCdpCheck,
    [switch]$AutoLaunchBrowser,
    [string]$BrowserPath,
    [string]$BrowserUserDataDir = "C:\chrome-cdp-app-create",
    [int]$CdpPort = 9222,
    [int]$CdpWaitSeconds = 30,
    [string]$BrowserExtraArgs = ""
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

function Install-NodeJs {
    param([string]$Version)

    $arch = if ([Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" }
    $msiName = "node-v$Version-$arch.msi"
    $downloadUrl = "https://nodejs.org/dist/v$Version/$msiName"
    $localMsi = Join-Path $env:TEMP $msiName

    Write-Step "Downloading Node.js $Version ($arch) ..."
    Invoke-WebRequest -Uri $downloadUrl -OutFile $localMsi -UseBasicParsing

    Write-Step "Installing Node.js silently..."
    $msiArgs = "/i `"$localMsi`" /qn /norestart"
    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw "Node.js MSI install failed with exit code $($process.ExitCode)."
    }

    Remove-Item $localMsi -Force -ErrorAction SilentlyContinue
    Write-Ok "Node.js installed."
}

function Ensure-Node {
    $nodeVersion = Get-CommandVersion -Command "node" -PrefixToTrim "v"
    $nodeMajor = Get-NodeMajorVersion -Version $nodeVersion

    if ($nodeMajor -and $nodeMajor -ge 18) {
        Write-Ok "Node.js already installed: v$nodeVersion"
        return
    }

    if ($nodeVersion) {
        Write-WarnLog "Detected Node.js v$nodeVersion (< 18). Will upgrade."
    } else {
        Write-WarnLog "Node.js not found. Will install."
    }

    Install-NodeJs -Version $NodeVersion
    Refresh-Path

    $newVersion = Get-CommandVersion -Command "node" -PrefixToTrim "v"
    if (-not $newVersion) {
        throw "Node.js install finished but 'node' is still not available in PATH. Open a new terminal and rerun."
    }
    Write-Ok "Node.js ready: v$newVersion"
}

function Ensure-Npm {
    $npmVersion = Get-CommandVersion -Command "npm"
    if (-not $npmVersion) {
        throw "npm not found after Node.js setup."
    }
    Write-Ok "npm ready: v$npmVersion"
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
            & npm ci
        } else {
            & npm install
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
            if ($lnk.Name -match "HubVPS|防关联|Chrome|Edge|Browser") {
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

    Ensure-Node
    Ensure-Npm
    Ensure-ProjectDependencies -ProjectDir $projectDir
    Ensure-CdpEndpoint

    Write-Host ""
    Write-Ok "Bootstrap completed."
    Write-Host "Run app creation with:"
    Write-Host "  cd `"$projectDir`""
    Write-Host "  npm run start -- .\apps_test_data.xlsx"
}

Main
