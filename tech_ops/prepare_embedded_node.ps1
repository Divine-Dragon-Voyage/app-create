param(
    [string]$NodeVersion = "20.20.2",
    [string]$TargetDir = "",
    [switch]$IncludeX86
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

function Download-NodeZip {
    param(
        [Parameter(Mandatory = $true)][string]$Version,
        [Parameter(Mandatory = $true)][string]$Arch,
        [Parameter(Mandatory = $true)][string]$OutDir
    )

    $fileName = "node-v$Version-win-$Arch.zip"
    $url = "https://nodejs.org/dist/v$Version/$fileName"
    $outPath = Join-Path $OutDir $fileName

    Write-Step "Downloading $fileName ..."
    Invoke-WebRequest -Uri $url -OutFile $outPath -UseBasicParsing

    if (-not (Test-Path $outPath)) {
        throw "Download failed: $fileName"
    }

    $size = (Get-Item $outPath).Length
    if ($size -lt 1MB) {
        throw "Downloaded file seems invalid (too small): $outPath"
    }

    Write-Ok "Ready: $outPath"
}

function Main {
    $scriptDir = Split-Path -Parent $PSCommandPath
    $projectDir = Split-Path -Parent $scriptDir

    # 兼容旧版系统默认 TLS 配置，避免下载 nodejs.org 失败
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if (-not $TargetDir) {
        $TargetDir = Join-Path $projectDir "runtime\node"
    }

    Ensure-Directory -PathValue $TargetDir

    # Default strategy: embed both x64 and x86 for full offline coverage.
    $archList = @("x64", "x86")
    if ($IncludeX86) {
        Write-WarnLog "IncludeX86 is now the default behavior and can be omitted."
    }
    $successCount = 0

    foreach ($arch in $archList) {
        try {
            Download-NodeZip -Version $NodeVersion -Arch $arch -OutDir $TargetDir
            $successCount++
        } catch {
            Write-WarnLog "Failed to prepare $arch package: $($_.Exception.Message)"
        }
    }

    if ($successCount -eq 0) {
        throw "No embedded Node package was downloaded."
    }

    Write-Host ""
    Write-Ok "Embedded runtime preparation completed."
    Write-Host "Target directory: $TargetDir"
    Write-Host "Embedded architectures: $($archList -join ', ')"
    Write-Host "Next: run release packaging so runtime zips are included in dist zip."
}

Main
