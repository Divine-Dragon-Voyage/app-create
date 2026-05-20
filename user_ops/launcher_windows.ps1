param(
    [string]$ProjectRoot = "",
    [string]$DataDir = "C:\app-create-data",
    [bool]$AutoLaunchBrowser = $true
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Show-ErrorAndExit {
    param([string]$Message, [int]$Code = 1)
    [System.Windows.Forms.MessageBox]::Show($Message, "应用创建工具", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    exit $Code
}

function Show-RunSummaryDialog {
    param(
        [string]$SummaryPath,
        [int]$ExitCode
    )

    if ([string]::IsNullOrWhiteSpace($SummaryPath)) {
        return
    }
    if (-not (Test-Path $SummaryPath -PathType Leaf)) {
        return
    }

    try {
        $payload = Get-Content -Path $SummaryPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $failedCount = 0
        $successCount = 0
        $successItems = @()
        if ($payload.failedCount -ne $null) {
            $failedCount = [int]$payload.failedCount
        } elseif ($payload.failed -is [System.Array]) {
            $failedCount = $payload.failed.Count
        }
        if ($payload.success -ne $null) {
            $successCount = [int]$payload.success
        }
        if ($payload.successItems -is [System.Array]) {
            $successItems = $payload.successItems
        }

        $lines = @(
            "Total loaded: $($payload.totalLoaded)",
            "Planned: $($payload.planned)",
            "Success: $successCount",
            "Failed: $failedCount"
        )

        if ($successItems.Count -gt 0) {
            $lines += "Successful apps:"
            foreach ($item in $successItems) {
                $lines += "[OK] $($item.appName) ($($item.packageName))"
            }
        }

        if ($failedCount -gt 0 -and $payload.failed) {
            $lines += "Failed apps:"
            foreach ($item in $payload.failed) {
                $lines += "[FAIL] $($item.appName) ($($item.packageName))"
            }
        }

        $text = ($lines -join "`r`n")
        $icon = if ($ExitCode -eq 0) {
            [System.Windows.Forms.MessageBoxIcon]::Information
        } else {
            [System.Windows.Forms.MessageBoxIcon]::Warning
        }

        [System.Windows.Forms.MessageBox]::Show(
            $text,
            "Run Summary",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            $icon
        ) | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to read run summary: $($_.Exception.Message)`r`nFile: $SummaryPath",
            "Run Summary",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
}

function Build-DeveloperUrl {
    param([string]$InputText)

    if ([string]::IsNullOrWhiteSpace($InputText)) {
        $text = ""
    }
    else {
        $text = $InputText.Trim()
    }

    if (-not $text) {
        throw "请填写开发者 ID 或 Play Console 链接。"
    }

    if ($text -match "^\d+$") {
        return "https://play.google.com/console/u/0/developers/$text/app-list"
    }

    if ($text -match "^https?://") {
        try {
            $uri = [Uri]$text
        }
        catch {
            throw "开发者链接格式不正确。"
        }

        if ($uri.Host -ne "play.google.com") {
            throw "开发者链接域名必须是 play.google.com。"
        }

        $full = "$($uri.Scheme)://$($uri.Host)$($uri.AbsolutePath)"
        $match = [regex]::Match($full, "^(https?://play\.google\.com/(?:console(?:/u/\d+)?/)?developers/(\d+))")
        if (-not $match.Success) {
            throw "开发者链接必须包含 /developers/<id>。"
        }

        return "$($match.Groups[1].Value)/app-list"
    }

    throw "请输入纯数字开发者 ID，或完整的 Play Console 开发者链接。"
}

if (-not $ProjectRoot) {
    $ProjectRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
}

$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$bootstrapPath = Join-Path $ProjectRoot "bootstrap_windows.ps1"
if (-not (Test-Path $bootstrapPath)) {
    throw "未找到 bootstrap_windows.ps1: $bootstrapPath"
}

if (-not (Test-Path $DataDir)) {
    New-Item -Path $DataDir -ItemType Directory -Force | Out-Null
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "应用创建启动器"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(760, 440)
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)

$labelIntro = New-Object System.Windows.Forms.Label
$labelIntro.Text = "请填写开发者 ID（或链接），并选择本次运行要读取的数据文件（Excel/CSV）。"
$labelIntro.Location = New-Object System.Drawing.Point(24, 20)
$labelIntro.Size = New-Object System.Drawing.Size(700, 22)
$form.Controls.Add($labelIntro)

$labelDeveloper = New-Object System.Windows.Forms.Label
$labelDeveloper.Text = "开发者 ID 或链接"
$labelDeveloper.Location = New-Object System.Drawing.Point(24, 56)
$labelDeveloper.Size = New-Object System.Drawing.Size(180, 20)
$form.Controls.Add($labelDeveloper)

$txtDeveloper = New-Object System.Windows.Forms.TextBox
$txtDeveloper.Location = New-Object System.Drawing.Point(24, 78)
$txtDeveloper.Size = New-Object System.Drawing.Size(700, 24)
$form.Controls.Add($txtDeveloper)

$labelExcel = New-Object System.Windows.Forms.Label
$labelExcel.Text = "数据文件路径（Excel/CSV）"
$labelExcel.Location = New-Object System.Drawing.Point(24, 114)
$labelExcel.Size = New-Object System.Drawing.Size(180, 20)
$form.Controls.Add($labelExcel)

$txtExcel = New-Object System.Windows.Forms.TextBox
$txtExcel.Location = New-Object System.Drawing.Point(24, 136)
$txtExcel.Size = New-Object System.Drawing.Size(590, 24)
$form.Controls.Add($txtExcel)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "浏览..."
$btnBrowse.Location = New-Object System.Drawing.Point(624, 134)
$btnBrowse.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnBrowse)

$labelWebUser = New-Object System.Windows.Forms.Label
$labelWebUser.Text = "网页账号（web_username）"
$labelWebUser.Location = New-Object System.Drawing.Point(24, 172)
$labelWebUser.Size = New-Object System.Drawing.Size(220, 20)
$form.Controls.Add($labelWebUser)

$txtWebUser = New-Object System.Windows.Forms.TextBox
$txtWebUser.Location = New-Object System.Drawing.Point(24, 194)
$txtWebUser.Size = New-Object System.Drawing.Size(700, 24)
$form.Controls.Add($txtWebUser)

$labelWebPass = New-Object System.Windows.Forms.Label
$labelWebPass.Text = "网页密码（web_password）"
$labelWebPass.Location = New-Object System.Drawing.Point(24, 230)
$labelWebPass.Size = New-Object System.Drawing.Size(220, 20)
$form.Controls.Add($labelWebPass)

$txtWebPass = New-Object System.Windows.Forms.TextBox
$txtWebPass.Location = New-Object System.Drawing.Point(24, 252)
$txtWebPass.Size = New-Object System.Drawing.Size(700, 24)
$txtWebPass.UseSystemPasswordChar = $true
$form.Controls.Add($txtWebPass)

$labelContactEmail = New-Object System.Windows.Forms.Label
$labelContactEmail.Text = "联系邮箱（contact_email）"
$labelContactEmail.Location = New-Object System.Drawing.Point(24, 288)
$labelContactEmail.Size = New-Object System.Drawing.Size(220, 20)
$form.Controls.Add($labelContactEmail)

$txtContactEmail = New-Object System.Windows.Forms.TextBox
$txtContactEmail.Location = New-Object System.Drawing.Point(24, 310)
$txtContactEmail.Size = New-Object System.Drawing.Size(700, 24)
$form.Controls.Add($txtContactEmail)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "开始运行"
$btnStart.Location = New-Object System.Drawing.Point(544, 350)
$btnStart.Size = New-Object System.Drawing.Size(85, 30)
$form.Controls.Add($btnStart)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "取消"
$btnCancel.Location = New-Object System.Drawing.Point(639, 350)
$btnCancel.Size = New-Object System.Drawing.Size(85, 30)
$form.Controls.Add($btnCancel)

$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "数据文件 (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|所有文件 (*.*)|*.*"
$openFileDialog.Multiselect = $false
$openFileDialog.Title = "选择数据文件（Excel/CSV）"

$selectedDeveloperUrl = $null
$selectedExcelPath = $null
$selectedWebUsername = $null
$selectedWebPassword = $null
$selectedContactEmail = $null
$startRequested = $false

$btnBrowse.Add_Click({
    if ($txtExcel.Text -and (Test-Path $txtExcel.Text)) {
        $openFileDialog.InitialDirectory = Split-Path -Parent $txtExcel.Text
    }

    $result = $openFileDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtExcel.Text = $openFileDialog.FileName
    }
})

$btnCancel.Add_Click({
    $form.Close()
})

$btnStart.Add_Click({
    try {
        $devUrl = Build-DeveloperUrl -InputText $txtDeveloper.Text

        if ([string]::IsNullOrWhiteSpace($txtExcel.Text)) {
            $excelPathRaw = ""
        }
        else {
            $excelPathRaw = $txtExcel.Text.Trim()
        }

        if (-not $excelPathRaw) {
            throw "请填写数据文件路径。"
        }

        if (-not (Test-Path $excelPathRaw -PathType Leaf)) {
            throw "数据文件不存在，请检查路径。"
        }

        $excelPath = (Resolve-Path $excelPathRaw -ErrorAction Stop).Path
        $ext = [IO.Path]::GetExtension($excelPath).ToLowerInvariant()
        if (@(".xlsx", ".xls", ".csv") -notcontains $ext) {
            throw "数据文件必须是 .xlsx / .xls / .csv 格式。"
        }

        $webUsername = $txtWebUser.Text.Trim()
        if (-not $webUsername) {
            throw "请填写网页账号（web_username）。"
        }

        $webPassword = $txtWebPass.Text
        if ([string]::IsNullOrWhiteSpace($webPassword)) {
            throw "请填写网页密码（web_password）。"
        }

        $contactEmail = $txtContactEmail.Text.Trim()
        if (-not $contactEmail) {
            throw "请填写联系邮箱（contact_email）。"
        }
        if ($contactEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            throw "联系邮箱格式不正确。"
        }

        $script:selectedDeveloperUrl = $devUrl
        $script:selectedExcelPath = $excelPath
        $script:selectedWebUsername = $webUsername
        $script:selectedWebPassword = $webPassword
        $script:selectedContactEmail = $contactEmail
        $script:startRequested = $true
        $form.Close()
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "输入检查", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
})

[void]$form.ShowDialog()

if (-not $startRequested) {
    exit 1
}

$developerConfigPath = Join-Path $DataDir "developer_url.txt"
Set-Content -Path $developerConfigPath -Value $selectedDeveloperUrl -Encoding UTF8
$runSummaryPath = Join-Path $DataDir "last_run_summary.json"
if (Test-Path $runSummaryPath -PathType Leaf) {
    Remove-Item -Path $runSummaryPath -Force -ErrorAction SilentlyContinue
}
$env:APP_CREATE_RUN_SUMMARY_PATH = $runSummaryPath
$env:APP_CREATE_WEB_USERNAME = $selectedWebUsername
$env:APP_CREATE_WEB_PASSWORD = $selectedWebPassword
$env:APP_CREATE_CONTACT_EMAIL = $selectedContactEmail

$args = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", "`"$bootstrapPath`"",
    "-RunApp",
    "-HoldWindowOnSuccess",
    "-ExcelFile", "`"$selectedExcelPath`"",
    "-DeveloperUrl", "`"$selectedDeveloperUrl`""
)
if ($AutoLaunchBrowser) {
    $args += "-AutoLaunchBrowser"
}

$process = Start-Process -FilePath "powershell.exe" -ArgumentList $args -Wait -PassThru
Show-RunSummaryDialog -SummaryPath $runSummaryPath -ExitCode $process.ExitCode
exit $process.ExitCode
