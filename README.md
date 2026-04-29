# app-create

使用 Playwright 自动在 Google Play Console 创建应用，并从 Excel 读取每条应用的名称与包名。

## 功能说明

脚本文件：`create_app.js`

当前实现会：

1. 连接本地已开启 CDP 的 Chrome（`http://localhost:9222`）。
2. 从项目根目录的 Excel 读取 `应用名称`、`应用包名`（或英文同义列名）。
3. 逐行进入开发者控制台并点击 `Create app`。
4. 填写 `App name` 与 `App package name`，并点击 `Check availability`。
5. 随机选择 `App` 或 `Game`，选择 `Free`。
6. 自动勾选创建页声明项并提交创建。
7. 进入 `App content`，依次处理 7 类声明（Ads / App access / Target audience / Advertising ID / Government apps / Financial features / Health apps）。

## Excel 格式

脚本会读取第一个工作表，支持以下列名（任一组合即可）：

- 应用名称 / 应用名 / App Name / Application Name
- 应用包名 / 包名 / App Package Name / Package Name / Application ID

示例：

| 应用名称 | 应用包名 |
|---|---|
| PixelGrid | com.pixelgridabcde.artfghijk |
| FlowBreath | com.flowbreathkjmnt.toolsqwerty |

校验规则：

- `应用名称` 不能为空，且长度 <= 30
- `应用包名` 不能为空，且需符合 `com.example.appname` 格式（小写字母/数字/下划线 + 点分段）

## 运行前准备

### 1) 环境

- Node.js 18+（推荐 LTS）
- Google Chrome（或 Chromium）
- 可访问 Google Play Console 的账号，且已登录

### 2) 安装依赖

在项目根目录执行：

```bash
npm install
```

### Windows VPS 一键初始化（推荐）

如果你的 VPS 是 Windows（例如 HubVPS），可以直接双击：

`setup_windows.cmd`

或 PowerShell 执行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\bootstrap_windows.ps1 -AutoLaunchBrowser
```

或在项目目录执行：

```bash
npm run bootstrap:win:auto
```

这个脚本会自动做以下事情：

- 检测 Node.js（要求 >= 18），没有就自动安装
- 检测 npm
- 检测项目依赖（`playwright` / `xlsx`），缺失就自动安装
- 检测 CDP 端口是否可访问（`http://127.0.0.1:9222/json/version`）
- 可选自动启动浏览器并打开 CDP 调试端口（默认端口 `9222`）

首次运行会自动弹出 UAC 管理员授权窗口，请点允许。

如果你只想做安装，不检查 CDP，可加参数：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\bootstrap_windows.ps1 -SkipCdpCheck
```

如果你要明确指定 HubVPS 浏览器路径：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\bootstrap_windows.ps1 -AutoLaunchBrowser -BrowserPath "C:\Path\To\HubVPSBrowser.exe"
```

### 3) 启动带远程调试端口的 Chrome

macOS 示例：

```bash
open -na "Google Chrome" --args --remote-debugging-port=9222 --user-data-dir=/tmp/chrome-cdp-app-create
```

然后在这个浏览器里手动登录 Google 账号与 Play Console。

Windows（Chrome/HubVPS）示例：

```powershell
"C:\Path\To\Your\Browser.exe" --remote-debugging-port=9222 --user-data-dir="C:\chrome-cdp-app-create"
```

## 使用方式

在项目目录执行：

```bash
npm run start
```

默认行为：

- 自动读取根目录第一个 `.xlsx/.xls` 文件
- 按 Excel 全部有效行执行

指定 Excel 文件：

```bash
npm run start -- ./apps.xlsx
```

脚本在多次执行之间会随机等待 60-180 秒（用于降低频率）。

## 常见问题

### `Cannot find module 'playwright'`

未安装依赖，执行：

```bash
npm install
```

### 找不到 Excel 或列名不匹配

- 确认根目录存在 `.xlsx/.xls` 文件，或显式传文件路径
- 确认表头包含：`应用名称` 与 `应用包名`（或 README 中列出的英文同义名）
- 确认数据在第一个工作表

### `connectOverCDP` 超时 / 连接失败

- 确认 Chrome 已用 `--remote-debugging-port=9222` 启动
- 确认本机 `9222` 端口未被占用或拦截
- 确认脚本机器与 Chrome 在同一台主机

### 页面元素找不到

Google Play Console 页面会变更，若报错可优先检查：

- 当前登录账号是否正确
- 是否进入了预期开发者账户
- 按钮文案或 DOM 结构是否变化（需更新 selector）

## 注意事项

- 本脚本会执行真实创建操作，请先在测试账号验证。
- 声明信息需与应用真实功能一致，避免合规风险。
- 包名冲突时需要更换 Excel 中对应包名后重试。
- 若用于批量操作，建议增加日志持久化和失败截图。
