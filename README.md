# app-create

使用 Playwright 自动在 Google Play Console 创建应用，并从 Excel 读取每条应用的名称与包名。

快速查看操作步骤请直接看：[WORKFLOW.md](./WORKFLOW.md)。

目录分工：

- `tech_ops/`：技术人员操作（发布/打包）
- `user_ops/`：普通用户操作（安装/更新/运行）

说明：根目录不再放用户入口脚本，请只使用上述两个目录。

## 功能说明

脚本文件：`create_app.js`

当前实现会：

1. 连接本地已开启 CDP 的 Chrome（`http://localhost:9222`）。
2. 从配置文件 `developer_url.txt` 读取开发者控制台链接（不再写死开发者 ID）。
3. 从项目根目录的 Excel 读取 `应用名称`、`应用包名`（或英文同义列名）。
4. 逐行进入开发者控制台并点击 `Create app`。
5. 填写 `App name` 与 `App package name`，并点击 `Check availability`。
6. 随机选择 `App` 或 `Game`，选择 `Free`。
7. 自动勾选创建页声明项并提交创建。
8. 进入 `App content`，依次处理 7 类声明（Ads / App access / Target audience / Advertising ID / Government apps / Financial features / Health apps）。

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

### Windows VPS（给非技术用户）

本项目新增了三份双击脚本：

- `user_ops/install_windows.cmd`：首次安装/覆盖安装（会下载最新代码并自动初始化环境）
- `user_ops/update_windows.cmd`：后续更新（重新下载并覆盖）
- `user_ops/run_windows.cmd`：正式执行创建任务

它们内部会调用 `deploy_windows.ps1` + `bootstrap_windows.ps1`，自动完成：

- 下载你发布的 zip 包
- 自动检测/安装 Node.js 与 npm 依赖
- 自动检测 CDP 端口
- 可选自动启动浏览器调试模式（`9222`）

首次运行会弹出 UAC 管理员授权，请点允许。

发布包下载地址由 `release_url.txt` 控制。  
默认值是 GitHub `main.zip`，你可以改成你自己的文件服务器链接。

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

### 非技术用户（Windows）

1. 首次：双击 `user_ops/install_windows.cmd`
2. 在 `C:\app-create-data\developer_url.txt` 里粘贴你的 Play Console 开发者链接
3. 把 Excel 放到 `C:\app-create-data\apps.xlsx`
4. 双击 `user_ops/run_windows.cmd`
5. 后续更新：双击 `user_ops/update_windows.cmd`

### 技术用户（命令行）

在项目目录执行：

```bash
npm run start
```

或指定 Excel：

```bash
npm run start -- ./apps.xlsx
```

脚本在多次执行之间会随机等待 60-180 秒（用于降低频率）。

### 维护者发布流程（你本地）

1. 本地开发和测试，提交到 GitHub
2. 执行打包（macOS/Linux）：
```bash
./tech_ops/release_mac_linux.sh
```
Windows 打包可用：
```powershell
tech_ops\release_windows.cmd
```
3. 上传 `dist/app-create-latest.zip` 到你的文件服务器固定链接（例如 `https://your-domain/app-create-latest.zip`）
4. 在仓库里把 `release_url.txt` 改成这个固定链接，并再次发布
5. 通知用户在 VPS 双击 `user_ops/update_windows.cmd`

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
