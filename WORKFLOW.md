# 操作流程文档（技术人员 / 普通用户）

这份文档只讲“怎么做”，不讲原理。

目录分工：

- `tech_ops/`：技术人员操作
- `user_ops/`：普通用户操作

注意：根目录没有用户入口脚本，请只在对应目录里执行。

## 一、技术人员操作流程

### A. 首次准备（只做一次）

1. 本地拉取项目并安装依赖：
```bash
npm install
```
2. 修改 `release_url.txt`，写成你的固定下载链接（例如：`https://your-domain/app-create-latest.zip`）。

### B. 每次发布更新

1. 本地开发并测试。
2. 提交代码到 GitHub。
3. 打包发布文件（macOS/Linux）：
```bash
./tech_ops/release_mac_linux.sh
```
Windows 可执行：
```powershell
tech_ops\release_windows.cmd
```
4. 上传 `dist/app-create-latest.zip` 到你的固定下载链接。
5. 通知用户：在 VPS 双击 `user_ops/update_windows.cmd`。

### C. 你需要交付给用户的文件

直接把整个项目打包给用户即可。  
如果只给最小集，请至少包含：

- `user_ops/` 目录
- `deploy_windows.ps1`
- `bootstrap_windows.ps1`
- `create_app.js`
- `package.json`
- `package-lock.json`
- `developer_url.txt`
- `release_url.txt`

## 二、普通用户操作流程（Windows VPS）

### A. 首次安装

1. 解压技术人员发来的安装包。
2. 进入 `user_ops` 目录，双击 `install_windows.cmd`。
3. 看到管理员权限弹窗（UAC）时，点击“是/允许”。
4. 等待脚本执行完成（自动安装环境、自动下载代码、自动检查浏览器调试端口）。

### B. 配置开发者链接

1. 打开：
`C:\app-create-data\developer_url.txt`
2. 把里面的示例链接替换成你自己账号登录后的 Play Console 链接（包含 `/developers/你的ID`）。

### C. 放置 Excel

1. 把 Excel 文件命名为 `apps.xlsx`。
2. 放到：
`C:\app-create-data\apps.xlsx`

### D. 执行任务

1. 双击 `user_ops/run_windows.cmd`。
2. 保持浏览器已登录 Google 账号并可访问 Play Console。
3. 等待脚本跑完。

### E. 后续更新

1. 双击 `user_ops/update_windows.cmd`。
2. 更新完成后，再次双击 `user_ops/run_windows.cmd` 执行。

## 三、常见问题（简版）

1. 双击没反应：右键“以管理员身份运行”再试。
2. 提示 Excel 不存在：确认 `C:\app-create-data\apps.xlsx` 路径和文件名正确。
3. 提示 CDP 连接失败：先手动打开 HubVPS/Chrome 并登录 Play Console，再重试。
4. 更新失败：检查 VPS 网络是否能访问 `release_url.txt` 中的下载链接。
