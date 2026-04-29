# 技术人员操作流程

## 1. 首次准备（只做一次）

1. 本地拉取项目并安装依赖：
```bash
npm install
```
2. 修改 `release_url.txt`，写成你的固定下载链接（例如：`https://your-domain/app-create-latest.zip`）。

## 2. 每次发布更新

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
5. 首次上线时，先把“初始安装包”（本项目目录或 zip）发到 VPS。
6. 后续更新时，通知用户在 VPS 双击 `user_ops/update_windows.cmd`。

## 3. 交付给用户的内容

建议：直接把整个项目目录打包给用户。  
如果只给最小集，请至少包含：

- `user_ops/` 目录
- `deploy_windows.ps1`
- `bootstrap_windows.ps1`
- `create_app.js`
- `package.json`
- `package-lock.json`
- `developer_url.txt`
- `release_url.txt`
