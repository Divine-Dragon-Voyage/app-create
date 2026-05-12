# 启动器版（在线更新）使用说明

补充文档：

- 用户操作手册：`用户使用文档.md`
- 开发者维护手册：`开发者维护文档.md`

## 目标
给用户一个小体积启动器安装包：`AppCreateLauncherSetup-*.exe`。
用户双击后，启动器会检查更新并决定是否部署新版本，再启动自动化。

## 更新策略（当前版本）
1. 每次用户点击桌面 `App Create`，启动器都会检查远端包元数据。
2. 检测到新版本时会弹窗询问用户：
   - 是：立即更新并启动
   - 否：跳过本次更新，直接启动当前版本
   - 取消：本次退出，不启动
3. 检测不到新版本时，直接启动本地版本。

## 你每次发布新版本要做什么
1. 先更新主程序代码。
2. 生成主程序 zip 并上传到固定下载地址。
3. 下载地址不变时：用户不需要重装启动器。
4. 下载地址变化时：需要重打并重发启动器安装包。

## 你本地如何打“启动器安装包”
```powershell
npm run release:launcher:exe:win -- -InnoCompilerPath "D:\Inno Setup 6\ISCC.exe" -PackageUrl "https://your-domain.com/app-create/app-create-latest.zip"
```

输出目录：`dist\AppCreateLauncherSetup-<version>-<timestamp>.exe`

## 用户侧如何使用
1. 安装启动器安装包（会创建桌面快捷方式 `App Create`）。
2. 双击桌面 `App Create`。
3. 如果检测到新版本，按弹窗选择是否更新。

## 关键文件
- 启动器逻辑：`launcher/AppCreateLauncher.ps1`
- 启动入口：`launcher/AppCreateLauncher.cmd`
- 启动器配置：`launcher/release_url.txt`
- 在线部署脚本：`deploy_windows.ps1`
