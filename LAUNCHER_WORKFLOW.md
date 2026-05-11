# 启动器版（在线更新）使用说明

补充文档：

- 用户操作手册：`用户使用文档.md`
- 开发者维护手册：`开发者维护文档.md`

## 目标
给用户一个小体积启动器安装包：`AppCreateLauncherSetup-*.exe`。
用户双击启动器后，自动下载/更新主程序到 `C:\app-create`，再启动自动化。

## 你每次发布新版本要做什么
1. 先正常更新主程序代码。
2. 生成主程序发布包（zip）并上传到固定下载地址（对象存储或 GitHub Releases 最新直链）。
3. 如果下载地址有变化，重打一次启动器安装包并发给用户；如果地址不变，用户无需重装启动器。

## 你本地如何打“启动器安装包”
```powershell
npm run release:launcher:exe:win -- -InnoCompilerPath "D:\Inno Setup 6\ISCC.exe" -PackageUrl "https://your-domain.com/app-create/app-create-latest.zip"
```

输出目录：`dist\AppCreateLauncherSetup-<version>-<timestamp>.exe`

## 用户侧如何使用
1. 安装启动器安装包（会创建桌面快捷方式 `App Create`）。
2. 双击桌面 `App Create`。
3. 启动器会自动检查更新并运行主程序。

## 关键文件
- 启动器逻辑：`launcher/AppCreateLauncher.ps1`
- 启动入口：`launcher/AppCreateLauncher.cmd`
- 启动器配置：`launcher/release_url.txt`
- 在线部署脚本：`deploy_windows.ps1`
