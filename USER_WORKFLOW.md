# 普通用户操作流程（Windows VPS）

## 1. 首次安装

1. 先拿到技术人员给你的“初始安装包”并解压到 VPS（例如 `C:\app-create\`）。
2. 进入 `user_ops` 目录，双击 `install_windows.cmd`。
3. 看到管理员权限弹窗（UAC）时，点击“是/允许”。
4. 等待脚本执行完成（自动安装环境、自动下载代码、自动检查浏览器调试端口）。

## 2. 配置开发者链接

1. 打开：`C:\app-create-data\developer_url.txt`
2. 把里面的示例链接替换成你自己账号登录后的 Play Console 链接（包含 `/developers/你的ID`）。

## 3. 放置 Excel

1. 把 Excel 文件命名为 `apps.xlsx`。
2. 放到：`C:\app-create-data\apps.xlsx`

## 4. 执行任务

1. 双击 `user_ops/run_windows.cmd`。
2. 保持浏览器已登录 Google 账号并可访问 Play Console。
3. 等待脚本跑完。

## 5. 后续更新

1. 先确认 VPS 能访问技术人员配置的下载链接（`release_url.txt` 对应地址）。
2. 双击 `user_ops/update_windows.cmd`。
3. 更新完成后，再次双击 `user_ops/run_windows.cmd` 执行。

## 6. 常见问题

1. 双击没反应：右键“以管理员身份运行”再试。
2. 提示 Excel 不存在：确认 `C:\app-create-data\apps.xlsx` 路径和文件名正确。
3. 提示 CDP 连接失败：先手动打开 HubVPS/Chrome 并登录 Play Console，再重试。
4. 更新失败：检查 VPS 网络是否能访问 `release_url.txt` 中的下载链接。
