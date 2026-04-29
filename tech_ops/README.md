# tech_ops

技术人员只需要使用这个目录里的发布命令：

- Windows：`release_windows.cmd`
- macOS / Linux：`release_mac_linux.sh`

发布前请确认：

1. `release_url.txt` 已设置为你的固定下载链接
2. 代码已提交并测试通过
3. 执行发布后，把 `dist/app-create-latest.zip` 上传到该链接
