# user_ops

请优先阅读根目录文档：[使用说明.md](../使用说明.md)

本目录给普通用户使用：

- `install_windows.cmd`：首次安装或覆盖安装
- `update_windows.cmd`：更新到新版本
- `run_windows.cmd`：启动任务（会弹出输入窗口）

说明：
- `run_windows.cmd` 现在会拉起可视化启动器，要求每次填写 Developer ID/URL 和数据文件路径（`.xlsx/.xls/.csv`）。
- 选择 `.csv` 时会自动生成同目录 `.__work.xlsx` 工作簿，并在该工作簿里维护断点续跑状态。
- 不再要求手工编辑固定的 `developer_url.txt` 或固定复制 `apps.xlsx` 到指定目录。
