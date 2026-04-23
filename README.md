# app-create

使用 Playwright 自动在 Google Play Console 创建应用并完成常见声明流程。

## 功能说明

脚本文件：`create_app.js`

当前实现会：

1. 连接本地已开启 CDP 的 Chrome（`http://localhost:9222`）。
2. 进入指定开发者控制台的应用列表页。
3. 点击 `Create app` 并填写应用名（时间戳）。
4. 随机选择 `App` 或 `Game`，选择 `Free`。
5. 自动勾选创建页声明项并提交创建。
6. 进入 `App content`，依次处理 7 类声明（Ads / App access / Target audience / Advertising ID / Government apps / Financial features / Health apps）。

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

### 3) 启动带远程调试端口的 Chrome

macOS 示例：

```bash
open -na "Google Chrome" --args --remote-debugging-port=9222 --user-data-dir=/tmp/chrome-cdp-app-create
```

然后在这个浏览器里手动登录 Google 账号与 Play Console。

## 使用方式

在项目目录执行：

```bash
npm run start
```

默认执行 1 次。指定执行次数：

```bash
npm run start -- 5
```

脚本在多次执行之间会随机等待 60-180 秒。

也可以直接使用预设脚本：

```bash
npm run start:once
npm run start:5
```

## 关键行为与默认选项

- 应用名：按当前时间生成（`YYYYMMDDHHmmss`）
- 类型：`App` / `Game` 随机
- 收费方式：`Free`
- 声明页中多数选项偏向“无相关功能”或最小功能路径

请根据你的真实业务自行调整 `create_app.js` 中各声明步骤，避免与应用实际功能不一致。

## 常见问题

### `Cannot find module 'playwright'`

未安装依赖，执行：

```bash
npm install
```

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
- 若用于批量操作，建议增加日志持久化和失败截图。
