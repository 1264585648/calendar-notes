# Outlook 本地同步实现说明

## 架构

Calendar Notes 当前主路径使用经典 Outlook + COM 做只读同步。Rust 后端启动临时 PowerShell 脚本，通过当前 Windows 用户的经典 Outlook Object Model 读取默认日历快照，再写入 SQLite 日程缓存。

Microsoft Graph 相关后端代码仍保留为兼容能力，但页面不再展示账号连接、断开、Graph 登录等入口。

## 同步流程

1. 页面点击“刷新 Outlook 日程”或开启“自动同步 Outlook”。
2. 前端调用 `refresh_local_outlook()`。
3. Rust 后端把内置脚本 `src-tauri/scripts/outlook-com-sync.ps1` 写入临时目录，并以 UTF-8 BOM 形式执行，兼容 Windows PowerShell 5.1。
4. PowerShell 创建 `Outlook.Application` COM 对象，读取 `GetNamespace("MAPI").GetDefaultFolder(9)` 默认日历。
5. 脚本按当前月 ±1 月窗口执行 `Items.Sort("[Start]")`、`IncludeRecurrences = true`、`Restrict()`，输出 UTF-8 JSON。
6. Rust 将快照 upsert 到 `external_events`，并把本窗口内本次未返回的旧事件标记 `deleted_at`。
7. 页面重新加载月视图，将本地待办与 Outlook 只读日程合并展示。

## 自动同步策略

- 页面只保留一个“自动同步 Outlook”勾选项。
- 勾选后立即刷新一次，之后每 10 分钟自动刷新一次经典 Outlook 日程。
- 自动同步开关保存在浏览器 localStorage：`calendar-notes:auto-sync-classic-outlook`。
- 手动“刷新 Outlook 日程”按钮始终可用，会复用同一条经典 Outlook 同步命令。
- 后端不再无条件启动 Outlook 后台轮询，避免与页面开关冲突。

## Tauri Commands

- `refresh_local_outlook()`：刷新本机经典 Outlook 默认日历；如果本机来源尚未创建，会自动创建 `outlook-com` 账号缓存。
- `get_month_view(year, month)`：返回当前月 Outlook 只读日程和本地待办，前端按日期分组展示。
- `get_external_event_detail(event_id)`：返回 Outlook 日程完整详情。

兼容保留但页面不直接使用：

- `connect_outlook()`
- `connect_local_outlook()`
- `disconnect_outlook(account_id)`
- `sync_outlook_now(account_id?)`

## 数据存储

SQLite 文件位于 Tauri app data 目录下的 `calendar-notes.sqlite3`。核心表：

- `accounts`
- `external_calendars`
- `external_events`
- `sync_state`

`accounts.provider = 'outlook-com'` 表示本机经典 Outlook COM 快照账号。

经典 Outlook COM 事件标识优先使用 `GlobalAppointmentID`，重复日程实例会附加实例开始时间；没有全局 ID 时退化为 `EntryID`。

## 已知边界

- 只支持 Windows + 经典 Outlook，不支持新 Outlook。
- 需要在当前登录用户桌面会话中运行，不适合 Windows Service 或服务器无人值守任务。
- 默认隐藏私密日程详情；如需同步私密详情，可设置 `CALENDAR_NOTES_OUTLOOK_COM_INCLUDE_PRIVATE_DETAILS=true`。
- 当前不支持双向写回 Outlook。
