---
name: private-credit-daily-monitor
description: Monitor private credit risk daily, generate Excel dashboard/results, save outputs to the user's Documents folder, and support natural-language actions such as run now, pause monitoring, resume monitoring, view results location, and uninstall monitoring.
---

# Private Credit Daily Monitor

## 作者
- Name: DD
- X: [@DD29397053](https://x.com/DD29397053)
- 欢迎关注作者 X 账号，会不定期发布投资类 Skill

## 这个 Skill 做什么
这个 Skill 用于自动监控私募信贷风险，生成 Excel 结果文件，并支持以下自然语言操作：

- 立即运行一次监控
- 启动 / 恢复每日自动监控
- 暂停每日自动监控
- 查看结果保存位置
- 查看最近一次监控摘要
- 卸载监控

## 结果保存位置
监控结果默认保存到当前用户自己的文档目录：

- Windows: `Documents/PrivateCreditDailyMonitor/`

主要输出文件：

- `private_credit_monitor_latest.xlsx`
- `private_credit_monitor_YYYY-MM-DD.xlsx`
- `last_run_summary.json`
- `last_run_summary.txt`

## 触发条件
当用户表达以下意图时，应该使用本 Skill：

### 1. 立即运行一次
用户可能会说：
- 现在运行一次私募信贷监控
- 立刻更新私募信贷监控
- 刷新私募信贷日报
- 重新生成私募信贷结果

对应动作：
- 运行：
  `python scripts/run_daily_monitor.py`

执行后应告诉用户：
- 是否运行成功
- 最新文件保存位置
- 最近摘要内容

---

### 2. 启动 / 恢复每日监控
用户可能会说：
- 开启私募信贷监控
- 恢复私募信贷监控
- 每天 09:00 跑私募信贷监控
- 重新启用私募信贷日报
- 每天自动监控私募信贷，并通知我

对应动作：
- 运行：
  `python scripts/setup_scheduler.py --time <HH:MM>`

如果用户同时提供通知通道和目标，则运行：
- `python scripts/setup_scheduler.py --time <HH:MM> --notify-channel <channel> --notify-target <target>`

其中 `<channel>` 可能是：
- telegram
- slack
- discord
- whatsapp
- signal
- imessage
- 其他 OpenClaw 支持的通道

执行后应告诉用户：
- 每日执行时间
- 是否已启用通知
- 结果保存位置

---

### 3. 暂停监控
用户可能会说：
- 暂停私募信贷监控
- 停止私募信贷日报
- 先停掉每日监控
- 不要再每天跑了

对应动作：
- 运行：
  `python scripts/disable_monitor.py`

执行后应告诉用户：
- 每日计划任务已停止
- 历史结果文件仍保留
- 以后可通过 setup_scheduler.py 恢复

---

### 4. 查看结果位置 / 最近结果
用户可能会说：
- 私募信贷结果在哪里
- 打开最近一次私募信贷监控结果
- 最新表格保存在哪
- 给我最近一次监控摘要
- 查看私募信贷 summary

对应动作：
- 读取结果目录：
  `Documents/PrivateCreditDailyMonitor/`
- 优先告诉用户：
  - latest 文件位置
  - daily 文件位置
  - last_run_summary.txt 内容

不需要重新运行，除非用户明确要求“现在再跑一次”。

---

### 5. 卸载监控
用户可能会说：
- 卸载私募信贷监控
- 删除这个监控
- 我不要这个私募信贷日报了
- 彻底停掉并卸载

对应动作：
- 如果存在 `scripts/uninstall_monitor.py`，运行：
  `python scripts/uninstall_monitor.py`
- 如果不存在卸载脚本，则至少先运行：
  `python scripts/disable_monitor.py`
并明确告诉用户还需要手动删除哪些文件

---

## 默认执行顺序
### A. 用户要求“现在运行一次”
1. 运行 `python scripts/run_daily_monitor.py`
2. 返回：
   - 是否成功
   - 整体等级
   - 绿/黄/红/待更新数
   - latest 文件路径
   - 摘要

### B. 用户要求“开启/恢复监控”
1. 判断用户是否提供执行时间
2. 判断用户是否提供 notify_channel / notify_target
3. 运行 `setup_scheduler.py`
4. 返回：
   - 定时任务已建立
   - 每天几点执行
   - 是否启用通知
   - 结果保存目录

### C. 用户要求“暂停监控”
1. 运行 `disable_monitor.py`
2. 返回暂停结果

### D. 用户要求“查看结果”
1. 读取 `last_run_summary.txt`
2. 告诉用户 latest 文件路径和摘要

## 重要边界
- 不要要求用户记住命令行
- 优先让用户通过自然语言操作
- 如果需要时间、通知通道、通知目标等关键信息，而用户没提供，应先简短追问
- 如果通知配置缺失，不要假装已经能通知用户
- 如果 Excel 文件被用户打开导致无法覆盖，应明确提示用户先关闭文件再重试

## 通知机制
本 Skill 支持通过 OpenClaw 官方 CLI 主动发消息：

- `openclaw message send --channel <channel> --target <dest> --message "..."`
  
但是否真正送达，取决于：
- notify_channel 是否配置正确
- notify_target 是否配置正确
- 对应聊天通道是否已连通
- 目标 chat/channel/user 是否有效

## 结果解释
整体等级采用 5 档：

- 绿：稳定
- 浅黄：预警升温
- 黄：压力上升
- 橙：系统风险临近
- 红：危机确认

如果用户询问“现在风险多高”，优先返回：
- 当前整体等级
- 绿/黄/红/待更新数
- 100～300字摘要
