# Private Credit Daily Monitor

## 作者信息
- Name: DD
- X: [@DD29397053](https://x.com/DD29397053)
- 欢迎关注作者 X 账号，会不定期发布投资类 Skill

## 这个 Skill 是做什么的
这是一个用于监控**私募信贷风险变化**的自动化监控面板。  
它会基于公开市场数据和主流媒体新闻证据，自动更新：

- Dashboard 汇总页
- Checklist 16 项监控项
- Evidence 新闻证据页

当前版本支持：

- 手动运行更新
- Windows 自动定时运行
- GitHub 仓库安装 / 更新
- 运行完成后生成文字摘要与结构化结果
- 可选通过 OpenClaw 当前配置的聊天通道主动发送结果通知

---

## 安装方式（推荐）
如果你的 OpenClaw agent 支持 GitHub 仓库安装，可以直接把这个仓库链接发给它，让它安装：

```text
https://github.com/DDpixel-creator/private-credit-daily-monitor
```

推荐让 agent 按下面要求执行：

```text
请从这个 GitHub 仓库安装这个 skill：
https://github.com/DDpixel-creator/private-credit-daily-monitor

要求：
1. 安装到当前 workspace 的 skills 目录
2. 安装后刷新 skills 或重启 gateway
3. 然后执行：
   openclaw skills info private-credit-daily-monitor
4. 不要先执行 setup_scheduler.py
5. 不要先执行 run_daily_monitor.py
6. 先告诉我 skill 是否已经成功加载
```

---

## 首次初始化
安装成功后，建议按下面顺序运行：

```bash
python scripts/setup_scheduler.py --time 09:00
python scripts/run_daily_monitor.py
```

如果你希望每日任务跑完后主动通过聊天通道通知你，可以在初始化时同时配置通知参数：

```bash
python scripts/setup_scheduler.py --time 09:00 --notify-channel <channel> --notify-target <target>
```

例如：

```bash
python scripts/setup_scheduler.py --time 09:00 --notify-channel telegram --notify-target 123456789
```

说明：

- `setup_scheduler.py`：创建或更新 Windows 每日自动任务
- `run_daily_monitor.py`：立即运行一次完整流程（更新表格 + 生成摘要 + 生成通知正文）
- 如果配置了通知参数，任务跑完后会尝试调用 OpenClaw 官方 CLI 主动发消息给主人

---

## 自动执行主流程
每日自动任务实际执行的是：

```bash
python scripts/run_daily_monitor.py
```

它会自动完成：

1. 运行私募信贷监控
2. 生成最新 Excel 结果
3. 生成每日归档文件
4. 生成 `last_run_summary.txt`
5. 生成 `last_run_summary.json`
6. 如果配置了通知参数，则尝试通过 OpenClaw 聊天通道主动发送通知给主人

---

## 结果保存位置
监控结果默认保存到当前用户自己的文档目录：

```text
Documents/PrivateCreditDailyMonitor/
```

主要输出文件包括：

- `private_credit_monitor_latest.xlsx`
- `private_credit_monitor_YYYY-MM-DD.xlsx`
- `last_run_summary.txt`
- `last_run_summary.json`

其中：

- `latest.xlsx`：最新结果
- `YYYY-MM-DD.xlsx`：每日归档
- `last_run_summary.txt`：适合直接阅读或转发的文字摘要
- `last_run_summary.json`：适合程序读取的结构化结果

---

## 主要输出内容
生成结果文件后，主要看这 3 个 sheet：

- `Dashboard`：总览面板
- `Checklist`：16 项监控项结果
- `Evidence`：新闻证据和链接

同时还会生成：

- `Lists`
- `AutomationMap`

用于规则映射与维护说明。

---

## Dashboard 设计
Dashboard 采用 5 档整体等级：

- 绿：稳定
- 浅黄：预警升温
- 黄：压力上升
- 橙：系统风险临近
- 红：危机确认

Dashboard 会自动写入：

- 更新日期
- 指标总数
- 绿灯数
- 黄灯数
- 红灯数
- 待更新数
- 加权分数
- 当前整体等级
- 等级释义
- 300 字内自动摘要

---

## 当前版本说明
当前版本支持：

- 16 项监控项全部自动输出状态
- 数值型指标使用公开市场数据
- 事件型指标使用主流媒体新闻聚类
- Evidence sheet 会写入新闻标题、来源、日期和链接，方便用户自行核验是否误报
- 自动输出 5 档整体风险等级
- 自动输出 300 字内摘要
- 自动保存到当前用户自己的文档目录
- 自动生成 summary 文件
- 可选通过 OpenClaw 聊天通道主动通知用户

---

## 16 项监控项
当前监控项分为三层：

### 早期预警层（EW）
- EW-01
- EW-02
- EW-03
- EW-04
- EW-05

### 传导层（TR）
- TR-01
- TR-02
- TR-03
- TR-04
- TR-05
- TR-06

### 系统确认层（SY）
- SY-01
- SY-02
- SY-03
- SY-04
- SY-05

整体等级不是简单按红灯数量决定，而是综合：

- 红灯数量
- 红灯所在层级
- 是否触发关键系统确认项

---

## 通知机制
本 Skill 支持在初始化时设置通知通道与通知目标。

例如：

```bash
python scripts/setup_scheduler.py --time 09:00 --notify-channel telegram --notify-target 123456789
```

当前通知发送依赖 OpenClaw 官方 CLI：

```bash
openclaw message send --channel <channel> --target <dest> --message "..."
```

说明：

- 如果 `notify_channel` 和 `notify_target` 配置正确，任务跑完后会尝试主动发送结果通知给主人
- 如果未配置通知参数，任务仍会正常运行，但只会在结果目录中生成：
  - `last_run_summary.txt`
  - `last_run_summary.json`

注意：

- 是否真正送达，取决于：
  - `notify_channel` 是否配置正确
  - `notify_target` 是否配置正确
  - 对应聊天通道是否已连通
  - 目标 chat/channel/user 是否有效

---

## 暂停、恢复与卸载

### 暂停监控
如果你暂时不想继续每日自动运行，可以直接对 OpenClaw 说：

- 暂停私募信贷监控
- 停止私募信贷日报
- 先停掉每日监控

对应脚本：

```bash
python scripts/disable_monitor.py
```

作用：

- 删除每日计划任务
- 停止后续自动监控
- 保留历史结果文件
- 保留通知配置

---

### 恢复监控
如果你已经暂停，想重新开启每日自动监控，可以直接对 OpenClaw 说：

- 恢复私募信贷监控
- 重新开启私募信贷日报
- 每天 09:00 恢复私募信贷监控

对应脚本：

```bash
python scripts/setup_scheduler.py --time 09:00
```

如果你还想同时恢复通知：

```bash
python scripts/setup_scheduler.py --time 09:00 --notify-channel <channel> --notify-target <target>
```

例如：

```bash
python scripts/setup_scheduler.py --time 09:00 --notify-channel telegram --notify-target 123456789
```

---

### 卸载监控
如果你不再需要这个监控，可以直接对 OpenClaw 说：

- 卸载私募信贷监控
- 删除这个监控
- 彻底停掉并卸载

对应脚本：

```bash
python scripts/uninstall_monitor.py
```

这会：

- 删除每日计划任务
- 保留历史结果目录
- 保留通知配置（默认）

如果你希望连结果文件和通知配置一起清除：

```bash
python scripts/uninstall_monitor.py --delete-results --clear-notify-config
```

---

## 推荐使用方式

### 初始化一次
```bash
python scripts/setup_scheduler.py --time 09:00 --notify-channel <channel> --notify-target <target>
```

### 立即手动运行一次
```bash
python scripts/run_daily_monitor.py
```

### 暂停
```bash
python scripts/disable_monitor.py
```

### 恢复
```bash
python scripts/setup_scheduler.py --time 09:00 --notify-channel <channel> --notify-target <target>
```

### 卸载
```bash
python scripts/uninstall_monitor.py
```

---

## 自然语言使用方式
用户不需要记命令，可以直接对 OpenClaw 说：

- 现在运行一次私募信贷监控
- 暂停私募信贷监控
- 恢复私募信贷监控，每天 09:00
- 查看私募信贷监控结果
- 给我最近一次私募信贷摘要
- 卸载私募信贷监控

Skill 会根据意图自动执行对应脚本。

---

## 常见问题

### 1. 为什么运行时报 PermissionError
如果 `private_credit_monitor_latest.xlsx` 或归档文件正被 Excel 打开，Windows 可能拒绝覆盖写入。  
请先关闭对应文件，再重新运行。

### 2. 为什么任务跑完了但没收到通知
常见原因包括：

- `notify_channel` 未配置
- `notify_target` 未配置
- 通道类型不对
- 目标 chat/channel/user 无效
- 对应聊天通道未连通

此时可先检查：

- `notify_config.json`
- `last_run_summary.txt`
- `last_run_summary.json`

### 3. 如果不配置通知会怎样
监控仍会正常运行，只是不会主动发消息给主人。  
结果仍会保存在：

```text
Documents/PrivateCreditDailyMonitor/
```

### 4. 如果暂停后想恢复怎么办
重新运行：

```bash
python scripts/setup_scheduler.py --time 09:00 --notify-channel <channel> --notify-target <target>
```

---

## 免责声明
本 Skill 仅用于监控和研究参考，不构成任何投资建议。  
公开数据和新闻源可能存在延迟、修订、误报或口径差异，使用时请结合原始来源自行核验。
