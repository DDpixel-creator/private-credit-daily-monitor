# Private Credit Daily Monitor

## 作者信息
- Name: DD
- X: [@DD29397053](https://x.com/DD29397053)
- 欢迎关注作者 X 账号，会不定期发布投资类 skill

## 这是什么
这是一个用于每日更新私募信贷监控表的本地可运行 skill 仓库。

它会：
- 抓取公开市场与信用数据
- 在 Excel 中更新状态灯
- 写入“状态依据（具体数据+来源）”
- 输出最新文件和按日期归档文件
- 支持本地自动调度运行

## 自动执行方式
本版本优先采用系统级调度器，目标是优先保证真正可运行：

- Windows：Task Scheduler
- macOS / Linux：cron

## 首次使用步骤
1. 运行环境检查：
   `python scripts/self_check.py`

2. 创建自动任务：
   `python scripts/setup_scheduler.py --workspace .`

3. 手动试跑一次更新：
   `python scripts/update_monitor.py --workspace .`

## 输出文件
执行成功后，应生成或更新：

- `private_credit_monitor_master_template.xlsx`
- `private_credit_monitor_latest.xlsx`
- `private_credit_monitor_YYYY-MM-DD.xlsx`

## 仓库结构
- `SKILL.md`：skill 说明
- `install.json`：安装识别信息
- `scripts/self_check.py`：环境检查
- `scripts/setup_scheduler.py`：创建自动任务
- `scripts/update_monitor.py`：更新 Excel

## 说明
如果 `assets/private_credit_monitor_template.xlsx` 不存在，
`update_monitor.py` 会自动生成一个最小可用模板。

## 免责声明
这不是投资建议。
公开数据可能有延迟、修订或临时不可用。
