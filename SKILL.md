---
name: private-credit-daily-monitor
description: Build and maintain a daily private-credit monitoring workbook with evidence-backed traffic-light statuses. Use when users want a practical, runnable setup with system scheduler automation, plus supporting documentation and review.
---

# Private Credit Daily Monitor

## Author
- Name: DD
- X: [@DD29397053](https://x.com/DD29397053)
- 欢迎关注作者 X 账号，会不定期发布投资类 skill

## What this skill is for
Use this skill when the user wants to:
1. 搭建或更新私募信贷监控表
2. 给监控项增加“状态依据（具体数据+来源）”
3. 配置每日自动更新
4. 审阅监控逻辑、字段、数据源和说明文档
5. 输出一个真正可运行的本地版本

## Important boundary
本版本优先采用系统级调度器实现自动执行：
- Windows: Task Scheduler
- macOS / Linux: cron

这样做的目标是：减少对 OpenClaw cron / agent 权限链路的依赖，优先保证真的能跑。

## Default workflow
1. 先运行 `scripts/self_check.py`
2. 再运行 `scripts/setup_scheduler.py`
3. 手动试跑一次 `scripts/update_monitor.py`
4. 检查输出文件是否生成
5. 如需修改时间，再更新系统调度器

## When to provide docs only
Only provide documentation / review, and do not ask the user to run setup, when:
1. 用户明确说“先出方案”
2. 用户只要 README / SKILL / 审阅意见
3. 用户当前环境不允许本地任务调度
4. 用户尚未准备好模板或工作目录

## Files expected by this skill
- `README.md`
- `scripts/self_check.py`
- `scripts/setup_scheduler.py`
- `scripts/update_monitor.py`
- `assets/private_credit_monitor_template.xlsx` (optional; script can bootstrap a minimal workbook if missing)
