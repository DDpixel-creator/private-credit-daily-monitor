import argparse
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path


TASK_NAME = "PrivateCreditDailyMonitor"


def build_python_command(skill_root: Path) -> str:
    python_exe = sys.executable
    runner = skill_root / "scripts" / "run_daily_monitor.py"
    return f'"{python_exe}" "{runner}"'


def create_or_update_task(run_time: str, skill_root: Path) -> None:
    task_cmd = build_python_command(skill_root)

    # 先删除旧任务（如果存在）
    subprocess.run(
        ["schtasks", "/Delete", "/TN", TASK_NAME, "/F"],
        capture_output=True,
        text=True,
    )

    # 创建新任务
    create_cmd = [
        "schtasks",
        "/Create",
        "/SC", "DAILY",
        "/TN", TASK_NAME,
        "/TR", task_cmd,
        "/ST", run_time,
        "/F",
    ]

    cp = subprocess.run(create_cmd, capture_output=True, text=True)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr.strip() or cp.stdout.strip() or "Failed to create scheduled task")

    print(f"Task created/updated: {TASK_NAME}")
    print(f"Run time (local machine time): {run_time}")
    print(f"Command: {task_cmd}")
    print(f'Verify with: schtasks /Query /TN "{TASK_NAME}"')


def parse_time(value: str) -> str:
    try:
        dt = datetime.strptime(value, "%H:%M")
        return dt.strftime("%H:%M")
    except Exception:
        raise argparse.ArgumentTypeError("Time must be in HH:MM format")


def main() -> None:
    parser = argparse.ArgumentParser(description="Create/update Windows scheduled task for Private Credit Daily Monitor")
    parser.add_argument("--time", required=True, type=parse_time, help="Daily run time in HH:MM local time")
    parser.add_argument("--workspace", default=".", help="Reserved for compatibility; not used by the scheduler")
    args = parser.parse_args()

    skill_root = Path(__file__).resolve().parents[1]
    create_or_update_task(args.time, skill_root)


if __name__ == "__main__":
    main()
