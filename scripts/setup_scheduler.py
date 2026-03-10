import argparse
import json
import subprocess
import sys
from pathlib import Path


TASK_NAME = "PrivateCreditDailyMonitor"


def build_python_command(skill_root: Path) -> str:
    python_exe = sys.executable
    runner = skill_root / "scripts" / "run_daily_monitor.py"
    return f'"{python_exe}" "{runner}"'


def write_notify_config(skill_root: Path, notify_channel: str, notify_target: str) -> Path:
    config_path = skill_root / "notify_config.json"
    config = {
        "notify_channel": notify_channel or "",
        "notify_target": notify_target or "",
    }
    config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    return config_path


def create_or_update_task(run_time: str, skill_root: Path) -> None:
    task_cmd = build_python_command(skill_root)

    subprocess.run(
        ["schtasks", "/Delete", "/TN", TASK_NAME, "/F"],
        capture_output=True,
        text=True,
    )

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
        from datetime import datetime
        dt = datetime.strptime(value, "%H:%M")
        return dt.strftime("%H:%M")
    except Exception:
        raise argparse.ArgumentTypeError("Time must be in HH:MM format")


def main() -> None:
    parser = argparse.ArgumentParser(description="Create/update Windows scheduled task for Private Credit Daily Monitor")
    parser.add_argument("--time", required=True, type=parse_time, help="Daily run time in HH:MM local time")
    parser.add_argument("--workspace", default=".", help="Reserved for compatibility; not used by the scheduler")
    parser.add_argument("--notify-channel", default="", help="Optional OpenClaw notify channel, e.g. telegram/slack/discord")
    parser.add_argument("--notify-target", default="", help="Optional OpenClaw notify target, e.g. chat id / channel id / username")
    args = parser.parse_args()

    skill_root = Path(__file__).resolve().parents[1]

    if args.notify_channel or args.notify_target:
        config_path = write_notify_config(skill_root, args.notify_channel, args.notify_target)
        print(f"Notify config written: {config_path}")

    create_or_update_task(args.time, skill_root)


if __name__ == "__main__":
    main()
