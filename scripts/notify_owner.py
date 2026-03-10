import argparse
import json
import shutil
import subprocess
from pathlib import Path


def get_default_output_dir() -> Path:
    home = Path.home()
    documents = home / "Documents"
    base = documents if documents.exists() and documents.is_dir() else home
    output_dir = base / "PrivateCreditDailyMonitor"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def read_summary_text(output_dir: Path) -> str:
    summary_txt = output_dir / "last_run_summary.txt"
    if not summary_txt.exists():
        raise FileNotFoundError(f"Summary file not found: {summary_txt}")
    return summary_txt.read_text(encoding="utf-8").strip()


def read_notify_config(skill_root: Path) -> dict:
    config_path = skill_root / "notify_config.json"
    if not config_path.exists():
        return {}
    try:
        return json.loads(config_path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def build_notification_message(summary_text: str) -> str:
    return "【Private Credit Daily Monitor】\n" + summary_text


def resolve_openclaw_executable() -> str:
    for name in ["openclaw", "openclaw.cmd", "openclaw.exe"]:
        path = shutil.which(name)
        if path:
            return path
    raise FileNotFoundError("openclaw executable not found in PATH")


def safe_text(value) -> str:
    return value if isinstance(value, str) else ""


def send_message_via_openclaw(channel: str, target: str, message: str) -> subprocess.CompletedProcess:
    openclaw_exe = resolve_openclaw_executable()
    cmd = [
        openclaw_exe,
        "message",
        "send",
        "--channel",
        channel,
        "--target",
        target,
        "--message",
        message,
    ]
    return subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Send owner notification from last_run_summary.txt")
    parser.add_argument("--output-dir", default="", help="Optional explicit output directory")
    args = parser.parse_args()

    skill_root = Path(__file__).resolve().parents[1]
    output_dir = Path(args.output_dir).resolve() if args.output_dir else get_default_output_dir()

    summary_text = read_summary_text(output_dir)
    message = build_notification_message(summary_text)

    print("notification_message_start")
    print(message)
    print("notification_message_end")

    cfg = read_notify_config(skill_root)
    channel = str(cfg.get("notify_channel", "")).strip()
    target = str(cfg.get("notify_target", "")).strip()

    if not channel or not target:
        print("notify_status: skipped (missing notify_channel/notify_target in notify_config.json)")
        return

    try:
        cp = send_message_via_openclaw(channel, target, message)
        stdout_text = safe_text(cp.stdout).strip()
        stderr_text = safe_text(cp.stderr).strip()

        if cp.returncode == 0:
            print("notify_status: sent")
            if stdout_text:
                print("notify_send_stdout_start")
                print(stdout_text)
                print("notify_send_stdout_end")
        else:
            print("notify_status: failed")
            if stdout_text:
                print("notify_send_stdout_start")
                print(stdout_text)
                print("notify_send_stdout_end")
            if stderr_text:
                print("notify_send_stderr_start")
                print(stderr_text)
                print("notify_send_stderr_end")

    except FileNotFoundError as exc:
        print(f"notify_status: error_openclaw_not_found ({exc})")
    except Exception as exc:
        print(f"notify_status: error ({exc})")


if __name__ == "__main__":
    main()
