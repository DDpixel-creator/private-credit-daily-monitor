import argparse
import json
import shutil
import subprocess
import tempfile
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


def build_single_line_message(summary_text: str) -> str:
    lines = [line.strip() for line in summary_text.splitlines() if line.strip()]
    body = " | ".join(lines)
    return f"【Private Credit Daily Monitor】 {body}"


def resolve_openclaw_executable() -> str:
    for name in ["openclaw", "openclaw.cmd", "openclaw.exe"]:
        path = shutil.which(name)
        if path:
            return path
    raise FileNotFoundError("openclaw executable not found in PATH")


def safe_text(value) -> str:
    return value if isinstance(value, str) else ""


def run_openclaw_send(channel: str, target: str, message: str) -> subprocess.CompletedProcess:
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

    env = dict()
    env.update(**subprocess.os.environ)
    env["PYTHONUTF8"] = "1"

    return subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
    )


def write_debug_message_file(message: str) -> Path:
    temp_dir = Path(tempfile.gettempdir())
    path = temp_dir / "private_credit_notify_message.txt"
    path.write_text(message, encoding="utf-8")
    return path


def send_message_via_openclaw(channel: str, target: str, message: str) -> subprocess.CompletedProcess:
    """
    Windows Scheduled Task 下，长多行中文参数可能被截断。
    为了提高稳定性，这里统一发送单行压缩版。
    完整多行正文仍会打印到日志，便于排查。
    """
    single_line_message = build_single_line_message(message.replace("【Private Credit Daily Monitor】\n", "", 1))
    return run_openclaw_send(channel, target, single_line_message)


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

    try:
        debug_path = write_debug_message_file(message)
        print(f"notification_debug_file: {debug_path}")
    except Exception as exc:
        print(f"notification_debug_file_error: {exc}")

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

            # 兜底：失败时尝试发送更短的极简单行版本
            fallback_message = "【Private Credit Daily Monitor】 通知发送失败，请查看本地输出目录：C:\\Users\\darry\\Documents\\PrivateCreditDailyMonitor"
            print("notify_fallback_attempt_start")
            cp2 = run_openclaw_send(channel, target, fallback_message)
            stdout_text2 = safe_text(cp2.stdout).strip()
            stderr_text2 = safe_text(cp2.stderr).strip()

            if cp2.returncode == 0:
                print("notify_fallback_status: sent")
            else:
                print("notify_fallback_status: failed")

            if stdout_text2:
                print("notify_fallback_stdout_start")
                print(stdout_text2)
                print("notify_fallback_stdout_end")
            if stderr_text2:
                print("notify_fallback_stderr_start")
                print(stderr_text2)
                print("notify_fallback_stderr_end")

    except FileNotFoundError as exc:
        print(f"notify_status: error_openclaw_not_found ({exc})")
    except Exception as exc:
        print(f"notify_status: error ({exc})")


if __name__ == "__main__":
    main()
