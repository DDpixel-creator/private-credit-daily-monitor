import argparse
import json
import shutil
import subprocess
from pathlib import Path

TASK_NAME = "PrivateCreditDailyMonitor"


def get_default_output_dir() -> Path:
    home = Path.home()
    documents = home / "Documents"
    base = documents if documents.exists() and documents.is_dir() else home
    return base / "PrivateCreditDailyMonitor"


def remove_task() -> tuple[bool, str, str]:
    cp = subprocess.run(
        ["schtasks", "/Delete", "/TN", TASK_NAME, "/F"],
        capture_output=True,
        text=True,
    )
    return cp.returncode == 0, cp.stdout or "", cp.stderr or ""


def clear_notify_config(skill_root: Path) -> bool:
    config_path = skill_root / "notify_config.json"
    if not config_path.exists():
        return False
    config_path.write_text(
        json.dumps({"notify_channel": "", "notify_target": ""}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return True


def remove_output_dir(output_dir: Path) -> bool:
    if not output_dir.exists():
        return False
    shutil.rmtree(output_dir)
    return True


def main() -> None:
    parser = argparse.ArgumentParser(description="Uninstall Private Credit Daily Monitor")
    parser.add_argument(
        "--delete-results",
        action="store_true",
        help="Also delete Documents/PrivateCreditDailyMonitor output folder",
    )
    parser.add_argument(
        "--clear-notify-config",
        action="store_true",
        help="Also clear notify_config.json",
    )
    args = parser.parse_args()

    skill_root = Path(__file__).resolve().parents[1]
    output_dir = get_default_output_dir()

    removed, stdout_text, stderr_text = remove_task()

    if removed:
        print(f"Scheduled task removed: {TASK_NAME}")
    else:
        lower_out = stdout_text.lower()
        lower_err = stderr_text.lower()
        if "cannot find the file specified" in lower_out or "cannot find the file specified" in lower_err or "找不到指定的文件" in stdout_text or "找不到指定的文件" in stderr_text:
            print(f"Scheduled task already absent: {TASK_NAME}")
        else:
            print(f"Failed to remove scheduled task: {TASK_NAME}")
            if stdout_text.strip():
                print("stdout_start")
                print(stdout_text.strip())
                print("stdout_end")
            if stderr_text.strip():
                print("stderr_start")
                print(stderr_text.strip())
                print("stderr_end")

    if args.clear_notify_config:
        cleared = clear_notify_config(skill_root)
        print("Notify config cleared." if cleared else "Notify config not found; skipped.")

    if args.delete_results:
        deleted = remove_output_dir(output_dir)
        print(f"Output directory removed: {output_dir}" if deleted else f"Output directory not found; skipped: {output_dir}")
    else:
        print(f"Output directory preserved: {output_dir}")

    print("Uninstall completed.")
    print("If you also want to remove the skill itself, delete the skill folder from your OpenClaw workspace manually.")


if __name__ == "__main__":
    main()
