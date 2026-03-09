import argparse
import platform
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Tuple

JOB_NAME = "PrivateCreditDailyMonitor"
DEFAULT_TIME = "09:00"


def run_cmd(args: list[str], check: bool = True) -> subprocess.CompletedProcess:
    cp = subprocess.run(args, capture_output=True, text=True)
    if check and cp.returncode != 0:
        raise RuntimeError(cp.stderr.strip() or cp.stdout.strip() or f"Command failed: {args}")
    return cp


def parse_time(value: str) -> Tuple[str, str]:
    parts = value.split(":")
    if len(parts) != 2:
        raise ValueError("Time must be in HH:MM format, e.g. 09:00")

    hour, minute = parts
    if not (hour.isdigit() and minute.isdigit()):
        raise ValueError("Time must be numeric, e.g. 09:00")

    hh = int(hour)
    mm = int(minute)
    if not (0 <= hh <= 23 and 0 <= mm <= 59):
        raise ValueError("Time out of range")

    return f"{hh:02d}", f"{mm:02d}"


def setup_windows(workspace: Path, run_time: str) -> None:
    python_exe = Path(sys.executable).resolve()
    script_path = (Path(__file__).resolve().parents[0] / "update_monitor.py").resolve()

    task_cmd = f'"{python_exe}" "{script_path}" --workspace "{workspace}"'
    hh, mm = parse_time(run_time)
    st = f"{hh}:{mm}"

    args = [
        "schtasks",
        "/Create",
        "/TN",
        JOB_NAME,
        "/SC",
        "DAILY",
        "/ST",
        st,
        "/TR",
        task_cmd,
        "/F",
    ]
    run_cmd(args)

    print(f"Task created/updated: {JOB_NAME}")
    print(f"Run time (local machine time): {st}")
    print(f'Verify with: schtasks /Query /TN "{JOB_NAME}"')


def setup_posix(workspace: Path, run_time: str) -> None:
    python_exe = Path(sys.executable).resolve()
    script_path = (Path(__file__).resolve().parents[0] / "update_monitor.py").resolve()

    hh, mm = parse_time(run_time)
    cron_line = f'{mm} {hh} * * * "{python_exe}" "{script_path}" --workspace "{workspace}"\n'
    marker = f"# {JOB_NAME}"

    existing = ""
    cp = subprocess.run(["crontab", "-l"], capture_output=True, text=True)
    if cp.returncode == 0:
        existing = cp.stdout
    elif "no crontab" in (cp.stderr or "").lower():
        existing = ""
    else:
        raise RuntimeError(cp.stderr.strip() or cp.stdout.strip())

    lines = existing.splitlines()
    filtered: list[str] = []
    skip_next = False

    for line in lines:
        if line.strip() == marker:
            skip_next = True
            continue
        if skip_next:
            skip_next = False
            continue
        filtered.append(line)

    filtered.append(marker)
    filtered.append(cron_line.rstrip("\n"))
    new_content = "\n".join(filtered).strip() + "\n"

    cp_set = subprocess.run(["crontab", "-"], input=new_content, capture_output=True, text=True)
    if cp_set.returncode != 0:
        raise RuntimeError(cp_set.stderr.strip() or cp_set.stdout.strip())

    print(f"Cron entry created/updated: {JOB_NAME}")
    print(f"Run time (local machine time): {hh}:{mm}")
    print("Verify with: crontab -l")


def main() -> None:
    parser = argparse.ArgumentParser(description="Create or update daily system scheduler for private credit monitor")
    parser.add_argument("--workspace", default=".", help="Workspace directory for output files")
    parser.add_argument("--time", default=DEFAULT_TIME, help="Daily run time in HH:MM, local machine time")
    args = parser.parse_args()

    workspace = Path(args.workspace).resolve()
    workspace.mkdir(parents=True, exist_ok=True)

    system_name = platform.system().lower()

    try:
        if system_name == "windows":
            if not shutil.which("schtasks"):
                raise RuntimeError("schtasks not found")
            setup_windows(workspace, args.time)
        else:
            if not shutil.which("crontab"):
                raise RuntimeError("crontab not found")
            setup_posix(workspace, args.time)
    except Exception as exc:
        print(f"Failed to set up scheduler: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
