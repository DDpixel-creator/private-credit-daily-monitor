import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path


def get_default_output_dir() -> Path:
    home = Path.home()
    documents = home / "Documents"
    if documents.exists() and documents.is_dir():
        base = documents
    else:
        base = home
    output_dir = base / "PrivateCreditDailyMonitor"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def run_update(skill_root: Path, output_dir: Path) -> subprocess.CompletedProcess:
    update_script = skill_root / "scripts" / "update_monitor.py"
    if not update_script.exists():
        raise FileNotFoundError(f"update_monitor.py not found: {update_script}")

    cmd = [sys.executable, str(update_script), "--workspace", str(output_dir)]
    return subprocess.run(cmd, capture_output=True, text=True)


def build_summary_text(
    run_time: str,
    success: bool,
    latest_file: Path,
    daily_file: Path,
    stdout: str,
    stderr: str,
) -> str:
    lines = []
    lines.append("Private Credit Daily Monitor 运行结果")
    lines.append(f"时间：{run_time}")
    lines.append(f"状态：{'成功' if success else '失败'}")
    lines.append(f"最新文件：{latest_file}")
    lines.append(f"归档文件：{daily_file}")

    if success:
        lines.append("说明：本次监控已完成，请打开 latest 文件查看最新结果。")
    else:
        lines.append("说明：本次运行失败，请检查错误输出。")
        if stderr.strip():
            lines.append(f"错误：{stderr.strip()[:400]}")
        elif stdout.strip():
            lines.append(f"输出：{stdout.strip()[:400]}")

    return "\n".join(lines)


def main() -> None:
    skill_root = Path(__file__).resolve().parents[1]
    output_dir = get_default_output_dir()

    run_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    date_str = datetime.now().strftime("%Y-%m-%d")

    latest_file = output_dir / "private_credit_monitor_latest.xlsx"
    daily_file = output_dir / f"private_credit_monitor_{date_str}.xlsx"

    result = run_update(skill_root, output_dir)
    success = result.returncode == 0

    summary = {
        "run_time": run_time,
        "status": "success" if success else "failed",
        "output_dir": str(output_dir),
        "latest_file": str(latest_file),
        "daily_file": str(daily_file),
        "stdout": result.stdout,
        "stderr": result.stderr,
    }

    summary_json = output_dir / "last_run_summary.json"
    summary_txt = output_dir / "last_run_summary.txt"

    summary_json.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    summary_txt.write_text(
        build_summary_text(
            run_time=run_time,
            success=success,
            latest_file=latest_file,
            daily_file=daily_file,
            stdout=result.stdout,
            stderr=result.stderr,
        ),
        encoding="utf-8",
    )

    print(f"output_dir: {output_dir}")
    print(f"summary_json: {summary_json}")
    print(f"summary_txt: {summary_txt}")
    print(result.stdout)
    if result.stderr.strip():
        print(result.stderr)

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
