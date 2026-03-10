import ast
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path


def get_default_output_dir() -> Path:
    home = Path.home()
    documents = home / "Documents"
    base = documents if documents.exists() and documents.is_dir() else home
    output_dir = base / "PrivateCreditDailyMonitor"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def run_update(skill_root: Path, output_dir: Path) -> subprocess.CompletedProcess:
    update_script = skill_root / "scripts" / "update_monitor.py"
    if not update_script.exists():
        raise FileNotFoundError(f"update_monitor.py not found: {update_script}")

    cmd = [sys.executable, str(update_script), "--workspace", str(output_dir)]
    return subprocess.run(cmd, capture_output=True, text=True)


def run_notify(skill_root: Path, output_dir: Path) -> subprocess.CompletedProcess:
    notify_script = skill_root / "scripts" / "notify_owner.py"
    if not notify_script.exists():
        raise FileNotFoundError(f"notify_owner.py not found: {notify_script}")

    cmd = [sys.executable, str(notify_script), "--output-dir", str(output_dir)]
    return subprocess.run(cmd, capture_output=True, text=True)


def parse_update_output(stdout: str) -> dict:
    result = {
        "status_counts": {"绿灯": 0, "黄灯": 0, "红灯": 0, "待更新": 0},
        "overall_level": "",
        "overall_label": "",
        "latest_file": "",
        "daily_file": "",
        "evidence_rows": 0,
    }

    for line in stdout.splitlines():
        line = line.strip()
        if line.startswith("updated:"):
            result["latest_file"] = line.replace("updated:", "", 1).strip()
        elif line.startswith("daily:"):
            result["daily_file"] = line.replace("daily:", "", 1).strip()
        elif line.startswith("evidence_rows:"):
            value = line.replace("evidence_rows:", "", 1).strip()
            try:
                result["evidence_rows"] = int(value)
            except Exception:
                pass
        elif line.startswith("status_counts:"):
            raw = line.replace("status_counts:", "", 1).strip()
            try:
                parsed = ast.literal_eval(raw)
                if isinstance(parsed, dict):
                    result["status_counts"] = parsed
            except Exception:
                pass
        elif line.startswith("overall_level:"):
            result["overall_level"] = line.replace("overall_level:", "", 1).strip()
        elif line.startswith("overall_label:"):
            result["overall_label"] = line.replace("overall_label:", "", 1).strip()

    return result


def build_summary_sentence(parsed: dict) -> str:
    level = parsed.get("overall_level", "")
    label = parsed.get("overall_label", "")
    counts = parsed.get("status_counts", {})

    green = counts.get("绿灯", 0)
    yellow = counts.get("黄灯", 0)
    red = counts.get("红灯", 0)
    pending = counts.get("待更新", 0)

    if level == "绿":
        text = (
            f"当前私募信贷风险处于“绿（{label}）”阶段。"
            f"16项监测中绿灯{green}项、黄灯{yellow}项、红灯{red}项、待更新{pending}项，"
            f"整体环境平稳，尚未看到明显危机迹象。"
        )
    elif level == "浅黄":
        text = (
            f"当前私募信贷风险处于“浅黄（{label}）”阶段。"
            f"16项监测中绿灯{green}项、黄灯{yellow}项、红灯{red}项、待更新{pending}项，"
            f"异常开始增多，但仍主要停留在预警层。"
        )
    elif level == "黄":
        text = (
            f"当前私募信贷风险处于“黄（{label}）”阶段。"
            f"16项监测中绿灯{green}项、黄灯{yellow}项、红灯{red}项、待更新{pending}项，"
            f"事件压力明显，传导开始出现，但系统性确认仍不足。"
        )
    elif level == "橙":
        text = (
            f"当前私募信贷风险处于“橙（{label}）”阶段。"
            f"16项监测中绿灯{green}项、黄灯{yellow}项、红灯{red}项、待更新{pending}项，"
            f"传导压力与系统确认压力开始共振，需高度警惕。"
        )
    elif level == "红":
        text = (
            f"当前私募信贷风险处于“红（{label}）”阶段。"
            f"16项监测中绿灯{green}项、黄灯{yellow}项、红灯{red}项、待更新{pending}项，"
            f"系统层关键指标已形成共振，风险进入危机阶段。"
        )
    else:
        text = (
            f"本次监控已完成。16项监测中绿灯{green}项、黄灯{yellow}项、红灯{red}项、待更新{pending}项。"
        )

    return text[:280]


def build_summary_text(run_time: str, success: bool, parsed: dict, output_dir: Path) -> str:
    latest_file = parsed.get("latest_file", str(output_dir / "private_credit_monitor_latest.xlsx"))
    daily_file = parsed.get("daily_file", "")
    level = parsed.get("overall_level", "")
    label = parsed.get("overall_label", "")
    counts = parsed.get("status_counts", {})

    green = counts.get("绿灯", 0)
    yellow = counts.get("黄灯", 0)
    red = counts.get("红灯", 0)
    pending = counts.get("待更新", 0)

    lines = []
    lines.append("Private Credit Daily Monitor 已更新" if success else "Private Credit Daily Monitor 运行失败")
    lines.append(f"时间：{run_time}")

    if level and label:
        lines.append(f"整体等级：{level}（{label}）")

    lines.append(f"绿灯 {green} / 黄灯 {yellow} / 红灯 {red} / 待更新 {pending}")

    if success:
        lines.append(f"最新文件：{latest_file}")
        if daily_file:
            lines.append(f"归档文件：{daily_file}")
        lines.append(f"摘要：{build_summary_sentence(parsed)}")
    else:
        lines.append("摘要：本次运行失败，请检查 last_run_summary.json 中的 stdout/stderr。")
        lines.append(f"输出目录：{output_dir}")

    return "\n".join(lines)


def main() -> None:
    skill_root = Path(__file__).resolve().parents[1]
    output_dir = get_default_output_dir()

    run_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    result = run_update(skill_root, output_dir)
    success = result.returncode == 0

    parsed = parse_update_output(result.stdout)

    latest_file = parsed.get("latest_file") or str(output_dir / "private_credit_monitor_latest.xlsx")
    date_str = datetime.now().strftime("%Y-%m-%d")
    daily_file = parsed.get("daily_file") or str(output_dir / f"private_credit_monitor_{date_str}.xlsx")

    summary = {
        "run_time": run_time,
        "status": "success" if success else "failed",
        "output_dir": str(output_dir),
        "latest_file": latest_file,
        "daily_file": daily_file,
        "status_counts": parsed.get("status_counts", {}),
        "overall_level": parsed.get("overall_level", ""),
        "overall_label": parsed.get("overall_label", ""),
        "evidence_rows": parsed.get("evidence_rows", 0),
        "summary_text": build_summary_sentence(parsed),
        "stdout": result.stdout,
        "stderr": result.stderr,
    }

    summary_json = output_dir / "last_run_summary.json"
    summary_txt = output_dir / "last_run_summary.txt"

    summary_json.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    summary_txt.write_text(
        build_summary_text(run_time=run_time, success=success, parsed=parsed, output_dir=output_dir),
        encoding="utf-8",
    )

    print(f"output_dir: {output_dir}")
    print(f"summary_json: {summary_json}")
    print(f"summary_txt: {summary_txt}")
    print(result.stdout)
    if result.stderr.strip():
        print(result.stderr)

    try:
        notify_result = run_notify(skill_root, output_dir)
        print("notify_stdout_start")
        print(notify_result.stdout)
        print("notify_stdout_end")
        if notify_result.stderr.strip():
            print("notify_stderr_start")
            print(notify_result.stderr)
            print("notify_stderr_end")
    except Exception as exc:
        print(f"notify_error: {exc}")

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
