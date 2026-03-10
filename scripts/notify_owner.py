import argparse
import sys
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


def build_notification_message(summary_text: str) -> str:
    lines = []
    lines.append("【Private Credit Daily Monitor】")
    lines.append(summary_text)
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="Prepare owner notification text from last_run_summary.txt")
    parser.add_argument("--output-dir", default="", help="Optional explicit output directory")
    args = parser.parse_args()

    output_dir = Path(args.output_dir).resolve() if args.output_dir else get_default_output_dir()
    summary_text = read_summary_text(output_dir)
    message = build_notification_message(summary_text)

    print("notification_message_start")
    print(message)
    print("notification_message_end")


if __name__ == "__main__":
    main()
