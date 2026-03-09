import importlib
import platform
import shutil
import sys
from pathlib import Path

REQUIRED_MODULES = ["requests", "openpyxl"]


def ok(msg: str) -> None:
    print(f"[OK] {msg}")


def warn(msg: str) -> None:
    print(f"[WARN] {msg}")


def fail(msg: str) -> None:
    print(f"[FAIL] {msg}")


def check_python_modules() -> bool:
    success = True
    for module_name in REQUIRED_MODULES:
        try:
            importlib.import_module(module_name)
            ok(f"Python module available: {module_name}")
        except Exception:
            success = False
            fail(f"Missing Python module: {module_name}")
    return success


def check_workspace_writable(workspace: Path) -> bool:
    probe = workspace / ".private_credit_write_test.tmp"
    try:
        probe.write_text("ok", encoding="utf-8")
        probe.unlink(missing_ok=True)
        ok(f"Workspace writable: {workspace}")
        return True
    except Exception as exc:
        fail(f"Workspace not writable: {workspace} ({exc})")
        return False


def check_scheduler() -> bool:
    system_name = platform.system().lower()

    if system_name == "windows":
        if shutil.which("schtasks"):
            ok("Windows Task Scheduler command available: schtasks")
            return True
        fail("schtasks not found")
        return False

    if shutil.which("crontab"):
        ok("cron command available: crontab")
        return True

    fail("crontab not found")
    return False


def check_template_presence(root: Path) -> None:
    asset_template = root / "assets" / "private_credit_monitor_template.xlsx"
    if asset_template.exists():
        ok(f"Asset template found: {asset_template}")
    else:
        warn("Asset template not found. update_monitor.py will create a minimal workbook automatically.")


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    workspace = Path.cwd()

    print("Running self-check...\n")

    results = [
        check_python_modules(),
        check_workspace_writable(workspace),
        check_scheduler(),
    ]
    check_template_presence(root)

    if all(results):
        print("\nSelf-check passed.")
        sys.exit(0)

    print("\nSelf-check failed. Fix the items above and run again.")
    sys.exit(1)


if __name__ == "__main__":
    main()
