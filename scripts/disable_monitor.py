import subprocess

TASK_NAME = "PrivateCreditDailyMonitor"


def main() -> None:
    cp = subprocess.run(
        ["schtasks", "/Delete", "/TN", TASK_NAME, "/F"],
        capture_output=True,
        text=True,
    )

    if cp.returncode == 0:
        print(f"Monitor disabled: {TASK_NAME}")
        print("Daily scheduled task has been removed.")
    else:
        stdout_text = (cp.stdout or "").strip()
        stderr_text = (cp.stderr or "").strip()

        if "cannot find the file specified" in stdout_text.lower() or \
           "cannot find the file specified" in stderr_text.lower() or \
           "找不到指定的文件" in stdout_text or \
           "找不到指定的文件" in stderr_text:
            print(f"Monitor already disabled: {TASK_NAME}")
        else:
            print(f"Failed to disable monitor: {TASK_NAME}")
            if stdout_text:
                print("stdout_start")
                print(stdout_text)
                print("stdout_end")
            if stderr_text:
                print("stderr_start")
                print(stderr_text)
                print("stderr_end")


if __name__ == "__main__":
    main()
