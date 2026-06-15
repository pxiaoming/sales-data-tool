"""一键运行入口：自动读取 inputs/ 中最新的 Excel 并生成报表。"""

from __future__ import annotations

import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    APP_ROOT = Path(sys.executable).resolve().parent
    BUNDLE_ROOT = Path(getattr(sys, "_MEIPASS"))
else:
    APP_ROOT = Path(__file__).resolve().parent.parent
    BUNDLE_ROOT = APP_ROOT
    sys.path.insert(0, str(APP_ROOT / "src"))

from finance_excel_flow.console import configure_console_encoding  # noqa: E402
from finance_excel_flow.paths import is_frozen  # noqa: E402

INPUT_DIR = APP_ROOT / "inputs"
OUTPUT_DIR = APP_ROOT / "output"
CONFIG_PATH = BUNDLE_ROOT / "configs" / "template-example.yml"

EXCEL_SUFFIXES = {".xlsx", ".xlsm"}


def _ensure_dirs() -> None:
    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)


def _find_latest_input() -> Path | None:
    candidates = [
        path
        for path in INPUT_DIR.iterdir()
        if path.is_file()
        and path.suffix.lower() in EXCEL_SUFFIXES
        and not path.name.startswith("~$")
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def _friendly_error(exc: BaseException) -> str:
    message = str(exc)
    if isinstance(exc, FileNotFoundError):
        return f"找不到文件：{message}"
    if "must be a mapping" in message:
        return f"配置文件格式有误，请联系维护人员：{message}"
    if "Missing required column" in message or "required column" in message.lower():
        return f"源 Excel 缺少必要列，请检查文件是否正确：{message}"
    if "Worksheet" in message and "does not exist" in message:
        return f"源 Excel 中找不到指定的工作表，请检查文件：{message}"
    if "Cannot convert value to number" in message:
        return f"源 Excel 中存在无法识别的数值，请检查数据格式：{message}"
    if "Permission denied" in message or "PermissionError" in type(exc).__name__:
        return "无法读取或写入文件，请确认 Excel 文件已关闭后再试。"
    return message


def _wait_for_exit(code: int) -> int:
    if is_frozen():
        print()
        try:
            input("按回车键关闭...")
        except EOFError:
            pass
    return code


def main() -> int:
    configure_console_encoding()
    _ensure_dirs()

    if not CONFIG_PATH.is_file():
        print(f"错误：找不到配置文件 {CONFIG_PATH.name}，请联系维护人员。")
        return _wait_for_exit(1)

    input_path = _find_latest_input()
    if input_path is None:
        print("错误：inputs 文件夹里没有 Excel 文件。")
        print(f"请将源文件（.xlsx）放入：{INPUT_DIR}")
        return _wait_for_exit(1)

    print("=" * 50)
    print("  KA 促销费用报表生成")
    print("=" * 50)
    print(f"源文件：{input_path.name}")
    print(f"配置文件：{CONFIG_PATH.name}")
    print("正在处理，请稍候...")
    print()

    try:
        from finance_excel_flow.config import load_config
        from finance_excel_flow.engine import process_file

        config = load_config(CONFIG_PATH)
        result_path = process_file(
            input_path=input_path,
            config=config,
            output_dir=OUTPUT_DIR,
        )
    except Exception as exc:
        print("处理失败。")
        print(_friendly_error(exc))
        return _wait_for_exit(1)

    print("处理完成！")
    print(f"输出文件：{result_path}")
    print(f"输出目录：{OUTPUT_DIR}")
    return _wait_for_exit(0)


if __name__ == "__main__":
    raise SystemExit(main())
