#!/usr/bin/env python3

from datetime import date
from pathlib import Path
import sys


BASE_DIR = Path("/Users/xinxinhuashe/Documents/易达威实习")
SCRIPT_DIR = BASE_DIR / "scripts"

if str(SCRIPT_DIR.resolve()) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR.resolve()))

from teacher_profile_pipeline import update_aggregate_workbooks  # noqa: E402


def prompt_text(label: str, default_value: str) -> str:
    raw = input(f"{label}（回车使用默认值：{default_value}）\n> ").strip()
    return raw or default_value


def main() -> None:
    selected_date = prompt_text("请输入要同步的日期", date.today().isoformat())
    update_aggregate_workbooks(selected_date)
    print(f"已同步：输出结果/{selected_date}/按院系汇总.xlsx")


if __name__ == "__main__":
    main()
