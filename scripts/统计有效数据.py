#!/usr/bin/env python3

from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from xlsx_utils import read_first_sheet


BASE_DIR = Path("/Users/xinxinhuashe/Documents/易达威实习")
OUTPUT_DIR = BASE_DIR / "输出结果"
DEFAULT_DATE = date.today().isoformat()
REQUIRED_FIELDS = ("姓名", "职称", "邮箱", "学校", "学院", "采集人", "采集日期", "简介", "主页")
TEMPLATE_HEADERS = ["序号", *REQUIRED_FIELDS]


@dataclass
class FileStats:
    path: Path
    total_rows: int = 0
    valid_rows: int = 0

    @property
    def invalid_rows(self) -> int:
        return self.total_rows - self.valid_rows


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="统计当天最终输出文件中的有效数据条目。")
    parser.add_argument("--date", default=DEFAULT_DATE, help="日期目录，例如 2026-06-10")
    parser.add_argument("--dir", help="自定义输出目录，默认使用 输出结果/日期")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    selected_date = prompt_text("请输入要统计的日期", args.date)
    target_dir = resolve_target_dir(args.dir, selected_date)
    ensure_dir_exists(target_dir)

    output_files = list_output_csv_files(target_dir)
    if not output_files:
        print(f"未找到可统计的输出文件：{target_dir}")
        return

    stats_list = [count_workbook(path) for path in output_files]
    stats_list = [stats for stats in stats_list if stats is not None]
    if not stats_list:
        print(f"未找到可统计的输出记录：{target_dir}")
        return

    print_report(target_dir, stats_list)


def resolve_target_dir(raw_dir: str | None, folder_date: str) -> Path:
    if not raw_dir:
        return OUTPUT_DIR / folder_date

    path = Path(raw_dir).expanduser()
    if path.is_absolute():
        return path.resolve()
    return (BASE_DIR / path).resolve()


def prompt_text(label: str, default_value: str) -> str:
    raw = input(f"{label}（回车使用默认值：{default_value}）\n> ").strip()
    return raw or default_value


def ensure_dir_exists(target_dir: Path) -> None:
    if not target_dir.exists():
        raise FileNotFoundError(f"目录不存在：{target_dir}")
    if not target_dir.is_dir():
        raise NotADirectoryError(f"不是目录：{target_dir}")


def list_output_csv_files(target_dir: Path) -> list[Path]:
    files: list[Path] = []
    for school_dir in sorted(target_dir.iterdir()):
        if not school_dir.is_dir() or school_dir.name.startswith("."):
            continue
        for csv_file in sorted(school_dir.glob("*.csv")):
            if csv_file.name.endswith("_预览.csv"):
                continue
            files.append(csv_file.resolve())
    return files


def count_workbook(path: Path) -> FileStats | None:
    try:
        if path.suffix.lower() == ".csv":
            headers, rows = read_csv(path)
        else:
            headers, rows = read_first_sheet(path)
    except Exception:
        return None

    normalized_headers = [normalize_text(value) for value in headers[: len(TEMPLATE_HEADERS)]]
    if normalized_headers != TEMPLATE_HEADERS:
        return None

    stats = FileStats(path=path)
    for raw_row in rows:
        row = normalize_row(raw_row)
        if not row:
            continue
        stats.total_rows += 1
        if is_valid_row(row):
            stats.valid_rows += 1
    return stats


def normalize_row(raw_row: object) -> dict[str, str] | None:
    if not isinstance(raw_row, list) or len(raw_row) < len(TEMPLATE_HEADERS):
        return None
    return {
        "序号": normalize_text(raw_row[0]),
        "姓名": normalize_text(raw_row[1]),
        "职称": normalize_text(raw_row[2]),
        "邮箱": normalize_text(raw_row[3]),
        "学校": normalize_text(raw_row[4]),
        "学院": normalize_text(raw_row[5]),
        "采集人": normalize_text(raw_row[6]),
        "采集日期": normalize_text(raw_row[7]),
        "简介": normalize_text(raw_row[8]),
        "主页": normalize_text(raw_row[9]),
    }


def is_valid_row(row: dict[str, str]) -> bool:
    return all(row.get(field, "") for field in REQUIRED_FIELDS)


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def read_csv(path: Path) -> tuple[list[str], list[list[str]]]:
    import csv

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        return [], []
    return rows[0], rows[1:]


def print_report(target_dir: Path, stats_list: list[FileStats]) -> None:
    total_rows = sum(item.total_rows for item in stats_list)
    valid_rows = sum(item.valid_rows for item in stats_list)
    invalid_rows = total_rows - valid_rows

    print(f"统计目录：{target_dir}")
    print(f"统计规则：仅统计当天学校文件夹下的最终 CSV；仅当 {'、'.join(REQUIRED_FIELDS)} 都非空时，记为有效数据")
    print("")

    for item in sorted(stats_list, key=lambda current: current.path.as_posix()):
        display_path = safe_relative_path(item.path, target_dir)
        print(
            f"{display_path} | 总条目 {item.total_rows} | "
            f"有效 {item.valid_rows} | 无效 {item.invalid_rows}"
        )

    print("")
    print(
        f"汇总 | 文件数 {len(stats_list)} | 总条目 {total_rows} | "
        f"有效 {valid_rows} | 无效 {invalid_rows}"
    )


def safe_relative_path(path: Path, target_dir: Path) -> Path:
    try:
        return path.relative_to(target_dir)
    except ValueError:
        return path


if __name__ == "__main__":
    main()
