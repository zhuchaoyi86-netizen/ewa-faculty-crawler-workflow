#!/usr/bin/env python3

"""
八爪鱼教师列表清洗脚本

作用：
1. 把八爪鱼导出的原始 CSV 清洗成标准输入表
2. 统一输出字段：姓名、主页、学校、学院
3. 尽量处理列错位、非标准表头、HTML 残片、多个链接列等情况
"""

from __future__ import annotations

import argparse
import csv
import html
import re
from datetime import date
from pathlib import Path
from typing import Dict, List


OUTPUT_HEADERS = ["姓名", "主页", "学校", "学院"]
URL_RE = re.compile(r"^https?://", re.IGNORECASE)
NAME_RE = re.compile(r"[\u4e00-\u9fa5·]{2,10}")
BASE_DIR = Path("/Users/xinxinhuashe/Documents/易达威实习")
DEFAULT_INPUT_DIR = BASE_DIR / "原始输入"
CLEAN_OUTPUT_DIR = BASE_DIR / "清洗结果"
SCHOOL_SUFFIXES = ("大学", "学校", "研究院")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="清洗八爪鱼导出的教师列表 CSV。")
    parser.add_argument("--input", help="原始 CSV 路径")
    parser.add_argument("--output", help="清洗后 CSV 路径")
    parser.add_argument("--school", default="", help="默认学校")
    parser.add_argument("--college", default="", help="默认学院")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = resolve_input_path(args.input) if args.input else prompt_input_file()
    output_path = Path(args.output).expanduser().resolve() if args.output else build_default_output_path(input_path)

    ensure_file_exists(input_path)

    with input_path.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))

    if not rows:
        raise ValueError("输入文件为空，无法处理。")

    header_row, data_rows = rows[0], rows[1:]
    column_map = build_column_map(header_row, data_rows)

    cleaned_rows = []
    for row in data_rows:
        if not any(normalize_text(cell) for cell in row):
            continue
        cleaned_rows.extend(normalize_row(row, column_map, args.school, args.college))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_HEADERS)
        writer.writeheader()
        writer.writerows(cleaned_rows)

    print(f"清洗完成：{len(cleaned_rows)} 条")
    print(f"输出文件：{output_path}")


def ensure_file_exists(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(f"文件不存在：{path}")


def build_default_output_path(input_path: Path) -> Path:
    school_dir = infer_school_from_input_path(input_path)
    base_dir = CLEAN_OUTPUT_DIR / detect_date_folder(input_path)
    if school_dir:
        base_dir = base_dir / school_dir
    return base_dir / f"{input_path.stem}_清洗.csv"


def build_column_map(header_row: List[str], data_rows: List[List[str]]) -> Dict[str, int]:
    normalized_headers = [normalize_text(cell).lower() for cell in header_row]
    header_map = {
        "name": find_header_index(normalized_headers, ["姓名", "教师姓名", "name", "字段1"]),
        "url": find_header_index(normalized_headers, ["主页", "链接", "链接地址", "url", "profile_url", "字段"]),
        "school": find_header_index(normalized_headers, ["学校", "school"]),
        "college": find_header_index(normalized_headers, ["学院", "院系", "college"]),
    }

    inferred = infer_columns(data_rows)
    if header_map["name"] < 0:
        header_map["name"] = inferred.get("name", -1)
    if header_map["url"] < 0:
        header_map["url"] = inferred.get("url", -1)

    return header_map


def find_header_index(headers: List[str], aliases: List[str]) -> int:
    alias_set = {item.lower() for item in aliases}
    for index, header in enumerate(headers):
        if header in alias_set:
            return index
    return -1


def infer_columns(data_rows: List[List[str]]) -> Dict[str, int]:
    samples = [row for row in data_rows if any(normalize_text(cell) for cell in row)][:15]
    if not samples:
        return {}

    max_cols = max(len(row) for row in samples)
    scores = [{"url": 0, "name": 0} for _ in range(max_cols)]

    for row in samples:
        for index in range(max_cols):
            cell = normalize_text(row[index]) if index < len(row) else ""
            if not cell:
                continue
            if looks_like_url(cell):
                scores[index]["url"] += 1
            if looks_like_name(cell):
                scores[index]["name"] += 1

    best_url = max(range(max_cols), key=lambda i: scores[i]["url"], default=-1)
    best_name = max(range(max_cols), key=lambda i: scores[i]["name"], default=-1)

    result = {}
    if best_url >= 0 and scores[best_url]["url"] > 0:
        result["url"] = best_url
    if best_name >= 0 and scores[best_name]["name"] > 0:
        result["name"] = best_name
    return result


def normalize_row(row: List[str], column_map: Dict[str, int], default_school: str, default_college: str) -> List[Dict[str, str]]:
    urls = extract_urls(row, column_map)
    names = extract_names(row, column_map)
    school = get_cell(row, column_map.get("school", -1)) or normalize_text(default_school)
    college = get_cell(row, column_map.get("college", -1)) or normalize_text(default_college)

    if not names and urls:
        names = ["" for _ in urls]
    if names and not urls:
        urls = ["" for _ in names]
    if not names and not urls:
        return []

    if len(urls) == 1 and len(names) > 1:
        urls = urls * len(names)
    elif len(names) == 1 and len(urls) > 1:
        names = names * len(urls)

    count = max(len(names), len(urls))
    rows = []
    for index in range(count):
        name = clean_name(names[index]) if index < len(names) else ""
        url = urls[index] if index < len(urls) else ""
        if not name and not url:
            continue
        rows.append(
            {
                "姓名": name,
                "主页": url,
                "学校": school,
                "学院": college,
            }
        )
    return rows


def get_preferred_url(row: List[str], column_map: Dict[str, int]) -> str:
    direct = get_cell(row, column_map.get("url", -1))
    if looks_like_url(direct):
        return direct

    for cell in row:
        value = normalize_text(cell)
        if looks_like_url(value):
            return value
    return ""


def get_preferred_name(row: List[str], column_map: Dict[str, int]) -> str:
    direct = get_cell(row, column_map.get("name", -1))
    if looks_like_name(direct):
        return direct

    for cell in row:
        value = strip_html_tags(cell)
        if looks_like_name(value):
            return value
    return ""


def extract_urls(row: List[str], column_map: Dict[str, int]) -> List[str]:
    urls = []
    direct = get_cell(row, column_map.get("url", -1))
    if looks_like_url(direct):
        urls.append(direct)

    for cell in row:
        value = normalize_text(cell)
        if looks_like_url(value) and value not in urls:
            urls.append(value)
    return urls


def extract_names(row: List[str], column_map: Dict[str, int]) -> List[str]:
    names = []
    direct = get_cell(row, column_map.get("name", -1))
    names.extend(split_names_from_cell(direct))

    for cell in row:
        value = strip_html_tags(cell)
        for name in split_names_from_cell(value):
            if name not in names:
                names.append(name)
    return names


def split_names_from_cell(value: str) -> List[str]:
    text = normalize_name_text(strip_html_tags(value))
    if not text:
        return []

    if looks_like_name(text):
        return [text]

    delimiters = ["、", "，", ",", ";", "；", "/", "|", "\n", " "]
    parts = [text]
    for delimiter in delimiters:
        next_parts = []
        for part in parts:
            next_parts.extend(part.split(delimiter))
        parts = next_parts

    candidates = [normalize_name_text(part) for part in parts if looks_like_name(part)]
    if len(candidates) >= 2:
        return dedupe_keep_order(candidates)

    regex_names = NAME_RE.findall(text)
    if len(regex_names) >= 2:
        return dedupe_keep_order(regex_names)

    return [text] if looks_like_name(text) else []


def dedupe_keep_order(items: List[str]) -> List[str]:
    result = []
    seen = set()
    for item in items:
        key = normalize_name_text(item)
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(key)
    return result


def get_cell(row: List[str], index: int) -> str:
    if index < 0 or index >= len(row):
        return ""
    return normalize_text(row[index])


def clean_name(value: str) -> str:
    text = strip_html_tags(value)
    text = re.sub(r"(?<!\S)(教授|副教授|讲师|博士|硕导|博导|老师)(?!\S)", "", text)
    text = re.sub(r"[ ]+", " ", text).strip()
    return normalize_name_text(text)


def strip_html_tags(value: str) -> str:
    text = html.unescape(str(value or ""))
    text = re.sub(r"<[^>]+>", " ", text)
    return normalize_text(text)


def normalize_text(value: str) -> str:
    text = html.unescape(str(value or ""))
    text = text.replace("\ufeff", "").replace("ï»¿", "")
    text = text.replace("\u00a0", " ").replace("\u3000", " ").replace("\r", "\n").replace("\t", " ")
    text = re.sub(r"[ ]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def looks_like_url(value: str) -> bool:
    return bool(URL_RE.match(normalize_text(value)))


def looks_like_name(value: str) -> bool:
    text = normalize_name_text(strip_html_tags(value))
    if not text or looks_like_url(text):
        return False
    return bool(NAME_RE.fullmatch(text))


def normalize_name_text(value: str) -> str:
    text = normalize_text(value)
    text = re.sub(r"(?<=[\u4e00-\u9fa5])\s+(?=[\u4e00-\u9fa5])", "", text)
    return text


def prompt_input_file() -> Path:
    print("请输入原始输入 CSV 路径或文件名关键词。")
    print("输入完整路径/相对路径：直接使用该文件")
    print("输入关键词：自动搜索原始输入目录")

    while True:
        raw = input("> ").strip()
        if not raw:
            print("请输入文件路径或关键词。")
            continue

        direct = resolve_input_path(raw)
        if direct.exists():
            return direct

        matches = search_csv_files(raw)
        if not matches:
            print(f"没有找到包含“{raw}”的原始 CSV 文件。\n")
            continue
        if len(matches) == 1:
            print(f"已自动匹配：{matches[0]}\n")
            return matches[0]
        return choose_from_matches(raw, matches)


def resolve_input_path(raw: str) -> Path:
    candidate = Path(raw).expanduser()
    if candidate.is_absolute():
        return candidate.resolve()
    return (BASE_DIR / candidate).resolve()


def search_csv_files(keyword: str) -> List[Path]:
    keyword_lower = keyword.lower()
    matches = []
    for path in DEFAULT_INPUT_DIR.rglob("*.csv"):
        relative = path.relative_to(BASE_DIR).as_posix().lower()
        compact_relative = relative.replace("西安理工大学", "").replace("_", "").replace("-", "")
        compact_keyword = keyword_lower.replace("西安理工大学", "").replace("_", "").replace("-", "")
        if (
            keyword_lower in relative
            or keyword_lower in path.name.lower()
            or compact_keyword in compact_relative
            or is_subsequence(compact_keyword, compact_relative)
        ):
            matches.append(path.resolve())
    return sorted(matches)


def choose_from_matches(keyword: str, matches: List[Path]) -> Path:
    print(f"找到 {len(matches)} 个包含“{keyword}”的原始 CSV 文件，请输入编号：")
    for index, path in enumerate(matches, start=1):
        print(f"{index}. {path.relative_to(BASE_DIR)}")

    while True:
        raw = input("> ").strip()
        if not raw.isdigit():
            print("请输入数字编号。")
            continue
        selected = int(raw)
        if 1 <= selected <= len(matches):
            print(f"已选择：{matches[selected - 1]}\n")
            return matches[selected - 1]
        print("编号超出范围，请重新输入。")


def is_subsequence(needle: str, haystack: str) -> bool:
    if not needle:
        return False
    it = iter(haystack)
    return all(char in it for char in needle)


def detect_date_folder(path: Path) -> str:
    for part in path.parts:
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", part):
            return part
    return date.today().isoformat()


def infer_school_from_input_path(input_path: Path) -> str:
    for candidate in [input_path.stem, *input_path.parts]:
        school = extract_school_name(candidate)
        if school:
            return school
    return ""


def extract_school_name(value: str) -> str:
    normalized = normalize_text(value)
    for suffix in SCHOOL_SUFFIXES:
        matches = re.findall(rf"[\u4e00-\u9fa5A-Za-z0-9·（）()]+?{re.escape(suffix)}", normalized)
        if matches:
            matches.sort(key=len, reverse=True)
            return matches[0]
    return ""


if __name__ == "__main__":
    main()
