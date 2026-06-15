#!/usr/bin/env python3

"""
教师详情补采脚本（Python 版）

输入：教师列表 CSV
输出：符合导入模板字段的 XLSX

设计目标：
1. 尽量少依赖，使用 Python 标准库即可运行
2. 优先提取姓名附近的邮箱和职称，降低误抓
3. 保留日志，方便人工复查异常记录
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import functools
import hashlib
import html
import ast
import json
import os
import re
import shutil
import subprocess
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urldefrag, urljoin, urlsplit
from urllib.request import Request, urlopen

from xlsx_utils import read_first_sheet, write_csv, write_sheet_csv_preview, write_workbook


TEMPLATE_HEADERS = [
    "序号",
    "姓名",
    "职称",
    "邮箱",
    "学校",
    "学院",
    "采集人",
    "采集日期",
    "简介",
    "主页",
]

DEFAULT_HEADERS = {
    "name": ["name", "姓名", "教师姓名", "字段1", "姓名名称"],
    "profile_url": ["profile_url", "主页", "教师主页", "详情页", "链接", "url", "链接地址", "网址"],
    "school": ["school", "学校", "院校"],
    "college": ["college", "学院", "院系", "学部"],
    "title": ["title", "职称"],
    "email": ["email", "邮箱", "电子邮箱"],
    "intro": ["intro", "简介", "个人简介"],
}
GENERIC_PLACEHOLDER_HEADERS = {
    "字段",
    "字段1",
    "字段2",
    "字段3",
    "字段4",
    "字段1_文本",
    "字段1_链接",
    "字段2_文本",
    "字段2_链接",
    "字段3_文本",
    "字段3_链接",
}

KNOWN_TITLES = [
    "科研副教授",
    "科研教授",
    "特聘教授",
    "特聘副教授",
    "高级实验师",
    "助理教授",
    "副教授",
    "教授",
    "助理研究员",
    "副研究员",
    "研究员",
    "高级工程师",
    "助理工程师",
    "工程师",
    "实验师",
    "讲师",
    "教师",
]
PRIORITY_TITLES = [
    "副教授",
    "教授",
    "讲师",
]
TITLE_ALIASES = {
    "科研副教授": ["科研副教授"],
    "科研教授": ["科研教授"],
    "特聘教授": ["特聘教授"],
    "特聘副教授": ["特聘副教授"],
    "高级实验师": ["高级实验师"],
    "助理教授": ["助理教授", "assistant professor"],
    "副教授": ["副教授", "associate professor"],
    "教授": ["教授", "professor"],
    "助理研究员": ["助理研究员", "assistant researcher"],
    "副研究员": ["副研究员", "associate researcher"],
    "研究员": ["研究员", "researcher"],
    "高级工程师": ["高级工程师", "senior engineer"],
    "助理工程师": ["助理工程师", "assistant engineer"],
    "工程师": ["工程师", "engineer"],
    "实验师": ["实验师"],
    "讲师": ["讲师", "lecturer"],
    "教师": ["教师", "teacher"],
}
TITLE_PATTERN_ENTRIES = [
    (
        canonical,
        re.compile(rf"\b{re.escape(alias)}\b", re.IGNORECASE) if re.search(r"[A-Za-z]", alias) else re.compile(re.escape(alias)),
    )
    for canonical, aliases in TITLE_ALIASES.items()
    for alias in aliases
]

EMAIL_PATTERN = re.compile(
    r"[A-Z0-9._%+-]+@[A-Z0-9.-]+?\.(?:com|cn|edu|net|org|gov)(?:\.[A-Z]{2,3})?",
    re.IGNORECASE,
)
EMAIL_LABEL_PATTERNS = [
    re.compile(r"(电子邮件|电子信箱|电子邮箱|邮箱|E-mail|Email|Emial)\s*[:：]?\s*([^\n]{0,80})", re.IGNORECASE),
    re.compile(r"(联系方式|联系邮箱|邮箱地址)\s*[:：]?\s*([^\n]{0,120})", re.IGNORECASE),
]
TITLE_LABEL_CAPTURE_PATTERNS = [
    re.compile(r"职\s*称\s*[:：]?\s*([\s\S]{0,80})", re.IGNORECASE),
    re.compile(r"职\s*称\s*/\s*职\s*务\s*[:：]?\s*([\s\S]{0,80})", re.IGNORECASE),
    re.compile(r"职\s*务\s*/\s*职\s*称\s*[:：]?\s*([\s\S]{0,80})", re.IGNORECASE),
    re.compile(r"基本情况[\s\S]{0,120}?职\s*称\s*[:：]?\s*([\s\S]{0,80})", re.IGNORECASE),
]
PROFESSOR_LEVEL_PATTERN = re.compile(r"([一二三四1234])级?教授")
ASSOCIATE_PROFESSOR_LEVEL_PATTERN = re.compile(r"([五六七567])级?(?:副教授|教授)")
IMAGE_URL_PATTERN = re.compile(r'<img[^>]+src=["\']([^"\']+)["\']', re.IGNORECASE)
VSB_CONTENT_PATTERN = re.compile(r'<div id="vsb_content".*?>[\s\S]*?</div>', re.IGNORECASE)
SECTION_BLOCK_PATTERN = re.compile(
    r'<span id="t\d+">\s*([^<]+?)\s*</span>[\s\S]*?<div class="nr" id="n\d+">([\s\S]*?)</div><!--ckendn\d+-->',
    re.IGNORECASE,
)
TSITES_ENCRYPT_SPAN_PATTERN = re.compile(
    r'<span[^>]+_tsites_encrypt_field=["\']_tsites_encrypt_field["\'][^>]*id=["\']([^"\']+)["\'][^>]*>([\s\S]*?)</span>',
    re.IGNORECASE,
)
TEACHER_CARD_PATTERN = re.compile(
    r"<li>\s*<a[^>]+href=[\"']([^\"']+)[\"'][^>]*>\s*<div[^>]+class=[\"']pic fl[\"'][^>]*>[\s\S]{0,1200}?<div[^>]+class=[\"']con fl[\"'][^>]*>\s*<p[^>]+class=[\"']btn[\"'][^>]*>[\s\S]{0,200}?</p>\s*<p[^>]+class=[\"']bt[\"'][^>]*>([\s\S]*?)</p>\s*<p[^>]+class=[\"']zy[\"'][^>]*>([\s\S]*?)</p>[\s\S]{0,400}?</a>\s*</li>",
    re.IGNORECASE,
)
TEACHER_CARD_SECTION_PATTERN = re.compile(
    r"centerCutImg\.js[\s\S]*?<ul>([\s\S]*?)</ul>",
    re.IGNORECASE,
)
TEACHER_LIST_LINK_PATTERN = re.compile(
    r'<a[^>]+href=["\']([^"\']*?/info/\d+/\d+\.htm)["\'][^>]*>([\s\S]*?)</a>',
    re.IGNORECASE,
)
VSB_WIDGET_TITLE_PATTERN = re.compile(r"var\s+(u_u\d+)_title\s*=\s*(\[[\s\S]*?\]);", re.IGNORECASE)
VSB_WIDGET_ID_PATTERN_TEMPLATE = r"var\s+{prefix}_id\s*=\s*(\[[\s\S]*?\]);"
VSB_WIDGET_SUMMARY_PATTERN_TEMPLATE = r"var\s+{prefix}_summary\s*=\s*(\[[\s\S]*?\]);"
VSB_WIDGET_HIDDEN_LINK_PATTERN_TEMPLATE = r'<a[^>]+id=["\']{prefix}_(\d+)_a["\'][^>]+href=["\']([^"\']+)["\']'
SCRIPT_DIR = Path(__file__).resolve().parent
VISION_OCR_SCRIPT = SCRIPT_DIR / "vision_ocr.swift"
AGGREGATE_RECORDS_DIRNAME = ".汇总数据/records"
ALL_REQUIRED_TEMPLATE_HEADERS = TEMPLATE_HEADERS[1:]
SCHOOL_SUFFIXES = ("大学", "学校", "研究院")
COLLEGE_SUFFIXES = ("学院", "系", "中心", "部", "所", "研究院")
NAME_NOISE_TOKENS = ("大学", "学院", "学部", "研究院", "中心", "教师", "主页", "首页", "中文", "英文", "个人")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="根据姓名和教师主页链接，补采职称、邮箱、简介并输出导入 XLSX。"
    )
    parser.add_argument("--input", help="输入 CSV 路径")
    parser.add_argument("--list-page-url", help="单页列表型师资页 URL，直接从页面拆出多位老师信息")
    parser.add_argument("--profile-url", help="单个教师主页 URL，直接补采该老师信息")
    parser.add_argument("--output", help="输出 XLSX 路径")
    parser.add_argument("--collector", default="未填写", help="采集人")
    parser.add_argument("--date", default=today(), help="采集日期，例如 2026-06-01")
    parser.add_argument("--school", default="", help="默认学校，仅在输入表该列为空时补入")
    parser.add_argument("--college", default="", help="默认学院，仅在输入表该列为空时补入")
    parser.add_argument("--concurrency", type=int, default=5, help="并发数，默认 5")
    parser.add_argument("--limit", type=int, help="仅处理前 N 条，便于调试")
    parser.add_argument("--log", help="日志 JSON 路径")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.input and not args.list_page_url and not args.profile_url:
        raise ValueError("请提供 --input、--list-page-url 或 --profile-url。")

    input_path = Path(args.input).expanduser().resolve() if args.input else None
    log_entries: List[Dict[str, object]] = []
    selected_rows_count = 0
    effective_school, effective_college = infer_default_org_fields(
        input_path=input_path,
        list_page_url=args.list_page_url or "",
        profile_url=args.profile_url or "",
        fallback_school=args.school,
        fallback_college=args.college,
    )

    if args.profile_url:
        results = parse_teacher_profile_page(
            args.profile_url,
            default_school=effective_school,
            default_college=effective_college,
            log_entries=log_entries,
        )
        selected_rows_count = len(results)
    elif args.list_page_url:
        list_records = parse_teacher_list_page(
            args.list_page_url,
            default_school=effective_school,
            default_college=effective_college,
            log_entries=log_entries,
            limit=args.limit,
        )
        selected_rows_count = len(list_records)
        concurrency = max(1, min(args.concurrency, 20))
        results = merge_duplicate_profiles(process_records(list_records, concurrency, log_entries))
    else:
        assert input_path is not None
        fallback_title = infer_group_title_from_filename(input_path.stem)
        ensure_file_exists(input_path)

        with input_path.open("r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))

        if not rows:
            raise ValueError("输入文件为空，无法处理。")

        header_row, data_rows = rows[0], rows[1:]
        header_map = build_header_map(header_row, data_rows)
        selected_rows = [row for row in data_rows if any(cell.strip() for cell in row)]
        if args.limit:
            selected_rows = selected_rows[: max(1, args.limit)]
        selected_rows_count = len(selected_rows)

        records: List[Dict[str, str]] = []
        previous_school = normalize_whitespace(effective_school)
        previous_college = normalize_whitespace(effective_college)
        for index, row in enumerate(selected_rows):
            row_records = expand_seed_rows(
                row,
                header_map,
                index,
                default_school=effective_school,
                default_college=effective_college,
                fallback_title=fallback_title,
            )
            for record in row_records:
                school_value = normalize_whitespace(record.get("school", ""))
                if is_placeholder_school(school_value) or (school_value and not looks_like_school(school_value)):
                    record["school"] = previous_school or normalize_whitespace(effective_school)
                elif school_value:
                    previous_school = school_value

                college_value = normalize_whitespace(record.get("college", ""))
                if is_placeholder_college(college_value) or (college_value and not looks_like_college(college_value)):
                    record["college"] = previous_college or normalize_whitespace(effective_college)
                elif college_value:
                    previous_college = college_value
                records.append(record)

        concurrency = max(1, min(args.concurrency, 20))
        results = merge_duplicate_profiles(process_records(records, concurrency, log_entries))

    raw_output_rows = [
        to_template_row(
            detail,
            index=index + 1,
            collector=args.collector,
            collected_at=args.date,
        )
        for index, detail in enumerate(results)
    ]
    output_rows = renumber_template_rows(
        [row for row in raw_output_rows if is_valid_template_row(row)]
    )
    filtered_rows_count = len(raw_output_rows) - len(output_rows)

    resolved_school = infer_output_school(output_rows, effective_school)
    resolved_college = infer_output_college(output_rows, effective_college, input_path, args.list_page_url or args.profile_url or "")
    output_path = (
        Path(args.output).expanduser().resolve()
        if args.output
        else create_default_output_path(input_path, args.date, resolved_school, resolved_college)
        if input_path
        else create_default_output_path_from_url(args.list_page_url or args.profile_url or "", args.date, resolved_school, resolved_college)
    )
    output_path = ensure_csv_output_path(output_path)
    log_path = Path(args.log).expanduser().resolve() if args.log else create_default_log_path(output_path)
    existing_rows = load_existing_output_rows(output_path)
    output_rows = merge_template_output_rows(existing_rows, output_rows)
    manifest_path = create_output_manifest_path(output_path, args.date)

    if output_rows:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        write_csv(output_path, TEMPLATE_HEADERS, output_rows)
        write_output_manifest(
            manifest_path,
            input_path=input_path,
            list_page_url=args.list_page_url or args.profile_url or "",
            output_path=output_path,
            collector=args.collector,
            collected_at=args.date,
            school=effective_school,
            college=effective_college,
            rows=output_rows,
        )
    else:
        remove_path_if_exists(output_path)
        remove_path_if_exists(manifest_path)
    remove_path_if_exists(create_legacy_xlsx_path(output_path))
    remove_path_if_exists(create_preview_csv_path(output_path))
    update_aggregate_workbooks(args.date)

    log_payload = {
        "createdAt": dt.datetime.now().isoformat(),
        "inputPath": str(input_path) if input_path else "",
        "listPageUrl": args.list_page_url or "",
        "profileUrl": args.profile_url or "",
        "outputPath": str(output_path),
        "totalInputRows": selected_rows_count,
        "processedRows": selected_rows_count,
        "exportedRows": len(output_rows),
        "filteredInvalidRows": filtered_rows_count,
        "collector": args.collector,
        "collectedAt": args.date,
        "logs": log_entries,
    }

    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("w", encoding="utf-8") as f:
        json.dump(log_payload, f, ensure_ascii=False, indent=2)

    print(f"处理完成：{selected_rows_count} 条")
    if len(results) != selected_rows_count:
        print(f"合并后输出：{len(results)} 条")
    if filtered_rows_count:
        print(f"有效数据导出：{len(output_rows)} 条，过滤空数据：{filtered_rows_count} 条")
    print(f"输出文件：{output_path if output_rows else '未生成（无有效数据）'}")
    print(f"日志文件：{log_path}")


def process_records(records: List[Dict[str, str]], concurrency: int, log_entries: List[Dict[str, object]]) -> List[Dict[str, str]]:
    results: List[Dict[str, str] | None] = [None] * len(records)
    with ThreadPoolExecutor(max_workers=min(concurrency, len(records) or 1)) as executor:
        future_map = {
            executor.submit(enrich_record, record, 15, log_entries): index
            for index, record in enumerate(records)
        }
        for future in as_completed(future_map):
            index = future_map[future]
            results[index] = future.result()
    return [item for item in results if item is not None]


def merge_duplicate_profiles(records: List[Dict[str, str]]) -> List[Dict[str, str]]:
    merged_records: List[Dict[str, str]] = []

    for record in records:
        matched_index = find_merge_target_index(merged_records, record)
        if matched_index < 0:
            merged_records.append(dict(record))
            continue
        merged_records[matched_index] = merge_profile_pair(merged_records[matched_index], record)

    return merged_records


def find_merge_target_index(records: List[Dict[str, str]], candidate: Dict[str, str]) -> int:
    candidate_name = clean_name(candidate.get("name", ""))
    candidate_email = normalize_whitespace(candidate.get("email", "")).lower()
    candidate_school = normalize_whitespace(candidate.get("school", ""))
    candidate_college = normalize_whitespace(candidate.get("college", ""))
    candidate_profile_key = normalize_profile_identity(candidate.get("profile_url", ""))

    for index, existing in enumerate(records):
        existing_name = clean_name(existing.get("name", ""))
        existing_email = normalize_whitespace(existing.get("email", "")).lower()
        existing_school = normalize_whitespace(existing.get("school", ""))
        existing_college = normalize_whitespace(existing.get("college", ""))
        existing_profile_key = normalize_profile_identity(existing.get("profile_url", ""))

        same_org = candidate_school == existing_school and candidate_college == existing_college
        same_name = candidate_name and existing_name and candidate_name == existing_name
        same_email = candidate_email and existing_email and candidate_email == existing_email
        same_profile = candidate_profile_key and existing_profile_key and candidate_profile_key == existing_profile_key

        if same_org and (same_name or (same_email and same_profile)):
            return index

    return -1


def merge_profile_pair(left: Dict[str, str], right: Dict[str, str]) -> Dict[str, str]:
    primary, secondary = choose_primary_profile(left, right)
    merged = dict(primary)

    for field in ["email", "title", "intro", "profile_url", "school", "college"]:
        if not normalize_whitespace(merged.get(field, "")):
            merged[field] = secondary.get(field, "")

    if normalize_whitespace(secondary.get("email", "")) and not normalize_whitespace(merged.get("email", "")):
        merged["email"] = secondary["email"]
    if normalize_whitespace(secondary.get("title", "")) and not normalize_whitespace(merged.get("title", "")):
        merged["title"] = secondary["title"]

    return merged


def choose_primary_profile(left: Dict[str, str], right: Dict[str, str]) -> tuple[Dict[str, str], Dict[str, str]]:
    left_score = profile_preference_score(left)
    right_score = profile_preference_score(right)
    return (left, right) if left_score >= right_score else (right, left)


def profile_preference_score(record: Dict[str, str]) -> tuple[int, int, int, int]:
    raw_name = normalize_whitespace(record.get("name", ""))
    intro = normalize_whitespace(record.get("intro", ""))
    url = normalize_whitespace(record.get("profile_url", ""))

    score_non_english_name = 0 if "英文" in raw_name.lower() else 1
    score_chinese_intro = 1 if contains_more_chinese_than_english(intro) else 0
    score_has_title = 1 if normalize_whitespace(record.get("title", "")) else 0
    score_has_email = 1 if normalize_whitespace(record.get("email", "")) else 0
    score_non_english_url = 0 if "english" in url.lower() else 1

    return (
        score_non_english_name,
        score_non_english_url,
        score_chinese_intro + score_has_title,
        score_has_email,
    )


def normalize_profile_identity(url: str) -> str:
    normalized = normalize_whitespace(url)
    if not normalized:
        return ""
    split = urlsplit(urldefrag(normalized).url or normalized)
    path = split.path or ""
    path = re.sub(r"/(zh_CN|en|en_US|cn)(?=/|$)", "", path, flags=re.IGNORECASE)
    path = re.sub(r"/index\.htm[l]?$", "", path, flags=re.IGNORECASE)
    path = re.sub(r"/+$", "", path)
    return f"{split.netloc.lower()}{path.lower()}"


def contains_more_chinese_than_english(value: str) -> bool:
    chinese_count = len(re.findall(r"[\u4e00-\u9fa5]", value))
    english_count = len(re.findall(r"[A-Za-z]", value))
    return chinese_count >= english_count


def ensure_file_exists(file_path: Path) -> None:
    if not file_path.exists():
        raise FileNotFoundError(f"文件不存在：{file_path}")


def create_default_output_path(input_path: Path, collected_at: str, school: str = "", college: str = "") -> Path:
    if school or college:
        return build_school_college_output_path(aggregate_root_dir(collected_at), school, college, input_path.stem)
    return aggregate_root_dir(collected_at) / f"{build_output_stem(input_path.stem)}.csv"


def create_default_output_path_from_url(page_url: str, collected_at: str, school: str = "", college: str = "") -> Path:
    split = urlsplit(page_url)
    stem = Path(split.path).stem or "列表页输出"
    if school or college:
        return build_school_college_output_path(aggregate_root_dir(collected_at), school, college, stem)
    return aggregate_root_dir(collected_at) / f"{build_output_stem(stem)}.csv"


def ensure_csv_output_path(path: Path) -> Path:
    return path if path.suffix.lower() == ".csv" else path.with_suffix(".csv")


def create_default_log_path(output_path: Path) -> Path:
    return Path.cwd() / "日志" / f"{output_path.stem}_日志.json"


def create_preview_csv_path(output_path: Path) -> Path:
    return output_path.with_name(f"{output_path.stem}_预览.csv")


def create_legacy_xlsx_path(output_path: Path) -> Path:
    return output_path.with_suffix(".xlsx")


def build_school_college_output_path(base_dir: Path, school: str, college: str, fallback_name: str) -> Path:
    normalized_school = sanitize_path_component(school) or "未填写学校"
    normalized_college = sanitize_path_component(college) or sanitize_path_component(build_output_stem(fallback_name)) or "未命名学院"
    return base_dir / normalized_school / f"{normalized_college}.csv"


def build_output_stem(input_stem: str) -> str:
    stem = input_stem.strip()
    for suffix in ["_导入_Python", "_导入", "_导出", "导入_Python", "导入", "导出", "输出"]:
        if stem.endswith(suffix):
            stem = stem[: -len(suffix)].rstrip("_- ")
            break
    return f"{stem}输出"


def infer_default_org_fields(
    input_path: Path | None,
    list_page_url: str,
    profile_url: str,
    fallback_school: str,
    fallback_college: str,
) -> tuple[str, str]:
    school = normalize_whitespace(fallback_school)
    college = normalize_whitespace(fallback_college)
    source_url = list_page_url or profile_url
    page_school = ""
    page_college = ""
    if source_url:
        page_school, page_college = infer_org_fields_from_page(source_url)
    if page_school:
        school = page_school
    if page_college:
        college = page_college

    source_stem = ""
    if input_path:
        source_stem = clean_input_stem(input_path.stem)
    else:
        if source_url:
            source_stem = clean_input_stem(Path(urlsplit(source_url).path).stem)

    if not school:
        school = extract_school_name(source_stem)
    if not college:
        college = extract_college_name(source_stem, school)
    return school, college


def clean_input_stem(stem: str) -> str:
    cleaned = stem
    for token in ["官网", "网站", "欢迎您", "欢迎您！", "教师信息", "教师列表", "老师列表", "导入", "导出", "输出"]:
        cleaned = cleaned.replace(token, "")
    cleaned = re.sub(r"\(\d+\)$", "", cleaned)
    return re.sub(r"[-_]+", "-", cleaned).strip("- ")


@functools.lru_cache(maxsize=256)
def infer_org_fields_from_page(page_url: str) -> tuple[str, str]:
    try:
        html_content = fetch_with_timeout(page_url, 20)
    except Exception:
        return "", ""
    return infer_org_fields_from_html(page_url, html_content)


def infer_org_fields_from_html(page_url: str, html_content: str) -> tuple[str, str]:
    title_candidates: list[str] = []
    title_match = re.search(r"<title[^>]*>([\s\S]*?)</title>", html_content, re.IGNORECASE)
    if title_match:
        title_candidates.append(clean_intro_text(html_to_text(title_match.group(1))))
    page_title_match = re.search(
        r'<meta[^>]+name=["\']pageTitle["\'][^>]+content=["\']([^"\']+)["\']',
        html_content,
        re.IGNORECASE,
    )
    if page_title_match:
        title_candidates.append(clean_intro_text(html_to_text(page_title_match.group(1))))

    head_scope = clean_intro_text(html_to_text(html_content[:12000]))
    tail_scope = clean_intro_text(html_to_text(html_content[-6000:]))
    school = infer_school_name_from_url(page_url)
    college = ""
    for candidate in title_candidates + [tail_scope, head_scope]:
        if not school:
            school = extract_school_name(candidate)
        if not college:
            college = extract_college_name(candidate, school)
        if school and college:
            break
    return school, college


def infer_school_name_from_url(page_url: str) -> str:
    host = urlsplit(page_url).netloc.lower()
    domain_school_map = {
        "sust.edu.cn": "陕西科技大学",
        "xust.edu.cn": "西安科技大学",
        "xidian.edu.cn": "西安电子科技大学",
        "xaut.edu.cn": "西安理工大学",
        "xauat.edu.cn": "西安建筑科技大学",
        "xsyu.edu.cn": "西安石油大学",
    }
    for domain, school in domain_school_map.items():
        if host == domain or host.endswith(f".{domain}"):
            return school
    return ""


def extract_school_name(stem: str) -> str:
    school_matches = []
    for suffix in SCHOOL_SUFFIXES:
        pattern = re.compile(rf"([\u4e00-\u9fa5A-Za-z0-9·（）()]+?{re.escape(suffix)})")
        for match in pattern.finditer(stem):
            candidate = normalize_org_name(match.group(1))
            if looks_like_school(candidate):
                school_matches.append(candidate)

    if not school_matches:
        return ""
    school_matches.sort(key=len, reverse=True)
    return school_matches[0]


def extract_college_name(stem: str, school: str) -> str:
    candidates = []
    for part in split_filename_parts(stem):
        normalized = normalize_org_name(part)
        if not normalized:
            continue
        candidate = remove_school_prefix(normalized, school)
        if looks_like_college(candidate):
            candidates.append(candidate)

    for candidate in candidates:
        if any(candidate.endswith(suffix) for suffix in ("学院", "研究院")):
            return candidate
    return candidates[0] if candidates else ""


def split_filename_parts(stem: str) -> list[str]:
    parts = [normalize_org_name(part) for part in re.split(r"[-_]", stem)]
    return [part for part in parts if part]


def normalize_org_name(value: str) -> str:
    cleaned = re.sub(r"[（(].*?[）)]", "", value)
    cleaned = re.sub(r"\s+", "", cleaned)
    return cleaned.strip("-_ ")


def remove_school_prefix(candidate: str, school: str) -> str:
    if school and candidate.startswith(school):
        return candidate.removeprefix(school)
    return candidate


def infer_group_title_from_filename(file_stem: str) -> str:
    normalized = normalize_whitespace(file_stem)
    if "副教授" in normalized:
        return "副教授"
    if "讲师" in normalized:
        return "讲师"
    if "教授" in normalized:
        return "教授"
    return ""


def parse_teacher_list_page(
    page_url: str,
    default_school: str,
    default_college: str,
    log_entries: List[Dict[str, object]],
    limit: int | None = None,
) -> List[Dict[str, str]]:
    html_content = fetch_with_timeout(page_url, 20)
    section_html = extract_teacher_card_section(html_content) or html_content
    items = []
    for index, match in enumerate(TEACHER_CARD_PATTERN.finditer(section_html), start=1):
        title_text = clean_intro_text(html_to_text(match.group(2)))
        body_text = clean_intro_text(html_to_text(match.group(3)))
        merged_text = normalize_whitespace(f"{title_text}\n{body_text}")
        href = normalize_whitespace(match.group(1))
        profile_url = "" if href == "#" else urljoin(page_url, href)
        rough_name = clean_name(title_text)
        title = normalize_title(title_text) or extract_title(merged_text, merged_text, rough_name)
        name = extract_name_from_title_text(title_text, title)
        email = extract_email("", merged_text, merged_text, None)
        intro = build_card_intro(body_text)

        if not name:
            log_entries.append(
                {
                    "level": "warn",
                    "rowNumber": index,
                    "url": page_url,
                    "message": "列表页条目未识别到姓名，已跳过",
                }
            )
            continue

        items.append(
            {
                "row_number": str(index),
                "name": name,
                "profile_url": profile_url or page_url,
                "school": normalize_whitespace(default_school),
                "college": normalize_whitespace(default_college),
                "title": title,
                "email": email,
                "intro": intro,
            }
        )
        if limit and len(items) >= max(1, limit):
            break

    if not items:
        items = extract_teacher_links_from_list_page(
            page_url,
            html_content,
            default_school=default_school,
            default_college=default_college,
            limit=limit,
        )

    if not items:
        items = extract_vsb_teacher_widget_items(
            page_url,
            html_content,
            default_school=default_school,
            default_college=default_college,
            limit=limit,
        )

    if not items:
        raise ValueError(f"未从列表页识别到老师条目：{page_url}")

    return items


def parse_teacher_profile_page(
    page_url: str,
    default_school: str,
    default_college: str,
    log_entries: List[Dict[str, object]],
) -> List[Dict[str, str]]:
    html_content = fetch_with_timeout(page_url, 20)
    html_content = decode_tsites_encrypted_fields(html_content, page_url, 20)
    name = infer_name_from_profile_page(html_content, page_url)
    if not name:
        raise ValueError(f"未从教师主页识别到姓名：{page_url}")

    seed_record = {
        "row_number": "1",
        "name": name,
        "profile_url": page_url,
        "school": normalize_whitespace(default_school),
        "college": normalize_whitespace(default_college),
        "title": "",
        "email": "",
        "intro": "",
    }
    return [enrich_record(seed_record, 20, log_entries)]


def extract_teacher_links_from_list_page(
    page_url: str,
    html_content: str,
    default_school: str,
    default_college: str,
    limit: int | None = None,
) -> List[Dict[str, str]]:
    items: List[Dict[str, str]] = []
    seen_pairs: set[tuple[str, str]] = set()
    row_number = 1

    for href, inner_html in TEACHER_LIST_LINK_PATTERN.findall(html_content):
        raw_text = clean_intro_text(html_to_text(inner_html))
        compact_text = re.sub(r"\s+", "", raw_text)
        name = clean_name(compact_text)
        if not name or not looks_like_name(name):
            continue

        profile_url = urljoin(page_url, normalize_whitespace(href))
        pair_key = (name, profile_url)
        if pair_key in seen_pairs:
            continue
        seen_pairs.add(pair_key)

        items.append(
            {
                "row_number": str(row_number),
                "name": name,
                "profile_url": profile_url,
                "school": normalize_whitespace(default_school),
                "college": normalize_whitespace(default_college),
                "title": infer_title_from_name(compact_text),
                "email": "",
                "intro": "",
            }
        )
        row_number += 1
        if limit and len(items) >= max(1, limit):
            break

    return items


def extract_vsb_teacher_widget_items(
    page_url: str,
    html_content: str,
    default_school: str,
    default_college: str,
    limit: int | None = None,
) -> List[Dict[str, str]]:
    items: List[Dict[str, str]] = []
    seen_pairs: set[tuple[str, str]] = set()

    for widget_match in VSB_WIDGET_TITLE_PATTERN.finditer(html_content):
        prefix = widget_match.group(1)
        titles = parse_js_string_array(widget_match.group(2))
        ids_match = re.search(VSB_WIDGET_ID_PATTERN_TEMPLATE.format(prefix=re.escape(prefix)), html_content, re.IGNORECASE)
        if not ids_match:
            continue
        ids = parse_js_string_array(ids_match.group(1))
        summaries_match = re.search(
            VSB_WIDGET_SUMMARY_PATTERN_TEMPLATE.format(prefix=re.escape(prefix)),
            html_content,
            re.IGNORECASE,
        )
        summaries = parse_js_string_array(summaries_match.group(1)) if summaries_match else []
        hidden_links = {
            item_id: urljoin(page_url, normalize_whitespace(href))
            for item_id, href in re.findall(
                VSB_WIDGET_HIDDEN_LINK_PATTERN_TEMPLATE.format(prefix=re.escape(prefix)),
                html_content,
                re.IGNORECASE,
            )
        }

        for index, item_id in enumerate(ids):
            if index >= len(titles):
                continue
            name = clean_name(titles[index])
            if not name or not looks_like_name(name):
                continue
            profile_url = hidden_links.get(item_id, "")
            if not profile_url:
                continue
            pair_key = (name, profile_url)
            if pair_key in seen_pairs:
                continue
            seen_pairs.add(pair_key)
            summary = summaries[index] if index < len(summaries) else ""
            items.append(
                {
                    "row_number": str(len(items) + 1),
                    "name": name,
                    "profile_url": profile_url,
                    "school": normalize_whitespace(default_school),
                    "college": normalize_whitespace(default_college),
                    "title": infer_title_from_name(summary),
                    "email": "",
                    "intro": clean_intro_text(summary),
                }
            )
            if limit and len(items) >= max(1, limit):
                return items

    return items


def parse_js_string_array(raw_value: str) -> List[str]:
    try:
        parsed = ast.literal_eval(raw_value)
    except Exception:
        return []
    if not isinstance(parsed, list):
        return []
    return [normalize_whitespace(str(item)) for item in parsed]


def build_card_intro(value: str) -> str:
    cleaned = normalize_whitespace(value)
    cleaned = re.sub(r"邮箱\s*[:：]?\s*[A-Z0-9._%+-]+@[A-Z0-9.-]+?\.(?:com|cn|edu|net|org|gov)(?:\.[A-Z]{2,3})?", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"电子邮件\s*[:：]?\s*[A-Z0-9._%+-]+@[A-Z0-9.-]+?\.(?:com|cn|edu|net|org|gov)(?:\.[A-Z]{2,3})?", "", cleaned, flags=re.IGNORECASE)
    return clean_intro_text(cleaned)


def extract_name_from_title_text(title_text: str, title: str) -> str:
    cleaned = normalize_whitespace(title_text)
    if title:
        cleaned = cleaned.replace(title, " ")
    cleaned = re.sub(r"(硕导|博导|博士生导师|硕士生导师)$", "", cleaned).strip()
    return clean_name(cleaned)


def infer_name_from_profile_page(html_content: str, page_url: str) -> str:
    title_match = re.search(r"<title[^>]*>([\s\S]*?)</title>", html_content, re.IGNORECASE)
    candidates: list[str] = []
    if title_match:
        title_text = clean_intro_text(html_to_text(title_match.group(1)))
        candidates.extend(re.split(r"--|－|—|-", title_text))

    meta_candidates = re.findall(
        r'<meta[^>]+content=["\']([^"\']+)["\']',
        html_content,
        flags=re.IGNORECASE,
    )
    candidates.extend(meta_candidates[:3])
    candidates.append(Path(urlsplit(page_url).path).stem)

    for candidate in candidates:
        name = extract_name_candidate(candidate)
        if name:
            return name
    return ""


def extract_name_candidate(value: str) -> str:
    cleaned = normalize_whitespace(value)
    matches = re.findall(r"[\u4e00-\u9fa5·]{2,10}", cleaned)
    for match in matches:
        normalized = clean_name(match)
        if any(token in normalized for token in NAME_NOISE_TOKENS):
            continue
        if looks_like_name(normalized):
            return normalized
    return ""


def extract_teacher_card_section(html_content: str) -> str:
    match = TEACHER_CARD_SECTION_PATTERN.search(html_content)
    return match.group(1) if match else ""


def build_header_map(header_row: List[str], data_rows: List[List[str]]) -> Dict[str, int]:
    normalized_headers = [normalize_header(value) for value in header_row]
    mapping: Dict[str, int] = {}
    for field, aliases in DEFAULT_HEADERS.items():
        alias_set = {normalize_header(alias) for alias in aliases}
        mapping[field] = next(
            (
                i
                for i, header in enumerate(normalized_headers)
                if header in alias_set and header not in GENERIC_PLACEHOLDER_HEADERS
            ),
            -1,
        )

    inferred = infer_columns_from_samples(data_rows)
    if mapping["name"] < 0:
        mapping["name"] = inferred.get("name", -1)
    if mapping["profile_url"] < 0:
        mapping["profile_url"] = inferred.get("profile_url", -1)

    # 最后兜底：八爪鱼最常见是两列姓名+链接。
    if mapping["name"] < 0 and len(header_row) >= 1:
        mapping["name"] = 0
    if mapping["profile_url"] < 0 and len(header_row) >= 2:
        mapping["profile_url"] = 1
    return mapping


def infer_columns_from_samples(data_rows: List[List[str]]) -> Dict[str, int]:
    samples = [row for row in data_rows if any(normalize_whitespace(cell) for cell in row)][:10]
    if not samples:
        return {}

    max_cols = max(len(row) for row in samples)
    scores = [{"url": 0, "name": 0} for _ in range(max_cols)]

    for row in samples:
        for index in range(max_cols):
            cell = normalize_whitespace(row[index]) if index < len(row) else ""
            if not cell:
                continue
            if looks_like_url(cell):
                scores[index]["url"] += 1
            if looks_like_name(cell):
                scores[index]["name"] += 1

    profile_index = max(range(max_cols), key=lambda i: scores[i]["url"], default=-1)
    name_index = max(range(max_cols), key=lambda i: scores[i]["name"], default=-1)

    result = {}
    if profile_index >= 0 and scores[profile_index]["url"] > 0:
        result["profile_url"] = profile_index
    if name_index >= 0 and scores[name_index]["name"] > 0:
        result["name"] = name_index
    return result


def normalize_seed_row(
    row: List[str],
    header_map: Dict[str, int],
    index: int,
    default_school: str = "",
    default_college: str = "",
    fallback_title: str = "",
) -> Dict[str, str]:
    raw_name = get_field(row, header_map["name"])
    inferred_name_title = infer_title_from_name(raw_name)
    return {
        "row_number": str(index + 2),
        "name": raw_name,
        "profile_url": get_field(row, header_map["profile_url"]),
        "school": get_field(row, header_map["school"]) or normalize_whitespace(default_school),
        "college": get_field(row, header_map["college"]) or normalize_whitespace(default_college),
        "title": get_field(row, header_map["title"]) or inferred_name_title or normalize_whitespace(fallback_title),
        "email": get_field(row, header_map["email"]),
        "intro": get_field(row, header_map["intro"]),
    }


def expand_seed_rows(
    row: List[str],
    header_map: Dict[str, int],
    index: int,
    default_school: str = "",
    default_college: str = "",
    fallback_title: str = "",
) -> List[Dict[str, str]]:
    expanded_urls = expand_multi_url_row(
        row,
        index,
        default_school=default_school,
        default_college=default_college,
        fallback_title=fallback_title,
    )
    if expanded_urls:
        return expanded_urls

    expanded = expand_multi_teacher_row(
        row,
        index,
        default_school=default_school,
        default_college=default_college,
        fallback_title=fallback_title,
    )
    if expanded:
        return expanded
    return [
        normalize_seed_row(
            row,
            header_map,
            index,
            default_school=default_school,
            default_college=default_college,
            fallback_title=fallback_title,
        )
    ]


def expand_multi_url_row(
    row: List[str],
    index: int,
    default_school: str = "",
    default_college: str = "",
    fallback_title: str = "",
) -> List[Dict[str, str]]:
    normalized_cells = [normalize_whitespace(cell) for cell in row]
    urls = []
    seen_urls = set()
    for cell in normalized_cells:
        if not looks_like_url(cell):
            continue
        if cell in seen_urls:
            continue
        seen_urls.add(cell)
        urls.append(cell)

    if len(urls) < 2:
        return []

    return [
        {
            "row_number": str(index + 2),
            "name": "",
            "profile_url": profile_url,
            "school": normalize_whitespace(default_school),
            "college": normalize_whitespace(default_college),
            "title": normalize_whitespace(fallback_title),
            "email": "",
            "intro": "",
        }
        for profile_url in urls
    ]


def expand_multi_teacher_row(
    row: List[str],
    index: int,
    default_school: str = "",
    default_college: str = "",
    fallback_title: str = "",
) -> List[Dict[str, str]]:
    normalized_cells = [normalize_whitespace(cell) for cell in row]
    url_indices = [cell_index for cell_index, cell in enumerate(normalized_cells) if looks_like_url(cell)]
    name_indices = [cell_index for cell_index, cell in enumerate(normalized_cells) if looks_like_name(cell)]
    if len(url_indices) < 2 or len(name_indices) < 2:
        return []

    orientation = infer_row_pair_orientation(normalized_cells)
    records: List[Dict[str, str]] = []
    seen_pairs: set[tuple[str, str]] = set()

    for url_index in url_indices:
        profile_url = normalized_cells[url_index]
        name = ""
        candidate_indices: List[int]
        if orientation == "name_url":
            candidate_indices = [url_index - 1, url_index + 1]
        elif orientation == "url_name":
            candidate_indices = [url_index + 1, url_index - 1]
        else:
            candidate_indices = [url_index + 1, url_index - 1]

        for candidate_index in candidate_indices:
            if 0 <= candidate_index < len(normalized_cells):
                candidate_value = normalized_cells[candidate_index]
                if looks_like_name(candidate_value):
                    name = candidate_value
                    break

        if not name:
            continue

        pair_key = (name, profile_url)
        if pair_key in seen_pairs:
            continue
        seen_pairs.add(pair_key)
        records.append(
            {
                "row_number": str(index + 2),
                "name": name,
                "profile_url": profile_url,
                "school": normalize_whitespace(default_school),
                "college": normalize_whitespace(default_college),
                "title": infer_title_from_name(name) or normalize_whitespace(fallback_title),
                "email": "",
                "intro": "",
            }
        )

    return records


def infer_row_pair_orientation(cells: List[str]) -> str:
    name_url_pairs = 0
    url_name_pairs = 0
    for index in range(len(cells) - 1):
        left = cells[index]
        right = cells[index + 1]
        if looks_like_name(left) and looks_like_url(right):
            name_url_pairs += 1
        if looks_like_url(left) and looks_like_name(right):
            url_name_pairs += 1
    if name_url_pairs > url_name_pairs:
        return "name_url"
    if url_name_pairs > name_url_pairs:
        return "url_name"
    return "mixed"


def enrich_record(record: Dict[str, str], timeout_seconds: int, log_entries: List[Dict[str, object]]) -> Dict[str, str]:
    result = dict(record)

    if not record["profile_url"]:
        log_entries.append(
            {
                "level": "warn",
                "rowNumber": int(record["row_number"]),
                "name": record["name"],
                "message": "缺少主页链接，已跳过详情抓取",
            }
        )
        return result

    try:
        html_content = fetch_with_timeout(record["profile_url"], timeout_seconds)
        html_content = decode_tsites_encrypted_fields(html_content, record["profile_url"], timeout_seconds)
        page_school, page_college = infer_org_fields_from_html(record["profile_url"], html_content)
        if page_school:
            result["school"] = page_school
        if page_college:
            result["college"] = page_college
        if not normalize_whitespace(result.get("name", "")) or looks_like_url(result.get("name", "")):
            inferred_name = infer_name_from_profile_page(html_content, record["profile_url"])
            if inferred_name:
                result["name"] = inferred_name
        text = html_to_text(html_content)
        context = build_teacher_context(text, result["name"])
        image_ocr_text_cache: List[str | None] = [None]

        def get_image_ocr_text() -> str:
            if image_ocr_text_cache[0] is None:
                image_ocr_text_cache[0] = extract_image_text(html_content) if should_run_ocr(html_content, text) else ""
            return image_ocr_text_cache[0] or ""

        if not result["email"]:
            result["email"] = extract_email(html_content, text, context, get_image_ocr_text)
        if not result["title"]:
            result["title"] = extract_title(text, context, result["name"], get_image_ocr_text)
        if not result["intro"]:
            result["intro"] = extract_intro(html_content, text)

        if not result["email"]:
            log_entries.append(
                {
                    "level": "info",
                    "rowNumber": int(record["row_number"]),
                    "name": record["name"],
                    "url": record["profile_url"],
                    "message": "未提取到邮箱",
                }
            )
        if not result["title"]:
            log_entries.append(
                {
                    "level": "info",
                    "rowNumber": int(record["row_number"]),
                    "name": record["name"],
                    "url": record["profile_url"],
                    "message": "未提取到职称",
                }
            )
    except Exception as error:
        log_entries.append(
            {
                "level": "error",
                "rowNumber": int(record["row_number"]),
                "name": record["name"],
                "url": record["profile_url"],
                "message": str(error),
            }
        )

    return result


def fetch_with_timeout(url: str, timeout_seconds: int) -> str:
    normalized_url = urldefrag(url).url or url
    return fetch_with_timeout_cached(normalized_url, timeout_seconds)


@functools.lru_cache(maxsize=512)
def fetch_with_timeout_cached(url: str, timeout_seconds: int) -> str:
    request = Request(url, headers=build_browser_like_headers(url))
    last_error: Exception | None = None
    for attempt in range(3):
        try:
            with urlopen(request, timeout=timeout_seconds) as response:
                charset = response.headers.get_content_charset() or "utf-8"
                return response.read().decode(charset, errors="ignore")
        except HTTPError as error:
            raise RuntimeError(f"请求失败：HTTP {error.code}") from error
        except URLError as error:
            last_error = error
            if attempt < 2:
                time.sleep(1 + attempt)
                continue
            raise RuntimeError(f"请求失败：{error.reason}") from error
    raise RuntimeError(f"请求失败：{last_error}") if last_error else RuntimeError("请求失败：未知错误")


def build_browser_like_headers(url: str) -> dict[str, str]:
    split = urlsplit(url)
    site_root = f"{split.scheme}://{split.netloc}/" if split.scheme and split.netloc else url
    return {
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/137.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Cache-Control": "max-age=0",
        "Pragma": "no-cache",
        "Referer": site_root,
        "Upgrade-Insecure-Requests": "1",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
    }


def decode_tsites_encrypted_fields(html_content: str, page_url: str, timeout_seconds: int) -> str:
    matches = list(TSITES_ENCRYPT_SPAN_PATTERN.finditer(html_content))
    if not matches:
        return html_content

    base_url = build_site_root_url(page_url)
    if not base_url:
        return html_content

    replaced_html = html_content
    for match in reversed(matches):
        field_id = normalize_whitespace(match.group(1))
        encrypted_content = normalize_whitespace(decode_html(match.group(2)))
        if not field_id or not encrypted_content:
            continue
        decoded_content = decode_tsites_field(base_url, field_id, encrypted_content, timeout_seconds)
        if not decoded_content:
            continue
        new_span = re.sub(r'style=["\'][^"\']*display\s*:\s*none;?[^"\']*["\']', "", match.group(0), flags=re.IGNORECASE)
        new_span = re.sub(r'_tsites_encrypt_field=["\']_tsites_encrypt_field["\']', "", new_span, flags=re.IGNORECASE)
        new_span = re.sub(rf">{re.escape(match.group(2))}</span>$", f">{decoded_content}</span>", new_span, flags=re.IGNORECASE)
        replaced_html = replaced_html[: match.start()] + new_span + replaced_html[match.end() :]
    return replaced_html


def build_site_root_url(page_url: str) -> str:
    split = urlsplit(page_url)
    if not split.scheme or not split.netloc:
        return ""
    return f"{split.scheme}://{split.netloc}"


@functools.lru_cache(maxsize=2048)
def decode_tsites_field(base_url: str, field_id: str, encrypted_content: str, timeout_seconds: int) -> str:
    endpoint = urljoin(base_url, "/system/resource/tsites/tsitesencrypt.jsp")
    query = urlencode(
        {
            "id": field_id,
            "content": encrypted_content,
            "mode": "8",
        }
    )
    request = Request(
        f"{endpoint}?{query}",
        headers={"User-Agent": "Mozilla/5.0 Codex Teacher Pipeline/1.0"},
    )
    for attempt in range(3):
        try:
            with urlopen(request, timeout=timeout_seconds) as response:
                charset = response.headers.get_content_charset() or "utf-8"
                payload = json.loads(response.read().decode(charset, errors="ignore"))
            return normalize_whitespace(decode_html(str(payload.get("content", ""))))
        except Exception:
            if attempt < 2:
                time.sleep(1 + attempt)
                continue
            return ""


def extract_email(
    html_content: str,
    text: str,
    context: str = "",
    image_ocr_text_supplier=None,
) -> str:
    scored_candidates: List[tuple[int, str]] = []

    labeled_sources = [context, text, html_content]
    labeled_email = extract_labeled_email(labeled_sources)
    if labeled_email and is_reliable_email_candidate(labeled_email, labeled_sources):
        return labeled_email

    for source_name, source in [("context", context), ("text", text)]:
        if not source:
            continue
        normalized = normalize_email_like_text(source)
        scored_candidates.extend(score_email_candidates(normalized, source_name))

    best_text_email = pick_best_scored_email(scored_candidates, min_score=70)
    if best_text_email:
        return best_text_email

    normalized_html = normalize_email_like_text(html_content)
    scored_candidates.extend(score_email_candidates(normalized_html, "html"))

    ocr_email = extract_email_from_images(
        html_content,
        image_ocr_text_supplier() if callable(image_ocr_text_supplier) else "",
    )
    if ocr_email:
        scored_candidates.append((40, ocr_email))
    return pick_best_scored_email(scored_candidates)


def extract_title(
    text: str,
    context: str = "",
    name: str = "",
    image_ocr_text_supplier=None,
) -> str:
    sources = [normalize_whitespace(source) for source in [context, text] if source]
    for source in sources:
        labeled_title = extract_labeled_title(source)
        if labeled_title:
            if is_valid_title_value(labeled_title) and labeled_title != "教师":
                return labeled_title
        nearby_title = extract_title_near_name(source, name)
        if nearby_title:
            if is_valid_title_value(nearby_title) and nearby_title != "教师":
                return nearby_title
        matched_title = find_title_in_text(source)
        if matched_title:
            if is_valid_title_value(matched_title) and matched_title != "教师":
                return matched_title

    if callable(image_ocr_text_supplier):
        image_ocr_text = normalize_whitespace(image_ocr_text_supplier())
        if image_ocr_text:
            labeled_title = extract_labeled_title(image_ocr_text)
            if labeled_title:
                if is_valid_title_value(labeled_title) and labeled_title != "教师":
                    return labeled_title
            nearby_title = extract_title_near_name(image_ocr_text, name)
            if nearby_title:
                if is_valid_title_value(nearby_title) and nearby_title != "教师":
                    return nearby_title
            matched_title = find_title_in_text(image_ocr_text)
            if matched_title:
                if is_valid_title_value(matched_title) and matched_title != "教师":
                    return matched_title
    return ""


def extract_title_near_name(source: str, name: str) -> str:
    clean_teacher_name = clean_name(name)
    if not clean_teacher_name:
        return ""
    name_index = source.find(clean_teacher_name)
    if name_index < 0:
        return ""

    name_start = name_index
    name_end = name_index + len(clean_teacher_name)
    best_match = ""
    best_distance = None

    for start, end, canonical in find_title_matches(source):
        if end <= name_start:
            distance = name_start - end
        elif start >= name_end:
            distance = start - name_end
        else:
            distance = 0

        if distance > 40:
            continue

        if best_distance is None or distance < best_distance:
            best_distance = distance
            best_match = canonical
        elif distance == best_distance and title_priority(canonical) < title_priority(best_match):
            best_match = canonical

    return best_match


def extract_intro(html_content: str, text: str) -> str:
    section_intro = extract_section_intro(html_content)
    if section_intro and is_valid_intro_text(section_intro):
        return truncate_text(section_intro, 200)

    meta_description = match_meta_description(html_content)
    if meta_description and is_valid_intro_text(meta_description):
        return truncate_text(meta_description, 200)

    lines = [
        normalize_whitespace(line)
        for line in text.split("\n")
        if normalize_whitespace(line)
    ]
    lines = [line for line in lines if len(line) >= 20 and not looks_like_navigation(line) and is_valid_intro_text(line)]
    return truncate_text(lines[0] if lines else "", 200)


def extract_section_intro(html_content: str) -> str:
    preferred_titles = ("个人简介", "简介", "Biography", "Profile", "个人情况")
    for title, block_html in extract_named_sections(html_content):
        normalized_title = normalize_whitespace(title)
        if normalized_title not in preferred_titles:
            continue
        block_text = html_to_text(block_html)
        block_text = clean_intro_text(block_text)
        if is_valid_intro_text(block_text):
            return block_text
    return ""


def extract_named_sections(html_content: str) -> List[tuple[str, str]]:
    sections: List[tuple[str, str]] = []
    for match in SECTION_BLOCK_PATTERN.finditer(html_content):
        title = normalize_whitespace(decode_html(match.group(1)))
        block_html = match.group(2)
        sections.append((title, block_html))
    return sections


def match_meta_description(html_content: str) -> str:
    match = re.search(
        r'<meta[^>]+name=["\']description["\'][^>]+content=["\']([^"\']+)["\']',
        html_content,
        re.IGNORECASE,
    )
    return normalize_whitespace(decode_html(match.group(1))) if match else ""


def looks_like_navigation(line: str) -> bool:
    bad_tokens = [
        "首页",
        "当前位置",
        "联系我们",
        "版权所有",
        "学院概况",
        "人才培养",
        "科学研究",
        "招生就业",
        "师资队伍",
        "关闭",
        "打印",
    ]
    return any(token in line for token in bad_tokens)


def to_template_row(record: Dict[str, str], index: int, collector: str, collected_at: str) -> List[str]:
    normalized_title = normalize_title(record["title"])
    if not normalized_title and normalize_whitespace(record["email"]):
        normalized_title = "教授"
    return [
        str(index),
        compact_csv_value(clean_name(record["name"])),
        compact_csv_value(normalized_title),
        compact_csv_value(record["email"] or ""),
        compact_csv_value(record["school"] or ""),
        compact_csv_value(record["college"] or ""),
        compact_csv_value(collector),
        compact_csv_value(collected_at),
        compact_csv_value(truncate_text(normalize_whitespace(record["intro"] or ""), 200)),
        compact_csv_value(record["profile_url"] or ""),
    ]


def is_valid_template_row(row: List[str]) -> bool:
    for offset, _ in enumerate(ALL_REQUIRED_TEMPLATE_HEADERS, start=1):
        if not normalize_whitespace(row[offset]):
            return False
    return True


def renumber_template_rows(rows: List[List[str]]) -> List[List[str]]:
    renumbered: List[List[str]] = []
    for index, row in enumerate(rows, start=1):
        updated = list(row)
        updated[0] = str(index)
        renumbered.append(updated)
    return renumbered


def strip_shared_email_rows(rows: List[List[str]]) -> List[List[str]]:
    email_name_groups: dict[tuple[str, str, str], set[str]] = {}
    for row in rows:
        if len(row) < len(TEMPLATE_HEADERS):
            continue
        email = normalize_whitespace(row[3]).lower()
        if not email:
            continue
        school = normalize_whitespace(row[4])
        college = normalize_whitespace(row[5])
        name = clean_name(normalize_whitespace(row[1]))
        if not school or not college or not name:
            continue
        email_name_groups.setdefault((school, college, email), set()).add(name)

    cleaned_rows: List[List[str]] = []
    for row in rows:
        normalized_row = list(row)
        if len(normalized_row) >= len(TEMPLATE_HEADERS):
            email = normalize_whitespace(normalized_row[3]).lower()
            school = normalize_whitespace(normalized_row[4])
            college = normalize_whitespace(normalized_row[5])
            shared_names = email_name_groups.get((school, college, email), set())
            if email and len(shared_names) >= 2:
                normalized_row[3] = ""
        cleaned_rows.append(normalized_row)
    return cleaned_rows


def sanitize_template_rows(rows: List[List[str]], *, keep_invalid: bool = False) -> List[List[str]]:
    sanitized_rows = strip_shared_email_rows(rows)
    normalized_rows: List[List[str]] = []
    for row in sanitized_rows:
        if len(row) < len(TEMPLATE_HEADERS):
            continue
        normalized_row = [normalize_whitespace(str(cell)) for cell in row[: len(TEMPLATE_HEADERS)]]
        normalized_row = normalize_template_row_name(normalized_row)
        if keep_invalid or is_valid_template_row(normalized_row):
            normalized_rows.append(normalized_row)
    return normalized_rows


def merge_template_output_rows(existing_rows: List[List[str]], new_rows: List[List[str]]) -> List[List[str]]:
    merged_by_key: dict[tuple[str, ...], list[str]] = {}
    for row in sanitize_template_rows(existing_rows + new_rows):
        normalized_row = list(row)
        unique_key = build_template_row_identity(normalized_row)
        existing_row = merged_by_key.get(unique_key)
        merged_by_key[unique_key] = (
            choose_preferred_template_row(existing_row, normalized_row)
            if existing_row
            else normalized_row
        )
    merged_rows = sanitize_template_rows(list(merged_by_key.values()))
    return renumber_template_rows(merged_rows)


def load_existing_output_rows(output_path: Path) -> List[List[str]]:
    if output_path.exists() and output_path.suffix.lower() == ".csv":
        return sanitize_template_rows(read_csv_output_rows(output_path), keep_invalid=True)

    legacy_xlsx_path = create_legacy_xlsx_path(output_path)
    if legacy_xlsx_path.exists():
        try:
            headers, rows = read_first_sheet(legacy_xlsx_path)
        except Exception:
            return []
        if [normalize_whitespace(value) for value in headers[: len(TEMPLATE_HEADERS)]] != TEMPLATE_HEADERS:
            return []
        return sanitize_template_rows(normalize_output_rows(rows), keep_invalid=True)
    return []


def read_csv_output_rows(path: Path) -> List[List[str]]:
    rows: List[List[str]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        next(reader, None)
        for row in reader:
            if not any(normalize_whitespace(cell) for cell in row):
                continue
            padded = list(row[: len(TEMPLATE_HEADERS)])
            while len(padded) < len(TEMPLATE_HEADERS):
                padded.append("")
            rows.append(normalize_template_row_name(padded))
    return rows


def normalize_output_rows(rows: List[List[str]]) -> List[List[str]]:
    normalized_rows: List[List[str]] = []
    for row in rows:
        if not any(normalize_whitespace(cell) for cell in row):
            continue
        padded = list(row[: len(TEMPLATE_HEADERS)])
        while len(padded) < len(TEMPLATE_HEADERS):
            padded.append("")
        normalized_row = [normalize_whitespace(str(cell)) for cell in padded]
        normalized_rows.append(normalize_template_row_name(normalized_row))
    return normalized_rows


def normalize_template_row_name(row: List[str]) -> List[str]:
    normalized_row = list(row)
    if len(normalized_row) > 1:
        normalized_row[1] = clean_name(normalized_row[1])
    return normalized_row


def infer_output_school(output_rows: List[List[str]], fallback_school: str) -> str:
    normalized_fallback = normalize_whitespace(fallback_school)
    if normalized_fallback:
        return normalized_fallback
    for row in output_rows:
        if len(row) > 4 and normalize_whitespace(row[4]):
            return normalize_whitespace(row[4])
    return "未填写学校"


def infer_output_college(
    output_rows: List[List[str]],
    fallback_college: str,
    input_path: Path | None,
    list_page_url: str,
) -> str:
    normalized_fallback = normalize_whitespace(fallback_college)
    if normalized_fallback:
        return normalized_fallback
    for row in output_rows:
        if len(row) > 5 and normalize_whitespace(row[5]):
            return normalize_whitespace(row[5])
    if input_path:
        return infer_college_name_from_stem(input_path.stem)
    if list_page_url:
        return infer_college_name_from_stem(Path(urlsplit(list_page_url).path).stem)
    return "未填写学院"


def infer_college_name_from_stem(stem: str) -> str:
    normalized = normalize_whitespace(stem)
    parts = [part for part in re.split(r"[-_—\s]+", normalized) if part]
    for part in parts:
        if any(token in part for token in ["学院", "学部", "研究院", "中心", "系"]):
            return part
    return normalized or "未填写学院"


def create_output_manifest_path(output_path: Path, collected_at: str) -> Path:
    relative_output = output_path.as_posix()
    digest = hashlib.sha1(relative_output.encode("utf-8")).hexdigest()[:12]
    filename = sanitize_filename(f"{output_path.stem}_{digest}.json")
    return aggregate_root_dir(collected_at) / AGGREGATE_RECORDS_DIRNAME / filename


def write_output_manifest(
    manifest_path: Path,
    *,
    input_path: Path | None,
    list_page_url: str,
    output_path: Path,
    collector: str,
    collected_at: str,
    school: str,
    college: str,
    rows: List[List[str]],
) -> None:
    payload = {
        "createdAt": dt.datetime.now().isoformat(),
        "inputPath": str(input_path) if input_path else "",
        "listPageUrl": list_page_url,
        "outputPath": str(output_path),
        "collector": collector,
        "collectedAt": collected_at,
        "school": school,
        "college": college,
        "headers": TEMPLATE_HEADERS,
        "rows": rows,
    }
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    with manifest_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def update_aggregate_workbooks(collected_at: str) -> None:
    manifests = load_output_manifests(collected_at)
    root_dir = aggregate_root_dir(collected_at)

    # 先把最终 CSV 重建到稳定状态，再只用最终 CSV 生成汇总 XLSX，
    # 确保手动修过的 CSV 会反映到提交用的汇总文件里。
    rebuild_college_output_files(root_dir, manifests)

    college_groups: dict[str, list[list[str]]] = {}
    college_seen: dict[str, dict[tuple[str, ...], list[str]]] = {}
    append_existing_output_rows_to_groups(root_dir, college_groups, college_seen)
    if not college_groups:
        remove_path_if_exists(root_dir / "按院系汇总.xlsx")
        remove_path_if_exists(root_dir / "按院校汇总.xlsx")
        remove_path_if_exists(root_dir / "按院校汇总_预览")
        remove_path_if_exists(root_dir / "按院系汇总_预览")
        return
    aggregate_rows = build_aggregate_rows(college_groups)
    write_workbook(
        root_dir / "按院系汇总.xlsx",
        [{"name": "汇总", "headers": TEMPLATE_HEADERS, "rows": aggregate_rows}],
    )
    remove_path_if_exists(root_dir / "按院校汇总.xlsx")
    remove_path_if_exists(root_dir / "按院校汇总_预览")
    remove_path_if_exists(root_dir / "按院系汇总_预览")


def append_manifest_rows_to_groups(
    manifests: list[dict[str, object]],
    college_groups: dict[str, list[list[str]]],
    college_seen: dict[str, dict[tuple[str, ...], list[str]]],
) -> None:
    for manifest in manifests:
        rows = manifest.get("rows", [])
        if not isinstance(rows, list):
            continue
        for raw_row in rows:
            if not isinstance(raw_row, list) or len(raw_row) < len(TEMPLATE_HEADERS):
                continue
            row = [normalize_whitespace(str(cell)) for cell in raw_row[: len(TEMPLATE_HEADERS)]]
            row = normalize_template_row_name(row)
            if not is_valid_template_row(row):
                continue

            school = row[4] or normalize_whitespace(str(manifest.get("school", ""))) or "未填写学校"
            college = row[5] or normalize_whitespace(str(manifest.get("college", ""))) or "未填写学院"
            college_key = f"{school}-{college}"
            unique_key = build_template_row_identity(row)

            append_unique_group_row(college_groups, college_seen, college_key, row, unique_key)


def append_existing_output_rows_to_groups(
    root_dir: Path,
    college_groups: dict[str, list[list[str]]],
    college_seen: dict[str, dict[tuple[str, ...], list[str]]],
) -> None:
    for output_path in iter_mergeable_output_files(root_dir):
        try:
            headers, rows = read_output_file_rows(output_path)
        except Exception:
            continue
        if headers[: len(TEMPLATE_HEADERS)] != TEMPLATE_HEADERS:
            continue
        for raw_row in rows:
            if len(raw_row) < len(TEMPLATE_HEADERS):
                continue
            row = [normalize_whitespace(str(cell)) for cell in raw_row[: len(TEMPLATE_HEADERS)]]
            row = normalize_template_row_name(row)
            if not is_valid_template_row(row):
                continue
            school = row[4] or "未填写学校"
            college = row[5] or "未填写学院"
            if is_placeholder_school(school) or is_placeholder_college(college):
                continue
            college_key = f"{school}-{college}"
            unique_key = build_template_row_identity(row)
            append_unique_group_row(college_groups, college_seen, college_key, row, unique_key)


def rebuild_college_output_files(root_dir: Path, manifests: list[dict[str, object]]) -> None:
    preferred_rows = collect_preferred_rows_from_manifests(manifests)
    for unique_key, row in collect_preferred_rows_from_output_files(root_dir).items():
        preferred_rows[unique_key] = choose_preferred_template_row(preferred_rows.get(unique_key), row)

    grouped_rows: dict[Path, list[list[str]]] = {}
    for row in sanitize_template_rows(list(preferred_rows.values())):
        school = row[4] or "未填写学校"
        college = row[5] or "未填写学院"
        if is_placeholder_school(school) or is_placeholder_college(college):
            continue
        target_path = build_school_college_output_path(root_dir, school, college, college).resolve()
        grouped_rows.setdefault(target_path, []).append(list(row))

    for target_path, rows in grouped_rows.items():
        ordered_rows = renumber_template_rows(sanitize_template_rows(rows))
        target_path.parent.mkdir(parents=True, exist_ok=True)
        write_csv(target_path, TEMPLATE_HEADERS, ordered_rows)
        remove_path_if_exists(create_legacy_xlsx_path(target_path))
        remove_path_if_exists(create_preview_csv_path(target_path))

    valid_targets = {path.resolve() for path in grouped_rows}
    for legacy_path in iter_mergeable_output_files(root_dir):
        if legacy_path.resolve() in valid_targets:
            continue
        remove_path_if_exists(legacy_path)
        remove_path_if_exists(create_preview_csv_path(legacy_path))
        if legacy_path.suffix.lower() == ".csv":
            remove_path_if_exists(create_legacy_xlsx_path(legacy_path))


def collect_preferred_rows_from_manifests(manifests: list[dict[str, object]]) -> dict[tuple[str, ...], list[str]]:
    preferred_rows: dict[tuple[str, ...], list[str]] = {}
    for manifest in manifests:
        rows = manifest.get("rows", [])
        if not isinstance(rows, list):
            continue

        manifest_school = normalize_whitespace(str(manifest.get("school", "")))
        manifest_college = normalize_whitespace(str(manifest.get("college", "")))
        for raw_row in rows:
            row = normalize_template_row(raw_row, manifest_school, manifest_college)
            if row is None:
                continue
            unique_key = build_template_row_identity(row)
            preferred_rows[unique_key] = choose_preferred_template_row(preferred_rows.get(unique_key), row)
    return preferred_rows


def collect_preferred_rows_from_output_files(root_dir: Path) -> dict[tuple[str, ...], list[str]]:
    preferred_rows: dict[tuple[str, ...], list[str]] = {}
    for output_path in iter_mergeable_output_files(root_dir):
        try:
            headers, rows = read_output_file_rows(output_path)
        except Exception:
            continue
        if headers[: len(TEMPLATE_HEADERS)] != TEMPLATE_HEADERS:
            continue
        for raw_row in rows:
            row = normalize_template_row(raw_row)
            if row is None:
                continue
            unique_key = build_template_row_identity(row)
            preferred_rows[unique_key] = choose_preferred_template_row(preferred_rows.get(unique_key), row)
    return preferred_rows


def normalize_template_row(
    raw_row: object,
    fallback_school: str = "",
    fallback_college: str = "",
) -> list[str] | None:
    if not isinstance(raw_row, list) or len(raw_row) < len(TEMPLATE_HEADERS):
        return None
    row = [normalize_whitespace(str(cell)) for cell in raw_row[: len(TEMPLATE_HEADERS)]]
    row = normalize_template_row_name(row)
    if fallback_school and is_placeholder_school(row[4]):
        row[4] = fallback_school
    if fallback_college and is_placeholder_college(row[5]):
        row[5] = fallback_college
    sanitized_rows = sanitize_template_rows([row])
    return sanitized_rows[0] if sanitized_rows else None


def remove_path_if_exists(path: Path) -> None:
    if path.is_dir():
        shutil.rmtree(path, ignore_errors=True)
        return
    if path.exists():
        path.unlink()


def is_path_within_root(path: Path, root_dir: Path) -> bool:
    try:
        path.relative_to(root_dir)
        return True
    except ValueError:
        return False


def iter_mergeable_output_files(root_dir: Path):
    for path in sorted(root_dir.rglob("*")):
        if not path.is_file():
            continue
        if path.suffix.lower() not in {".csv", ".xlsx"}:
            continue
        if path.name == "按院系汇总.xlsx" or path.name.endswith("_预览.csv"):
            continue
        if AGGREGATE_RECORDS_DIRNAME in path.parts:
            continue
        yield path


def read_output_file_rows(path: Path) -> tuple[list[str], list[list[str]]]:
    if path.suffix.lower() == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            rows = list(reader)
        if not rows:
            return [], []
        headers = [normalize_whitespace(value) for value in rows[0][: len(TEMPLATE_HEADERS)]]
        data_rows = normalize_output_rows(rows[1:])
        return headers, data_rows
    headers, rows = read_first_sheet(path)
    normalized_headers = [normalize_whitespace(value) for value in headers[: len(TEMPLATE_HEADERS)]]
    return normalized_headers, normalize_output_rows(rows)


def append_unique_group_row(
    group_map: dict[object, list[list[str]]],
    seen_map: dict[object, dict[tuple[str, ...], list[str]]],
    group_key: object,
    row: list[str],
    unique_key: tuple[str, ...],
) -> None:
    seen = seen_map.setdefault(group_key, {})
    rows = group_map.setdefault(group_key, [])
    existing_row = seen.get(unique_key)
    if existing_row is None:
        copied = list(row)
        seen[unique_key] = copied
        rows.append(copied)
        return
    preferred = choose_preferred_template_row(existing_row, row)
    if preferred is existing_row:
        return
    existing_row[:] = preferred


def build_template_row_identity(row: list[str]) -> tuple[str, ...]:
    profile_url = normalize_whitespace(row[9]) if len(row) > 9 else ""
    name = clean_name(normalize_whitespace(row[1]) if len(row) > 1 else "")
    school = normalize_whitespace(row[4]) if len(row) > 4 else ""
    email = normalize_whitespace(row[3]).lower() if len(row) > 3 else ""
    if profile_url:
        return ("url", profile_url.lower())
    if email:
        return ("email", school, email)
    if name:
        return ("name", school, name)
    college = normalize_whitespace(row[5]) if len(row) > 5 else ""
    return ("fallback", name, school, college)


def choose_preferred_template_row(left: list[str] | None, right: list[str]) -> list[str]:
    if left is None:
        return list(right)
    left_score = score_template_row(left)
    right_score = score_template_row(right)
    if right_score > left_score:
        return list(right)
    return left


def score_template_row(row: list[str]) -> tuple[int, int, int, int, int, int, int, int, int]:
    non_empty = sum(1 for cell in row[1:] if normalize_whitespace(cell))
    name = normalize_whitespace(row[1]) if len(row) > 1 else ""
    title = normalize_whitespace(row[2]) if len(row) > 2 else ""
    email = normalize_whitespace(row[3]) if len(row) > 3 else ""
    school = normalize_whitespace(row[4]) if len(row) > 4 else ""
    college = normalize_whitespace(row[5]) if len(row) > 5 else ""
    intro = normalize_whitespace(row[8]) if len(row) > 8 else ""
    normalized_name = clean_name(name)
    intro_name_score = 1 if normalized_name and normalized_name in normalize_person_name(intro) else 0
    clean_name_score = 1 if normalized_name and normalized_name == name else 0
    name_score = score_name_quality(name)
    title_score = 1 if title and title != "教师" and is_valid_title_value(title) else 0
    email_score = 1 if email else 0
    school_score = 1 if school and not is_placeholder_school(school) and looks_like_school(school) else 0
    college_score = 1 if college and not is_placeholder_college(college) and looks_like_college(college) else 0
    intro_score = min(len(intro), 200)
    return (
        school_score,
        college_score,
        non_empty,
        intro_name_score,
        clean_name_score,
        name_score,
        title_score,
        email_score,
        intro_score,
    )


def score_name_quality(name: str) -> int:
    normalized = clean_name(name)
    if not normalized:
        return 0
    if any(token in normalized for token in NAME_NOISE_TOKENS):
        return 0
    if looks_like_name(normalized):
        return 2
    if re.search(r"[\u4e00-\u9fa5]", normalized):
        return 1
    return 0


def build_group_sheets(group_map: dict[str, list[list[str]]]) -> list[dict[str, object]]:
    sheets = []
    for sheet_name in sorted(group_map):
        rows = renumber_template_rows(group_map[sheet_name])
        sheets.append({"name": sheet_name, "headers": TEMPLATE_HEADERS, "rows": rows})
    return sheets


def build_aggregate_rows(group_map: dict[str, list[list[str]]]) -> list[list[str]]:
    rows: list[list[str]] = []
    for group_name in sorted(group_map):
        group_rows = sorted(
            group_map[group_name],
            key=lambda row: (
                normalize_whitespace(row[4]) if len(row) > 4 else "",
                normalize_whitespace(row[5]) if len(row) > 5 else "",
                normalize_whitespace(row[1]) if len(row) > 1 else "",
            ),
        )
        rows.extend(group_rows)
    return renumber_template_rows(rows)


def load_output_manifests(collected_at: str) -> list[dict[str, object]]:
    records_dir = aggregate_root_dir(collected_at) / AGGREGATE_RECORDS_DIRNAME
    if not records_dir.exists():
        return []

    manifests = []
    for path in sorted(records_dir.glob("*.json")):
        try:
            with path.open("r", encoding="utf-8") as f:
                payload = json.load(f)
            if isinstance(payload, dict):
                manifests.append(payload)
        except Exception:
            continue
    return manifests


def aggregate_root_dir(collected_at: str) -> Path:
    return Path.cwd() / "输出结果" / collected_at


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", value)
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned.strip("._") or "output"


def sanitize_path_component(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", normalize_whitespace(value))
    cleaned = cleaned.strip(" .")
    return cleaned or ""


def clean_name(name: str) -> str:
    cleaned = normalize_whitespace(name)
    cleaned = re.sub(r"[（(][^（）()]{1,30}[）)]", "", cleaned)
    cleaned = re.sub(r"(老师简介|导师简介|个人简介|教师简介|简介)$", "", cleaned)
    title_suffix_pattern = r"(教授|副教授|讲师|博士|硕导|博导|老师|高级实验师|实验师|高级工程师|助理工程师|工程师|助理研究员|副研究员|研究员)$"
    title_inline_pattern = r"(?<!\S)(教授|副教授|讲师|博士|硕导|博导|老师|高级实验师|实验师|高级工程师|助理工程师|工程师|助理研究员|副研究员|研究员)(?!\S)"
    title_prefix_tokens = sorted(
        set(KNOWN_TITLES + ["博士", "硕导", "博导", "老师"]),
        key=len,
        reverse=True,
    )
    title_prefix_pattern = rf"^({'|'.join(re.escape(token) for token in title_prefix_tokens)})\s*[-—－]?\s*"
    cleaned = re.sub(title_prefix_pattern, "", cleaned)
    cleaned = re.sub(title_suffix_pattern, "", cleaned)
    cleaned = re.sub(title_inline_pattern, "", cleaned)
    cleaned = re.sub(r"项目$", "", cleaned)
    cleaned = re.sub(title_prefix_pattern, "", cleaned)
    cleaned = re.sub(title_suffix_pattern, "", cleaned)
    cleaned = re.sub(title_inline_pattern, "", cleaned)
    cleaned = re.sub(r"[ ]+", " ", cleaned).strip()
    return normalize_person_name(cleaned)


def infer_title_from_name(name: str) -> str:
    normalized = normalize_whitespace(name)
    if not normalized:
        return ""
    if "副教授" in normalized:
        return "副教授"
    if "讲师" in normalized:
        return "讲师"
    if "教授" in normalized:
        return "教授"
    return ""


def normalize_title(title: str) -> str:
    normalized = normalize_whitespace(title)
    specific_title = extract_specific_title_phrase(normalized)
    if specific_title:
        return collapse_special_title(specific_title)
    alias_title = find_title_in_text(normalized)
    if alias_title:
        return collapse_special_title(alias_title)
    priority_title = pick_priority_title(normalized)
    if priority_title:
        return priority_title
    for known in KNOWN_TITLES:
        if known in normalized:
            return collapse_special_title(known)
    return collapse_special_title(normalized)


def collapse_special_title(title: str) -> str:
    normalized = normalize_whitespace(title)
    if normalized == "特聘教授":
        return "教授"
    if normalized == "特聘副教授":
        return "副教授"
    return normalized


def get_field(row: List[str], index: int) -> str:
    if index < 0 or index >= len(row):
        return ""
    return normalize_whitespace(row[index])


def looks_like_url(value: str) -> bool:
    normalized = normalize_whitespace(value).lower()
    return normalized.startswith("http://") or normalized.startswith("https://")


def looks_like_name(value: str) -> bool:
    normalized = normalize_person_name(normalize_whitespace(value))
    if not normalized or looks_like_url(normalized):
        return False
    if "<" in normalized or ">" in normalized:
        return False
    return bool(re.fullmatch(r"[\u4e00-\u9fa5·]{2,10}", normalized))


def looks_like_school(value: str) -> bool:
    normalized = normalize_whitespace(value)
    if len(normalized) < 4:
        return False
    return any(normalized.endswith(suffix) for suffix in SCHOOL_SUFFIXES) and "系" not in normalized


def looks_like_college(value: str) -> bool:
    normalized = normalize_whitespace(value)
    if len(normalized) < 2:
        return False
    return any(normalized.endswith(suffix) for suffix in COLLEGE_SUFFIXES)


def is_placeholder_college(value: str) -> bool:
    normalized = normalize_whitespace(value).lower()
    if not normalized:
        return True
    normalized = normalized.replace(" ", "")
    if normalized in {"xx学院", "xx学部", "xx系", "xx中心", "xx部", "xx研究院"}:
        return True
    return bool(re.fullmatch(r"x{2,}(学院|学部|系|中心|部|研究院)", normalized))


def is_placeholder_school(value: str) -> bool:
    normalized = normalize_whitespace(value).lower()
    if not normalized:
        return True
    normalized = normalized.replace(" ", "")
    if normalized in {"xx大学", "xx学校", "xx研究院"}:
        return True
    return bool(re.fullmatch(r"x{2,}(大学|学校|研究院)", normalized))


def normalize_header(value: str) -> str:
    return normalize_whitespace(str(value)).lower()


def normalize_whitespace(value: str) -> str:
    text = decode_html(str(value or ""))
    text = text.replace("\ufeff", "").replace("ï»¿", "")
    text = text.replace("\u00a0", " ").replace("\u3000", " ").replace("\r", "\n").replace("\t", " ")
    text = re.sub(r"[ ]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def normalize_person_name(value: str) -> str:
    text = normalize_whitespace(value)
    return re.sub(r"(?<=[\u4e00-\u9fa5])\s+(?=[\u4e00-\u9fa5])", "", text)


def truncate_text(value: str, max_length: int) -> str:
    normalized = clean_intro_text(normalize_whitespace(value))
    return f"{normalized[:max_length]}..." if len(normalized) > max_length else normalized


def compact_csv_value(value: str) -> str:
    normalized = normalize_whitespace(value)
    return re.sub(r"\s*\n\s*", " ", normalized)


def build_teacher_context(text: str, name: str) -> str:
    normalized_text = text.strip()
    clean_teacher_name = clean_name(name)
    if not normalized_text or not clean_teacher_name:
        return ""

    index = normalized_text.find(clean_teacher_name)
    if index < 0:
        return ""

    start = max(0, index - 20)
    end = min(len(normalized_text), index + len(clean_teacher_name) + 400)
    return normalized_text[start:end]


def normalize_email_like_text(value: str) -> str:
    text = str(value or "")
    text = re.sub(r"\[at\]|\(at\)|（at）", "@", text, flags=re.IGNORECASE)
    text = re.sub(r"\[dot\]|\(dot\)|（dot）", ".", text, flags=re.IGNORECASE)
    text = text.replace("&#64;", "@").replace("＠", "@").replace("。", ".")
    text = re.sub(r"(?<=[A-Za-z0-9._%+-])\s*[#＃]\s*(?=[A-Za-z0-9.-]+\.(?:com|cn|edu|net|org|gov)\b)", "@", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*@\s*", "@", text)
    text = re.sub(r"\s*\.\s*", ".", text)
    text = re.sub(r"(?<=\w)\s+(?=[A-Za-z0-9._%+-]+@)", "", text)
    return re.sub(r"mailto:", "", text, flags=re.IGNORECASE)


def normalize_email_candidate(value: str) -> str:
    candidate = normalize_whitespace(value)
    candidate = candidate.strip(".,;:，。；：()[]{}<>")
    candidate = re.sub(r"^[\W_]+(?=[A-Za-z0-9])", "", candidate)
    return candidate


def extract_labeled_email(sources: List[str]) -> str:
    scored_candidates: List[tuple[int, str]] = []
    for source in sources:
        if not source:
            continue
        for pattern in EMAIL_LABEL_PATTERNS:
            for match in pattern.finditer(source):
                candidate_text = match.group(2)
                if "<" in candidate_text and ">" in candidate_text:
                    candidate_text = html_to_text(candidate_text)
                normalized = normalize_email_like_text(candidate_text)
                compacted = re.sub(r"\s+", "", normalized)
                scored_candidates.extend(score_email_candidates(normalized, "context"))
                if compacted != normalized:
                    scored_candidates.extend(score_email_candidates(compacted, "context"))
    return pick_best_scored_email(scored_candidates, min_score=70)


def is_reliable_email_candidate(email: str, sources: List[str]) -> bool:
    normalized_email = normalize_email_candidate(email).lower()
    best_score = -10**9
    for source_name, source in zip(("context", "text", "html"), sources):
        if not source:
            continue
        normalized_source = normalize_email_like_text(source)
        for match in EMAIL_PATTERN.finditer(normalized_source):
            candidate = normalize_email_candidate(match.group(0)).lower()
            if candidate != normalized_email:
                continue
            source_bonus = {"context": 80, "text": 60, "html": 20}.get(source_name, 0)
            score = source_bonus + score_email_match(normalized_source, match.start(), match.end(), match.group(0))
            best_score = max(best_score, score)
    return best_score >= 70


def extract_labeled_title(source: str) -> str:
    for pattern in TITLE_LABEL_CAPTURE_PATTERNS:
        match = pattern.search(source)
        if not match:
            continue
        value = clean_labeled_field_value(match.group(1))
        normalized = map_title_by_level(value)
        if normalized:
            return normalized
        alias_title = find_title_in_text(value)
        if alias_title:
            return alias_title
        priority_title = pick_priority_title(value)
        if priority_title:
            return priority_title
        for title in KNOWN_TITLES:
            if title in value:
                return title
        cleaned = clean_labeled_title_value(value)
        if cleaned:
            return cleaned
    return ""


def extract_email_from_images(html_content: str, image_ocr_text: str = "") -> str:
    recognized_text = image_ocr_text or extract_image_text(html_content)
    if recognized_text:
        normalized = normalize_email_like_text(recognized_text)
        preferred = pick_best_scored_email(score_email_candidates(normalized, "ocr"))
        if preferred:
            return preferred
    return ""


def extract_image_text(html_content: str) -> str:
    image_urls = extract_image_urls(html_content)
    recognized_blocks = []
    for image_url in image_urls[:3]:
        recognized_text = recognize_image_text(image_url)
        if recognized_text:
            recognized_blocks.append(recognized_text)
    return "\n".join(recognized_blocks)


def should_run_ocr(html_content: str, text: str) -> bool:
    image_urls = extract_image_urls(html_content)
    if not image_urls:
        return False

    vsb_content = extract_vsb_content(html_content)
    if vsb_content:
        vsb_text = normalize_whitespace(html_to_text(vsb_content))
        if "img_vsb_content" in vsb_content and len(vsb_text) < 40:
            return True

    body_text = normalize_whitespace(text)
    strong_info_tokens = ("邮箱", "电子邮箱", "电子信箱", "@", "职称", "教授", "副教授", "讲师", "研究员", "工程师")
    if any(token in body_text for token in strong_info_tokens):
        return False

    if "img_vsb_content" in html_content:
        return True

    return len(body_text) < 200


def extract_vsb_content(html_content: str) -> str:
    match = VSB_CONTENT_PATTERN.search(html_content)
    return match.group(0) if match else ""


def clean_intro_text(value: str) -> str:
    cleaned = normalize_whitespace(value)
    cleaned = re.sub(r"^(View\s+English\s+CV|View\s+Chinese\s+CV)\s*", "", cleaned, flags=re.IGNORECASE)
    return cleaned


def is_valid_intro_text(value: str) -> bool:
    cleaned = clean_intro_text(value)
    if not cleaned or len(cleaned) < 8:
        return False
    bad_prefixes = (
        "西安电子科技大学教师个人主页系统采用PHP和MySQL开源技术开发",
        "信息来源：不",
        "信息来源:",
        "信息来源：",
        "来源：不",
        "来源:",
        "来源：",
        "暂无",
        "无",
        "空",
    )
    if any(cleaned.startswith(prefix) for prefix in bad_prefixes):
        return False
    contact_tokens = ("学院地址", "邮编", "传真", "联系电话", "电话:")
    if sum(token in cleaned for token in contact_tokens) >= 2:
        return False
    return True


def extract_image_urls(html_content: str) -> List[str]:
    urls = []
    for match in IMAGE_URL_PATTERN.finditer(html_content):
        raw_url = normalize_whitespace(match.group(1))
        if not raw_url:
            continue
        if raw_url.startswith("//"):
            raw_url = f"https:{raw_url}"
        elif raw_url.startswith("/"):
            raw_url = f"https://rwxy.xaut.edu.cn{raw_url}"
        elif raw_url.startswith("../") or raw_url.startswith("./"):
            raw_url = f"https://rwxy.xaut.edu.cn/info/1180/{raw_url}"
        if raw_url not in urls:
            urls.append(raw_url)
    return urls


def recognize_image_text(image_url: str) -> str:
    if not VISION_OCR_SCRIPT.exists():
        return ""

    suffix = Path(image_url).suffix or ".img"
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
            temp_path = Path(temp_file.name)
        request = Request(
            image_url,
            headers={"User-Agent": "Mozilla/5.0 Codex Teacher Pipeline/1.0"},
        )
        with urlopen(request, timeout=15) as response:
            temp_path.write_bytes(response.read())

        result = subprocess.run(
            ["swift", str(VISION_OCR_SCRIPT), str(temp_path)],
            capture_output=True,
            text=True,
            timeout=30,
        )
        if result.returncode != 0:
            return ""
        return normalize_whitespace(result.stdout)
    except Exception:
        return ""
    finally:
        try:
            temp_path.unlink(missing_ok=True)  # type: ignore[name-defined]
        except Exception:
            pass


def map_title_by_level(value: str) -> str:
    if PROFESSOR_LEVEL_PATTERN.search(value):
        return "教授"
    if ASSOCIATE_PROFESSOR_LEVEL_PATTERN.search(value):
        return "副教授"
    return ""


def clean_labeled_title_value(value: str) -> str:
    cleaned = clean_labeled_field_value(value)
    specific_title = extract_specific_title_phrase(cleaned)
    if specific_title:
        return specific_title
    cleaned = re.split(r"[，,；;。/ ]", cleaned)[0]
    return cleaned[:20] if cleaned else ""


def is_valid_title_value(value: str) -> bool:
    cleaned = normalize_whitespace(value)
    if not cleaned:
        return False
    if not re.search(r"[\u4e00-\u9fa5A-Za-z]", cleaned):
        return False
    invalid_values = {
        "教",
        "teacher",
        "教师",
        "职务",
        "职务：",
        "职称",
        "职称：",
        "职称/职务",
        "职称/职务：",
    }
    return cleaned not in invalid_values


def clean_labeled_field_value(value: str) -> str:
    cleaned = normalize_whitespace(value)
    cleaned = re.sub(r"^(为|是)", "", cleaned)
    cleaned = re.sub(
        r"(硕导\s*/\s*博导|硕导|博导|毕业院校|联系方式|联系电话|办公电话|电话|办公室|邮箱|电子邮箱|电子信箱|Email|E-mail|Emial|个人简历).*$",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )
    return normalize_whitespace(cleaned)


def pick_priority_title(value: str) -> str:
    for title in PRIORITY_TITLES:
        if title in value or title_matches_alias(value, title):
            return title
    return ""


def extract_specific_title_phrase(value: str) -> str:
    cleaned = normalize_whitespace(value)
    patterns = [
        r"(科研副教授|科研教授|特聘副教授|特聘教授)",
    ]
    for pattern in patterns:
        match = re.search(pattern, cleaned)
        if match:
            return match.group(1)
    return ""


def find_title_in_text(source: str) -> str:
    for _, _, title in find_title_matches(source):
        if not is_valid_title_value(title):
            continue
        if title == "教师":
            continue
        return title
    return ""


def find_title_matches(source: str) -> List[tuple[int, int, str]]:
    matches: List[tuple[int, int, str]] = []
    for canonical, pattern in TITLE_PATTERN_ENTRIES:
        for match in pattern.finditer(source):
            matches.append((match.start(), match.end(), canonical))
    matches.sort(key=lambda item: (item[0], title_priority(item[2]), -(item[1] - item[0])))

    deduped: List[tuple[int, int, str]] = []
    seen = set()
    for item in matches:
        key = (item[0], item[1], item[2])
        if key in seen:
            continue
        seen.add(key)
        deduped.append(item)
    return deduped


def iter_title_alias_matches(source: str, alias: str):
    if re.search(r"[A-Za-z]", alias):
        return re.finditer(rf"\b{re.escape(alias)}\b", source, re.IGNORECASE)
    return re.finditer(re.escape(alias), source)


def title_matches_alias(value: str, canonical: str) -> bool:
    aliases = TITLE_ALIASES.get(canonical, [canonical])
    return any(next(iter_title_alias_matches(value, alias), None) for alias in aliases)


def title_priority(title: str) -> int:
    try:
        return KNOWN_TITLES.index(title)
    except ValueError:
        return len(KNOWN_TITLES)


def collect_emails(value: str) -> List[str]:
    unique = []
    seen = set()
    for email in EMAIL_PATTERN.findall(value):
        candidate = normalize_email_candidate(email)
        if not candidate:
            continue
        key = candidate.lower()
        if key not in seen:
            seen.add(key)
            unique.append(candidate)
    return unique


def pick_best_email(candidates: List[str]) -> str:
    if not candidates:
        return ""

    bad_domains = {"example.com", "test.com", "163.com.cn"}
    filtered = [
        email
        for email in candidates
        if not any(email.lower().endswith(f"@{domain}") for domain in bad_domains)
    ]
    return filtered[0] if filtered else candidates[0]


def score_email_candidates(source: str, source_name: str) -> List[tuple[int, str]]:
    scored_candidates: List[tuple[int, str]] = []
    seen = set()
    source_bonus = {"context": 80, "text": 60, "html": 20, "ocr": 10}.get(source_name, 0)
    for match in EMAIL_PATTERN.finditer(source):
        candidate = normalize_email_candidate(match.group(0))
        if not candidate:
            continue
        key = candidate.lower()
        if key in seen:
            continue
        seen.add(key)
        score = source_bonus + score_email_match(source, match.start(), match.end(), candidate)
        scored_candidates.append((score, candidate))
    return scored_candidates


def pick_best_scored_email(candidates: List[tuple[int, str]], min_score: int = 1) -> str:
    if not candidates:
        return ""

    deduped: Dict[str, int] = {}
    for score, email in candidates:
        key = email.lower()
        deduped[key] = max(score, deduped.get(key, -10**9))

    best_score, best_email = max(
        ((score, email) for email, score in deduped.items()),
        key=lambda item: (item[0], item[1]),
    )
    return best_email if best_score >= min_score else ""


def score_email_match(source: str, start: int, end: int, email: str) -> int:
    window_start = max(0, start - 200)
    window_end = min(len(source), end + 200)
    window = source[window_start:window_end]
    score = 10
    lowered = email.lower()
    relative_position = start / max(len(source), 1)
    local_part = lowered.split("@", 1)[0]
    has_email_label = any(token in window for token in ("电子邮件", "电子邮箱", "邮箱", "Email", "E-mail", "联系邮箱"))
    has_profile_contact_section = any(
        token in window
        for token in ("其他联系方式", "Contact information", "个人信息", "基本信息", "联系方式", "联系电话", "电子信箱")
    )
    has_footer_tokens = any(token in window for token in ("上一篇", "下一篇", "版权", "版权所有", "技术支持", "官方微信", "友情链接", "Copyright"))
    looks_like_footer_contact_block = (
        has_footer_tokens
        and any(token in window for token in ("地址：", "地址:", "联系电话", "电话：", "电话:"))
    )
    looks_like_public_mailbox = local_part in {
        "admin",
        "admission",
        "college",
        "contact",
        "dean",
        "department",
        "faculty",
        "info",
        "office",
        "service",
        "support",
        "webmaster",
        "xueyuan",
    } or any(token in local_part for token in ("noreply", "no-reply", "postmaster", "webmaster", "support", "office", "contact", "admin"))
    looks_like_personal_email = not looks_like_public_mailbox

    if has_email_label:
        score += 30
    if has_email_label and looks_like_personal_email:
        score += 50
    if any(token in window for token in ("办公地点", "研究方向", "职称", "职称/岗位", "毕业学校", "基本信息", "个人信息", "工作单位", "导师")):
        score += 15
    if any(token in window for token in ("投诉邮箱", "监督邮箱", "师德师风")):
        score -= 80
    if has_profile_contact_section:
        score += 45
    if looks_like_footer_contact_block:
        score -= 180
    if has_footer_tokens and not has_profile_contact_section and not (has_email_label and looks_like_personal_email):
        score -= 50
    if "地址：" in window and "电话" in window:
        score -= 40
    if relative_position <= 0.45:
        score += 15
    elif relative_position >= 0.82 and not has_profile_contact_section and not (has_email_label and looks_like_personal_email):
        score -= 80
    elif relative_position >= 0.70 and not has_profile_contact_section and not (has_email_label and looks_like_personal_email):
        score -= 35
    if looks_like_public_mailbox:
        score -= 50
    return score


def html_to_text(html_content: str) -> str:
    text = re.sub(r"<script[\s\S]*?</script>", " ", html_content, flags=re.IGNORECASE)
    text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</p>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    text = decode_html(text)
    text = re.sub(r"[ ]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def decode_html(text: str) -> str:
    return html.unescape(str(text or ""))


def today() -> str:
    return dt.date.today().isoformat()


if __name__ == "__main__":
    main()
