#!/usr/bin/env python3

from datetime import date
from pathlib import Path
import subprocess
import sys
import re
from urllib.parse import urlsplit


BASE_DIR = Path("/Users/xinxinhuashe/Documents/易达威实习")
SCRIPT_PATH = BASE_DIR / "scripts" / "teacher_profile_pipeline.py"
if str(SCRIPT_PATH.parent) not in sys.path:
    sys.path.insert(0, str(SCRIPT_PATH.parent))

try:
    from teacher_profile_pipeline import infer_org_fields_from_page, infer_school_name_from_url
except Exception:
    infer_org_fields_from_page = None
    infer_school_name_from_url = None

OUTPUT_DIR = BASE_DIR / "输出结果"
TEST_OUTPUT_DIR = BASE_DIR / "测试结果"
RAW_INPUT_DIR = BASE_DIR / "原始输入"
DEFAULT_INPUT_DATE = "2026-06-04"
DEFAULT_INPUT_PATH = RAW_INPUT_DIR / DEFAULT_INPUT_DATE / "xx大学-xx学院-教师信息.csv"
DEFAULT_COLLECTOR = "朱超逸"
DEFAULT_SCHOOL = "xx大学"
DEFAULT_COLLEGE = "xx学院"
DEFAULT_DATE = date.today().isoformat()
DEFAULT_CONCURRENCY = "5"
SCHOOL_SUFFIXES = ("大学", "学校", "研究院")
COLLEGE_SUFFIXES = ("学院", "系", "中心", "部", "所", "研究院")


def main() -> None:
    ensure_script_exists()

    input_source, is_url, is_profile_url = prompt_input_source()
    collected_at = prompt_text("请输入采集日期", DEFAULT_DATE)
    recent_school = infer_recent_school(collected_at)
    recent_college = infer_recent_college(collected_at, recent_school)
    default_school, default_college = infer_defaults(input_source, is_url, recent_school, recent_college)

    school = prompt_text("请输入学校", default_school)
    college = prompt_text("请输入学院", default_college)
    output_path = prompt_path(
        prompt_text="请输入输出 CSV 路径",
        default_path=build_default_output_path(input_source, is_url, collected_at, school, college),
        must_exist=False,
        date_folder=collected_at,
    )
    warn_if_output_path_mismatches(output_path, school, college)
    collector = prompt_text("请输入采集人", DEFAULT_COLLECTOR)
    concurrency = prompt_text("请输入并发数", DEFAULT_CONCURRENCY)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    command = [
        sys.executable,
        str(SCRIPT_PATH),
        "--output",
        str(output_path),
        "--school",
        school,
        "--college",
        college,
        "--collector",
        collector,
        "--date",
        collected_at,
    ]
    if is_url and is_profile_url:
        command.extend(["--profile-url", str(input_source)])
    elif is_url:
        command.extend(["--list-page-url", str(input_source)])
    else:
        command.extend(["--input", str(input_source)])
        command.extend(["--concurrency", concurrency])

    print("\n即将执行：")
    print(" ".join(command))
    print("")

    subprocess.run(command, check=True, cwd=BASE_DIR)


def ensure_script_exists() -> None:
    if not SCRIPT_PATH.exists():
        raise FileNotFoundError(f"找不到采集脚本：{SCRIPT_PATH}")


def prompt_path(prompt_text: str, default_path: Path, must_exist: bool, date_folder: str | None = None) -> Path:
    while True:
        raw = input(f"{prompt_text}（回车使用默认值）\n默认值: {default_path}\n> ").strip()
        candidate = (
            normalize_output_path(raw, default_path, date_folder or DEFAULT_DATE)
            if not must_exist
            else resolve_input_path(raw, default_path)
        )

        if must_exist and not candidate.exists():
            print(f"文件不存在：{candidate}\n")
            continue

        return candidate


def prompt_input_source() -> tuple[Path | str, bool, bool]:
    print("请输入输入 CSV 路径或文件名关键词。")
    print("也可以直接输入单页列表页 URL 或单个教师主页 URL。")
    print("直接回车：使用默认文件")
    print("输入完整路径或相对路径：直接使用该文件")
    print("输入关键词：自动搜索并让你选择")

    while True:
        raw = input(f"默认值: {DEFAULT_INPUT_PATH}\n> ").strip()
        if not raw:
            return DEFAULT_INPUT_PATH, False, False

        if looks_like_url(raw):
            return raw, True, looks_like_teacher_profile_url(raw)

        direct_candidate = Path(raw).expanduser()
        if direct_candidate.is_absolute() or "/" in raw or raw.endswith(".csv"):
            resolved = resolve_input_path(raw, DEFAULT_INPUT_PATH)
            if resolved.exists():
                return resolved, False, False
            print(f"文件不存在：{resolved}\n")
            continue

        matches = search_csv_files(raw)
        if not matches:
            print(f"没有找到包含“{raw}”的 CSV 文件。\n")
            continue
        if len(matches) == 1:
            print(f"已自动匹配：{matches[0]}\n")
            return matches[0], False, False
        return choose_from_matches(raw, matches), False, False


def resolve_input_path(raw: str, default_path: Path) -> Path:
    candidate = Path(raw).expanduser() if raw else default_path
    if candidate.is_absolute():
        return candidate.resolve()
    return (BASE_DIR / candidate).resolve()


def normalize_output_path(raw: str, default_path: Path, date_folder: str) -> Path:
    if not raw:
        return default_path

    candidate = Path(raw).expanduser()
    if candidate.is_absolute():
        return ensure_csv_suffix(candidate).resolve()

    if "/" not in raw and "\\" not in raw:
        return ensure_csv_suffix(select_output_base_dir(candidate.name, date_folder) / candidate.name).resolve()

    return ensure_csv_suffix(BASE_DIR / candidate).resolve()


def search_csv_files(keyword: str) -> list[Path]:
    keyword_lower = keyword.lower()
    matches = []
    for path in BASE_DIR.rglob("*.csv"):
        if "输出结果" in path.parts or "6.1输出结果" in path.parts or "日志" in path.parts:
            continue
        relative = path.relative_to(BASE_DIR).as_posix().lower()
        if keyword_lower in relative or keyword_lower in path.name.lower():
            matches.append(path.resolve())
    return sorted(matches)


def choose_from_matches(keyword: str, matches: list[Path]) -> Path:
    print(f"找到 {len(matches)} 个包含“{keyword}”的 CSV 文件，请输入编号：")
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


def prompt_text(label: str, default_value: str) -> str:
    raw = input(f"{label}（回车使用默认值：{default_value}）\n> ").strip()
    return raw or default_value


def build_default_output_path(
    input_source: Path | str,
    is_url: bool,
    collected_at: str,
    school: str,
    college: str = "",
) -> Path:
    school_dir = school or DEFAULT_SCHOOL
    college_name = college or (
        build_url_output_stem(str(input_source)) if is_url else build_output_stem(Path(input_source).stem)
    )
    filename = normalize_output_filename(college_name or (build_url_output_stem(str(input_source)) if is_url else build_output_stem(Path(input_source).stem)))
    return build_output_dir(collected_at, school_dir) / f"{filename}.csv"


def warn_if_output_path_mismatches(output_path: Path, school: str, college: str) -> None:
    normalized_school = normalize_org_name(school)
    normalized_college = normalize_org_name(college)
    normalized_parts = {normalize_org_name(part) for part in output_path.parts}
    normalized_stem = normalize_org_name(output_path.stem)
    if normalized_school and normalized_school not in normalized_parts:
        print(f"提示：当前输出路径不在“{school}”目录下：{output_path}")
    if normalized_college and normalized_college not in normalized_parts and normalized_college != normalized_stem:
        print(f"提示：当前输出文件名与学院“{college}”不一致：{output_path}")


def build_output_dir(collected_at: str, school: str = "") -> Path:
    base_dir = OUTPUT_DIR / collected_at
    return base_dir / school if school and school != DEFAULT_SCHOOL else base_dir


def build_test_output_dir(collected_at: str, school: str = "") -> Path:
    base_dir = TEST_OUTPUT_DIR / collected_at
    return base_dir / school if school and school != DEFAULT_SCHOOL else base_dir


def select_output_base_dir(filename: str, collected_at: str) -> Path:
    normalized = filename.lower()
    school = extract_school_name(clean_input_stem(Path(filename).stem)) or infer_recent_school(collected_at)
    if any(token in normalized for token in ["测试", "test", "修正", "规则"]):
        return build_test_output_dir(collected_at, school)
    return build_output_dir(collected_at, school)


def ensure_csv_suffix(path: Path) -> Path:
    return path if path.suffix.lower() == ".csv" else path.with_suffix(".csv")


def normalize_output_filename(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", value.strip())
    return cleaned.strip(" .") or "未命名学院"


def build_output_stem(input_stem: str) -> str:
    stem = input_stem.strip()
    for suffix in ["_导入_Python", "_导入", "_导出", "导入_Python", "导入", "导出", "输出"]:
        if stem.endswith(suffix):
            stem = stem[: -len(suffix)].rstrip("_- ")
            break
    return f"{stem}输出"


def build_url_output_stem(url: str) -> str:
    split = urlsplit(url)
    stem = Path(split.path).stem or "列表页"
    return build_output_stem(stem)


def infer_defaults_from_input_path(input_path: Path, recent_school: str = "", recent_college: str = "") -> tuple[str, str]:
    stem = clean_input_stem(input_path.stem)
    school = extract_school_name(stem) or recent_school or DEFAULT_SCHOOL
    college = extract_college_name(stem, school) or recent_college or DEFAULT_COLLEGE
    return school, college


def infer_defaults(
    input_source: Path | str,
    is_url: bool,
    recent_school: str = "",
    recent_college: str = "",
) -> tuple[str, str]:
    if is_url:
        page_school = ""
        page_college = ""
        if infer_org_fields_from_page is not None:
            try:
                page_school, page_college = infer_org_fields_from_page(str(input_source))
            except Exception:
                page_school, page_college = "", ""
        school = (
            page_school
            or (infer_school_name_from_url(str(input_source)) if infer_school_name_from_url is not None else "")
            or extract_school_name(str(input_source))
            or recent_school
            or DEFAULT_SCHOOL
        )
        college = (
            page_college
            or extract_college_name(clean_input_stem(Path(urlsplit(str(input_source)).path).stem), school)
            or recent_college
            or DEFAULT_COLLEGE
        )
        return school, college
    return infer_defaults_from_input_path(Path(input_source), recent_school, recent_college)


def infer_recent_school(collected_at: str) -> str:
    dated_output_dir = OUTPUT_DIR / collected_at
    candidates = []
    if dated_output_dir.exists():
        for path in dated_output_dir.iterdir():
            if not path.is_dir():
                continue
            if not looks_like_school(path.name):
                continue
            try:
                modified_at = path.stat().st_mtime
            except OSError:
                continue
            candidates.append((modified_at, path.name))

    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][1]

    all_candidates = []
    for path in OUTPUT_DIR.rglob("*"):
        if not path.is_dir():
            continue
        if not looks_like_school(path.name):
            continue
        try:
            modified_at = path.stat().st_mtime
        except OSError:
            continue
        all_candidates.append((modified_at, path.name))

    if all_candidates:
        all_candidates.sort(reverse=True)
        return all_candidates[0][1]

    return ""


def infer_recent_college(collected_at: str, school: str = "") -> str:
    candidates = []
    ignored_colleges = {DEFAULT_COLLEGE, "未命名学院"}

    def collect_from_dir(base_dir: Path) -> None:
        if not base_dir.exists():
            return
        for path in base_dir.rglob("*.csv"):
            if school and school not in path.parts:
                continue
            college = extract_college_name(clean_input_stem(path.stem), school)
            if not college or college in ignored_colleges:
                continue
            try:
                modified_at = path.stat().st_mtime
            except OSError:
                continue
            candidates.append((modified_at, college))

    collect_from_dir(OUTPUT_DIR / collected_at)
    collect_from_dir(TEST_OUTPUT_DIR / collected_at)

    if not candidates:
        collect_from_dir(OUTPUT_DIR)
        collect_from_dir(TEST_OUTPUT_DIR)

    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][1]

    return ""


def clean_input_stem(stem: str) -> str:
    cleaned = stem
    for token in ["官网", "网站", "欢迎您", "欢迎您！", "教师信息", "教师列表", "老师列表", "导入", "导出", "输出"]:
        cleaned = cleaned.replace(token, "")
    cleaned = re.sub(r"\(\d+\)$", "", cleaned)
    return re.sub(r"[-_]+", "-", cleaned).strip("- ")


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


def looks_like_school(value: str) -> bool:
    if len(value) < 4:
        return False
    return any(value.endswith(suffix) for suffix in SCHOOL_SUFFIXES) and "系" not in value


def looks_like_college(value: str) -> bool:
    if len(value) < 2:
        return False
    return any(value.endswith(suffix) for suffix in COLLEGE_SUFFIXES)


def looks_like_url(value: str) -> bool:
    normalized = value.strip().lower()
    return normalized.startswith("http://") or normalized.startswith("https://")


def looks_like_teacher_profile_url(value: str) -> bool:
    normalized = value.strip().lower()
    path = urlsplit(normalized).path
    if any(token in normalized for token in ["faculty.", "教师个人主页"]):
        return True
    if re.search(r"/info/\d+/\d+\.htm$", path):
        return True
    return path.endswith("/index.htm")


if __name__ == "__main__":
    main()
