#!/usr/bin/env python3

from __future__ import annotations

import csv
import shutil
from pathlib import Path
from xml.sax.saxutils import escape
import zipfile
import xml.etree.ElementTree as ET


INVALID_SHEET_CHARS = set('[]:*?/\\')


def write_workbook(path: Path, sheets: list[dict[str, object]]) -> None:
    prepared_sheets = prepare_sheets(sheets)
    path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", build_content_types_xml(len(prepared_sheets)))
        zf.writestr("_rels/.rels", ROOT_RELS_XML)
        zf.writestr("xl/workbook.xml", build_workbook_xml(prepared_sheets))
        zf.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels_xml(len(prepared_sheets)))
        zf.writestr("xl/styles.xml", STYLES_XML)
        for index, sheet in enumerate(prepared_sheets, start=1):
            zf.writestr(
                f"xl/worksheets/sheet{index}.xml",
                build_sheet_xml(sheet["headers"], sheet["rows"]),
            )


def write_csv(path: Path, headers: list[object], rows: list[list[object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([normalize_cell(value) for value in headers])
        writer.writerows(
            [[normalize_cell(cell) for cell in row] for row in rows]
        )


def write_sheet_csv_preview(preview_dir: Path, sheets: list[dict[str, object]]) -> None:
    if preview_dir.exists():
        shutil.rmtree(preview_dir)
    preview_dir.mkdir(parents=True, exist_ok=True)

    prepared_sheets = prepare_sheets(sheets)
    for index, sheet in enumerate(prepared_sheets, start=1):
        file_name = f"{index:02d}_{sheet['name']}.csv"
        write_csv(preview_dir / file_name, sheet["headers"], sheet["rows"])


def read_first_sheet(path: Path) -> tuple[list[str], list[list[str]]]:
    with zipfile.ZipFile(path) as zf:
        sheet_xml = zf.read("xl/worksheets/sheet1.xml")

    namespace = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    root = ET.fromstring(sheet_xml)
    rows: list[list[str]] = []
    for row_element in root.findall(".//a:sheetData/a:row", namespace):
        cells: dict[int, str] = {}
        max_index = 0
        for cell in row_element.findall("a:c", namespace):
            reference = cell.attrib.get("r", "")
            column_index = column_index_from_reference(reference)
            max_index = max(max_index, column_index)
            text_node = cell.find("a:is/a:t", namespace)
            value = text_node.text if text_node is not None and text_node.text is not None else ""
            cells[column_index] = value
        row_values = [cells.get(index, "") for index in range(1, max_index + 1)]
        rows.append(row_values)

    if not rows:
        return [], []
    return rows[0], rows[1:]


def prepare_sheets(sheets: list[dict[str, object]]) -> list[dict[str, object]]:
    used_names: set[str] = set()
    prepared = []
    for index, sheet in enumerate(sheets, start=1):
        raw_name = str(sheet.get("name") or f"Sheet{index}")
        name = uniquify_sheet_name(raw_name, used_names)
        headers = [normalize_cell(value) for value in list(sheet.get("headers") or [])]
        rows = [
            [normalize_cell(cell) for cell in list(row)]
            for row in list(sheet.get("rows") or [])
        ]
        prepared.append({"name": name, "headers": headers, "rows": rows})
        used_names.add(name)
    return prepared or [{"name": "Sheet1", "headers": [], "rows": []}]


def uniquify_sheet_name(name: str, used_names: set[str]) -> str:
    cleaned = "".join("_" if char in INVALID_SHEET_CHARS else char for char in name).strip()
    cleaned = cleaned.strip("'") or "Sheet"
    cleaned = cleaned[:31]
    if cleaned not in used_names:
        return cleaned

    base = cleaned[:28] or "Sheet"
    counter = 2
    while True:
        candidate = f"{base}_{counter}"[:31]
        if candidate not in used_names:
            return candidate
        counter += 1


def build_content_types_xml(sheet_count: int) -> str:
    overrides = [
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for index in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    joined = "".join(overrides)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f"{joined}"
        "</Types>"
    )


def build_workbook_xml(sheets: list[dict[str, object]]) -> str:
    parts = []
    for index, sheet in enumerate(sheets, start=1):
        parts.append(
            f'<sheet name="{escape(str(sheet["name"]))}" sheetId="{index}" '
            f'r:id="rId{index}"/>'
        )
    joined = "".join(parts)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{joined}</sheets>"
        "</workbook>"
    )


def build_workbook_rels_xml(sheet_count: int) -> str:
    parts = []
    for index in range(1, sheet_count + 1):
        parts.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    parts.append(
        f'<Relationship Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    joined = "".join(parts)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{joined}"
        "</Relationships>"
    )


def build_sheet_xml(headers: list[str], rows: list[list[str]]) -> str:
    data_rows = [headers] if headers else []
    data_rows.extend(rows)
    max_columns = max((len(row) for row in data_rows), default=1)
    max_rows = max(len(data_rows), 1)
    dimension = f"A1:{column_name(max_columns)}{max_rows}"

    xml_rows = []
    for row_index, row in enumerate(data_rows, start=1):
        cells = []
        for column_index, value in enumerate(row, start=1):
            if value == "":
                continue
            reference = f"{column_name(column_index)}{row_index}"
            escaped_value = escape(value)
            cells.append(
                f'<c r="{reference}" t="inlineStr"><is><t xml:space="preserve">{escaped_value}</t></is></c>'
            )
        xml_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    sheet_data = "".join(xml_rows)

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="{dimension}"/>'
        "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>"
        "<sheetFormatPr defaultRowHeight=\"15\"/>"
        f"<sheetData>{sheet_data}</sheetData>"
        "</worksheet>"
    )


def column_name(index: int) -> str:
    result = []
    current = index
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result)) or "A"


def column_index_from_reference(reference: str) -> int:
    letters = "".join(char for char in reference if char.isalpha()).upper()
    result = 0
    for char in letters:
        result = result * 26 + (ord(char) - 64)
    return result or 1


def normalize_cell(value: object) -> str:
    if value is None:
        return ""
    return str(value)


ROOT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""


STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
  </fonts>
  <fills count="1">
    <fill>
      <patternFill patternType="none"/>
    </fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
"""
