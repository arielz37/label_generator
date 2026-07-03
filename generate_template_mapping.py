from __future__ import annotations

import argparse
import csv
import re
import unicodedata
import zipfile
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
from xml.sax.saxutils import escape, unescape

try:
    import app_paths
    from app_paths import TEMPLATE_MAPPING_FILE, TEMPLATE_ROOT
except Exception:
    app_paths = None
    TEMPLATE_ROOT = Path(__file__).resolve().parent / "Templates"
    TEMPLATE_MAPPING_FILE = Path(__file__).resolve().parent / "template_mapping.xlsx"

DEFAULT_TEMPLATE_ROOT = TEMPLATE_ROOT
DEFAULT_OUTPUT = TEMPLATE_MAPPING_FILE

INACTIVE_KEYWORDS = ("旧版本", "勿用", "老版本", "暂停", "待确定", "不用", "弃用", "备用版本", "旧印刷")
INNER_KEYWORDS = ("内标", "内 标")
OUTER_KEYWORDS = ("外标", "外 标", "外箱", "外 箱")
INNER_OUTER_KEYWORDS = ("内外标", "内外 标", "外内标", "外内 标", "内外", "外内")
INNER_PACKAGE_HINTS = ("内罐", "内袋", "一罐", "一袋")
OUTER_PACKAGE_HINTS = ("外箱", "一箱", "-箱", "箱）", "箱)", "箱")
DELIVERY_KEYWORDS = ("送货单", "发货清单")
PALLET_KEYWORDS = ("托盘",)
INNER_SYSTEM_LABEL_PHRASES = ("内标是系统标", "内标是系统表", "内标签是系统标", "内标签是系统表")

NOISE_WORDS = (
    "内标用二维码的版本",
    "内标用二维码版本",
    "外标内标用二维码的版本",
    "外标内标用二维码版本",
    "发货清单二维码",
    "二维码版本",
    "二维码",
    "版本",
    "新版本",
    "内外标",
    "外内标",
    "内外",
    "外内",
    "内标",
    "外标",
    "外箱",
    "外盒",
    "送货单",
    "托盘标签",
    "托盘",
    "标签",
    "管装",
    "铝箔袋",
    "屏蔽袋",
    "正方形",
    "长方形",
    "方形",
    "天地盖",
    "样品",
    "OK的模板",
    "OK",
    "保质期一年半",
    "保质期一年",
    "保质期袋子2年",
)

HEADERS = (
    "客户编号",
    "匹配编号",
    "标签类型",
    "客户目录",
    "模板相对路径",
    "问题",
)


@dataclass
class MappingRow:
    customer_code: str
    manual_match_code: str
    manual_label_type: str
    customer_dir: str
    file_name: str
    relative_path: str
    issue: str
    status: str = ""
    enabled_suggestion: str = ""
    match_code: str = ""
    label_type: str = ""

    def as_list(self) -> List[str]:
        return [
            self.customer_code,
            self.manual_match_code,
            self.manual_label_type,
            self.customer_dir,
            self.relative_path,
            self.display_issue(),
        ]

    def display_issue(self) -> str:
        parts = [item for item in self.issue.split("；") if item]
        if self.status:
            parts.append(f"状态：{self.status}")
        if self.enabled_suggestion:
            parts.append(f"启用建议：{self.enabled_suggestion}")
        return "；".join(dict.fromkeys(parts))


def normalize_text(value: str) -> str:
    return unicodedata.normalize("NFKC", value or "")


def contains_any(value: str, keywords: Iterable[str]) -> bool:
    return any(keyword in value for keyword in keywords)


def detect_label_type(stem: str, relative_path: Path | None = None) -> Tuple[str, str]:
    text = normalize_text(stem)
    path_context = ""
    if relative_path is not None and len(relative_path.parts) > 2:
        # Use subfolders such as "内标 59-xxxx" but ignore the customer root
        # directory, which often contains operational notes rather than a label type.
        path_context = " ".join(relative_path.parts[1:-1])

    context = normalize_text(f"{path_context} {stem}")
    compact = re.sub(r"\s+", "", context)
    stem_compact = re.sub(r"\s+", "", text)
    note = ""

    inner_system_label = contains_any(compact, INNER_SYSTEM_LABEL_PHRASES)
    for phrase in INNER_SYSTEM_LABEL_PHRASES:
        compact = compact.replace(phrase, "")
        stem_compact = stem_compact.replace(phrase, "")

    has_inner = contains_any(compact, INNER_KEYWORDS)
    has_outer = contains_any(compact, OUTER_KEYWORDS)
    if contains_any(compact, INNER_OUTER_KEYWORDS):
        return "内外标", note
    if contains_any(text, DELIVERY_KEYWORDS):
        return "送货单", note
    if contains_any(text, PALLET_KEYWORDS):
        return "托盘", note
    if has_inner and has_outer:
        return "内外标", note
    if has_inner:
        return "内标", note
    if has_outer:
        return "外标", note

    # Many newly-added customer templates are named only by part number plus
    # package quantity, for example "TAPLOS230KX330X0X15（500一箱）".
    # Treat box-like names as outer labels and bag/can-like names as inner labels.
    if contains_any(stem_compact, OUTER_PACKAGE_HINTS):
        return "外标", note
    if contains_any(stem_compact, INNER_PACKAGE_HINTS):
        return "内标", note

    if inner_system_label:
        return "外标", "文件名提示内标为系统标，当前模板按外标建议"

    return "未识别", note


def strip_noise(stem: str) -> str:
    text = normalize_text(stem)
    for word in NOISE_WORDS:
        text = text.replace(word, " ")
    return text


def extract_match_code(stem: str) -> str:
    cleaned = strip_noise(stem)
    tokens = re.findall(r"[A-Za-z0-9][A-Za-z0-9_.\-]*[A-Za-z0-9]|[A-Za-z0-9]", cleaned)
    tokens = [token.strip("._-") for token in tokens if len(token.strip("._-")) >= 3]
    if not tokens:
        return ""

    # Prefer the longest technical-looking code. This handles files where "内标/外标"
    # appears before the actual customer part number.
    tokens.sort(key=lambda token: (len(token), bool(re.search(r"[A-Za-z]", token))), reverse=True)
    return tokens[0]


def is_inactive(path: Path) -> bool:
    full_path = normalize_text(str(path))
    return contains_any(full_path, INACTIVE_KEYWORDS)


def scan_templates(template_root: Path, relative_root: Path | None = None) -> List[MappingRow]:
    template_root = template_root.resolve()
    relative_root = (relative_root or template_root).resolve()
    files = sorted(
        {
            path
            for pattern in ("*.btw", "*.btW", "*.BTW")
            for path in template_root.rglob(pattern)
            if path.is_file()
        },
        key=lambda path: str(path).casefold(),
    )

    rows: List[MappingRow] = []
    for path in files:
        try:
            relative = path.relative_to(relative_root)
        except ValueError:
            relative = path.relative_to(template_root)
        customer_dir = relative.parts[0] if relative.parts else ""
        stem = path.stem
        label_type, label_note = detect_label_type(stem, relative_path=relative)
        match_code = extract_match_code(stem)
        inactive = is_inactive(path)

        issues: List[str] = []
        if inactive:
            issues.append("路径包含旧版/勿用关键字")
        if not match_code:
            issues.append("未识别匹配编号")
        if label_type == "未识别":
            if match_code:
                label_type = "外标"
                issues.append("未出现内/外标字样，按外标建议")
            else:
                issues.append("未识别标签类型")
        if label_note:
            issues.append(label_note)

        rows.append(
            MappingRow(
                customer_code="",
                manual_match_code=match_code,
                manual_label_type=label_type if label_type != "未识别" else "",
                customer_dir=customer_dir,
                file_name=path.name,
                relative_path=str(relative),
                issue="；".join(issues),
                status="待定",
                enabled_suggestion="否" if inactive else "是",
                match_code=match_code,
                label_type=label_type,
            )
        )

    duplicate_keys = Counter(
        (row.customer_dir, row.match_code.casefold(), row.label_type)
        for row in rows
        if row.match_code and row.label_type != "未识别" and row.enabled_suggestion == "是"
    )

    for row in rows:
        key = (row.customer_dir, row.match_code.casefold(), row.label_type)
        issues = [item for item in row.issue.split("；") if item]
        if row.enabled_suggestion == "否":
            row.status = "排除"
        elif not row.match_code or row.label_type == "未识别":
            row.status = "需要确认"
        elif any("按外标建议" in item for item in issues):
            row.status = "需要确认"
        elif duplicate_keys[key] > 1:
            row.status = "需要确认"
            issues.append("同客户+编号+标签类型存在多个模板")
        else:
            row.status = "自动确认"
        row.issue = "；".join(dict.fromkeys(issues))

    return rows


def column_letter(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def column_index(column: str) -> int:
    index = 0
    for char in column.upper():
        if "A" <= char <= "Z":
            index = index * 26 + ord(char) - ord("A") + 1
    return index


def write_csv(rows: List[MappingRow], output_path: Path) -> None:
    with output_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(HEADERS)
        writer.writerows(row.as_list() for row in rows)


def get_column(row: Dict[str, str], names: Iterable[str]) -> str:
    for name in names:
        if name in row:
            return row[name]
    return ""


def normalize_relative_path(value: str) -> str:
    return normalize_text(value).replace("/", "\\").strip().casefold()


def convert_existing_row(row: Dict[str, str]) -> MappingRow:
    status = get_column(row, ("状态",))
    enabled_suggestion = get_column(row, ("启用建议", "启用"))
    match_code = ""
    label_type = ""
    issue = get_column(row, ("问题",))
    if not status:
        match = re.search(r"(?:^|；)状态：([^；]+)", issue)
        status = match.group(1) if match else ""
    if not enabled_suggestion:
        match = re.search(r"(?:^|；)启用建议：([^；]+)", issue)
        enabled_suggestion = match.group(1) if match else ""
    return MappingRow(
        customer_code=get_column(row, ("客户编号", "ERP客户编号")),
        manual_match_code=get_column(row, ("匹配编号",)),
        manual_label_type=get_column(row, ("标签类型",)),
        customer_dir=get_column(row, ("客户目录",)),
        file_name=get_column(row, ("模板文件名",)),
        relative_path=get_column(row, ("模板相对路径", "模版相对路径", "模板路径", "模版路径", "相对路径")),
        issue=issue,
        status=status,
        enabled_suggestion=enabled_suggestion,
        match_code=match_code,
        label_type=label_type,
    )


def read_existing_rows(path: Path) -> List[MappingRow]:
    if not path.exists():
        return []

    from template_mapping_lookup import read_mapping_rows

    return [convert_existing_row(row) for row in read_mapping_rows(path)]


def merge_existing_and_scanned(existing_rows: List[MappingRow], scanned_rows: List[MappingRow]) -> List[MappingRow]:
    existing_keys = {
        normalize_relative_path(row.relative_path)
        for row in existing_rows
        if row.relative_path
    }
    new_rows = [
        row
        for row in scanned_rows
        if normalize_relative_path(row.relative_path) not in existing_keys
    ]
    return [*existing_rows, *new_rows]


def filter_new_rows(existing_rows: List[MappingRow], scanned_rows: List[MappingRow]) -> List[MappingRow]:
    existing_keys = {
        normalize_relative_path(row.relative_path)
        for row in existing_rows
        if row.relative_path
    }
    return [
        row
        for row in scanned_rows
        if normalize_relative_path(row.relative_path) not in existing_keys
    ]


def xml_cell(row_idx: int, col_idx: int, value: str, style: int = 0) -> str:
    ref = f"{column_letter(col_idx)}{row_idx}"
    value = escape(str(value or ""))
    style_attr = f' s="{style}"' if style else ""
    return f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{value}</t></is></c>'


def write_minimal_xlsx(rows: List[MappingRow], output_path: Path) -> None:
    data = [list(HEADERS), *[row.as_list() for row in rows]]
    widths = [16, 22, 14, 18, 70, 80]
    row_count = len(data)
    col_count = len(HEADERS)
    last_cell = f"{column_letter(col_count)}{row_count}"

    sheet_rows: List[str] = []
    for row_idx, values in enumerate(data, start=1):
        cells = []
        for col_idx, value in enumerate(values, start=1):
            if row_idx > 1 and value == "":
                continue
            style = 1 if row_idx == 1 else 0
            if row_idx > 1 and col_idx == len(HEADERS):
                if "状态：自动确认" in value:
                    style = 2
                elif "状态：需要确认" in value:
                    style = 3
                elif "状态：排除" in value:
                    style = 4
            cells.append(xml_cell(row_idx, col_idx, value, style))
        sheet_rows.append(f'<row r="{row_idx}">{"".join(cells)}</row>')

    cols = "".join(
        f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>'
        for idx, width in enumerate(widths, start=1)
    )

    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>{cols}</cols>
  <sheetData>{''.join(sheet_rows)}</sheetData>
  <autoFilter ref="A1:{last_cell}"/>
</worksheet>'''

    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="模板映射" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Arial"/></font>
    <font><b/><sz val="11"/><name val="Arial"/></font>
  </fonts>
  <fills count="5">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF4CCCC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    workbook_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    core_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Label Generator</dc:creator>
  <cp:lastModifiedBy>Label Generator</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''

    app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Label Generator</Application>
</Properties>'''

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as xlsx:
        xlsx.writestr("[Content_Types].xml", content_types_xml)
        xlsx.writestr("_rels/.rels", rels_xml)
        xlsx.writestr("xl/workbook.xml", workbook_xml)
        xlsx.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        xlsx.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        xlsx.writestr("xl/styles.xml", styles_xml)
        xlsx.writestr("docProps/core.xml", core_xml)
        xlsx.writestr("docProps/app.xml", app_xml)


def build_sheet_row_xml(row_idx: int, values: List[str]) -> str:
    cells = []
    for col_idx, value in enumerate(values, start=1):
        if value == "":
            continue
        cells.append(xml_cell(row_idx, col_idx, value))
    return f'<row r="{row_idx}">{"".join(cells)}</row>'


def update_range_last_row(xml_text: str, last_row: int) -> str:
    def replace_ref(match) -> str:
        ref = match.group(1)
        updated = re.sub(r"(\$?[A-Z]+\$?)\d+", rf"\g<1>{last_row}", ref)
        return f'ref="{updated}"'

    return re.sub(r'ref="([^"]+)"', replace_ref, xml_text)


def read_shared_strings(workbook: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []

    text = workbook.read("xl/sharedStrings.xml").decode("utf-8")
    values: List[str] = []
    for item in re.findall(r"<si\b[^>]*>(.*?)</si>", text, flags=re.S):
        parts = re.findall(r"<t(?:\s[^>]*)?>(.*?)</t>", item, flags=re.S)
        values.append(unescape("".join(parts)))
    return values


def get_cell_text(cell_xml: str, shared_strings: List[str]) -> str:
    cell_type_match = re.search(r'\bt="([^"]+)"', cell_xml)
    cell_type = cell_type_match.group(1) if cell_type_match else ""

    if cell_type == "inlineStr":
        parts = re.findall(r"<t(?:\s[^>]*)?>(.*?)</t>", cell_xml, flags=re.S)
        return unescape("".join(parts))

    value_match = re.search(r"<v>(.*?)</v>", cell_xml, flags=re.S)
    if not value_match:
        return ""

    value = unescape(value_match.group(1))
    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (ValueError, IndexError):
            return ""
    return value


def find_cell_xml(row_xml: str, row_idx: int, column: str) -> str:
    pattern = rf'<c\b(?=[^>]*\br="{column}{row_idx}")[^>]*(?:>.*?</c>|/>)'
    match = re.search(pattern, row_xml, flags=re.S)
    return match.group(0) if match else ""


def replace_or_insert_cell(row_xml: str, row_idx: int, col_idx: int, value: str) -> str:
    column = column_letter(col_idx)
    cell_pattern = rf'<c\b(?=[^>]*\br="{column}{row_idx}")[^>]*(?:>.*?</c>|/>)'
    new_cell = xml_cell(row_idx, col_idx, value)

    if re.search(cell_pattern, row_xml, flags=re.S):
        return re.sub(cell_pattern, new_cell, row_xml, count=1, flags=re.S)

    cells = list(re.finditer(r'<c\b[^>]*\br="([A-Z]+)\d+"[^>]*(?:>.*?</c>|/>)', row_xml, flags=re.S))
    for cell in cells:
        if column_index(cell.group(1)) > col_idx:
            return row_xml[: cell.start()] + new_cell + row_xml[cell.start() :]

    return row_xml.replace("</row>", new_cell + "</row>")


def replace_xlsx_file(temp_path: Path, output_path: Path) -> None:
    try:
        temp_path.replace(output_path)
    except PermissionError as exc:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise RuntimeError(
            f"无法更新 {output_path}：文件可能正在被 Excel/WPS 打开。"
            "请关闭 template_mapping.xlsx 后重新运行脚本。"
        ) from exc


def read_sheet_header_columns(sheet_xml: str, shared_strings: List[str]) -> Dict[str, str]:
    header_match = re.search(r'<row\b[^>]*\br="1"[^>]*>.*?</row>', sheet_xml, flags=re.S)
    if not header_match:
        return {}

    columns: Dict[str, str] = {}
    for cell in re.finditer(r'<c\b[^>]*\br="([A-Z]+)1"[^>]*(?:>.*?</c>|/>)', header_match.group(0), flags=re.S):
        text = get_cell_text(cell.group(0), shared_strings).strip()
        if text:
            columns[text] = cell.group(1)
    return columns


def find_header_column(header_columns: Dict[str, str], names: Iterable[str]) -> str:
    for name in names:
        column = header_columns.get(name)
        if column:
            return column
    return ""


def backfill_existing_generated_fields(output_path: Path, scanned_rows: List[MappingRow]) -> int:
    scanned_by_path = {
        normalize_relative_path(row.relative_path): row
        for row in scanned_rows
        if row.relative_path
    }
    if not scanned_by_path:
        return 0

    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    updated_count = 0
    with zipfile.ZipFile(output_path, "r") as source:
        shared_strings = read_shared_strings(source)
        sheet_xml = source.read("xl/worksheets/sheet1.xml").decode("utf-8")
        header_columns = read_sheet_header_columns(sheet_xml, shared_strings)
        relative_path_column = find_header_column(
            header_columns,
            ("模板相对路径", "模版相对路径", "模板路径", "模版路径", "相对路径"),
        )
        match_code_column = find_header_column(header_columns, ("匹配编号",))
        label_type_column = find_header_column(header_columns, ("标签类型",))
        if not relative_path_column or not match_code_column or not label_type_column:
            return 0

        def update_row(match) -> str:
            nonlocal updated_count
            row_xml = match.group(0)
            row_idx = int(match.group(1))
            if row_idx == 1:
                return row_xml

            relative_path = get_cell_text(find_cell_xml(row_xml, row_idx, relative_path_column), shared_strings)
            scanned = scanned_by_path.get(normalize_relative_path(relative_path))
            if scanned is None:
                return row_xml

            match_code = get_cell_text(find_cell_xml(row_xml, row_idx, match_code_column), shared_strings).strip()
            label_type = get_cell_text(find_cell_xml(row_xml, row_idx, label_type_column), shared_strings).strip()
            changed = False

            if not match_code and scanned.match_code:
                row_xml = replace_or_insert_cell(row_xml, row_idx, column_index(match_code_column), scanned.match_code)
                changed = True
            if not label_type and scanned.label_type and scanned.label_type != "未识别":
                row_xml = replace_or_insert_cell(row_xml, row_idx, column_index(label_type_column), scanned.label_type)
                changed = True

            if changed:
                updated_count += 1
            return row_xml

        sheet_xml = re.sub(r'<row\b[^>]*\br="(\d+)"[^>]*>.*?</row>', update_row, sheet_xml, flags=re.S)

        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                if item.filename == "xl/worksheets/sheet1.xml":
                    target.writestr(item, sheet_xml)
                else:
                    target.writestr(item, source.read(item.filename))

    if updated_count:
        replace_xlsx_file(temp_path, output_path)
    else:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass
    return updated_count


def append_rows_to_existing_xlsx(output_path: Path, rows: List[MappingRow]) -> None:
    if not rows:
        return

    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    with zipfile.ZipFile(output_path, "r") as source:
        sheet_xml = source.read("xl/worksheets/sheet1.xml").decode("utf-8")
        row_numbers = [int(value) for value in re.findall(r'<row[^>]*\sr="(\d+)"', sheet_xml)]
        next_row = max(row_numbers, default=1) + 1
        new_sheet_rows = "".join(
            build_sheet_row_xml(row_idx, row.as_list())
            for row_idx, row in enumerate(rows, start=next_row)
        )
        sheet_xml = sheet_xml.replace("</sheetData>", new_sheet_rows + "</sheetData>")
        last_row = next_row + len(rows) - 1
        sheet_xml = re.sub(
            r'<autoFilter\s+ref="[^"]+"\s*/>',
            lambda match: update_range_last_row(match.group(0), last_row),
            sheet_xml,
        )
        sheet_xml = re.sub(
            r'<dimension\s+ref="[^"]+"\s*/>',
            lambda match: update_range_last_row(match.group(0), last_row),
            sheet_xml,
        )

        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                if item.filename == "xl/worksheets/sheet1.xml":
                    target.writestr(item, sheet_xml)
                else:
                    target.writestr(item, source.read(item.filename))

    replace_xlsx_file(temp_path, output_path)


def is_path_relative_to(path: Path, base: Path) -> bool:
    try:
        path.resolve().relative_to(base.resolve())
        return True
    except ValueError:
        return False


def update_template_mapping_from_directory(
    template_root: Path,
    output: Optional[Path] = None,
) -> Dict[str, object]:
    template_root = Path(template_root)
    output = Path(output or (app_paths.TEMPLATE_MAPPING_FILE if app_paths is not None else DEFAULT_OUTPUT))

    if not template_root.exists():
        raise FileNotFoundError(f"模板目录不存在：{template_root}")

    relative_root = DEFAULT_TEMPLATE_ROOT if is_path_relative_to(template_root, DEFAULT_TEMPLATE_ROOT) else template_root
    scanned_rows = scan_templates(template_root, relative_root=relative_root)
    if output.exists():
        existing_rows = read_existing_rows(output)
        new_rows = filter_new_rows(existing_rows, scanned_rows)
        if new_rows:
            append_rows_to_existing_xlsx(output, new_rows)
        total_rows = len(existing_rows) + len(new_rows)
    else:
        existing_rows = []
        new_rows = scanned_rows
        total_rows = len(scanned_rows)
        write_minimal_xlsx(scanned_rows, output)

    counts = Counter(row.status for row in [*existing_rows, *new_rows])
    return {
        "template_root": str(template_root.resolve()),
        "output": str(output.resolve()),
        "scanned": len(scanned_rows),
        "existing": len(existing_rows),
        "added": len(new_rows),
        "total": total_rows,
        "auto_confirmed": counts["自动确认"],
        "need_confirm": counts["需要确认"],
        "excluded": counts["排除"],
    }


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate initial BarTender template mapping workbook.")
    parser.add_argument("--template-root", default=str(DEFAULT_TEMPLATE_ROOT), help="模板根目录")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="输出 .xlsx 路径")
    parser.add_argument("--csv-output", default="", help="可选：同时输出 CSV 路径")
    return parser


def main() -> int:
    args = build_arg_parser().parse_args()
    template_root = Path(args.template_root)
    output = Path(args.output)

    if not template_root.exists():
        raise SystemExit(f"模板根目录不存在：{template_root}")

    result = update_template_mapping_from_directory(template_root, output)
    scanned_rows = scan_templates(template_root)
    existing_rows = read_existing_rows(output) if output.exists() else []
    new_rows = filter_new_rows(existing_rows[: result["existing"]], scanned_rows) if result["added"] else []
    rows = existing_rows
    backfilled_count = 0

    if args.csv_output:
        write_csv(rows, Path(args.csv_output))

    print(f"扫描模板数：{result['scanned']}")
    print(f"已有映射数：{result['existing']}")
    print(f"本次新增：{result['added']}")
    print(f"本次回填匹配编号/标签类型：{backfilled_count}")
    print(f"输出总数：{result['total']}")
    print(f"自动确认：{result['auto_confirmed']}")
    print(f"需要确认：{result['need_confirm']}")
    print(f"排除：{result['excluded']}")
    print(f"已生成：{result['output']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
