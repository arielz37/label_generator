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
from typing import Iterable, List
from xml.sax.saxutils import escape


DEFAULT_TEMPLATE_ROOT = Path(__file__).resolve().parent / "Templates"
DEFAULT_OUTPUT = Path(__file__).resolve().parent / "template_mapping.xlsx"

INACTIVE_KEYWORDS = ("旧版本", "勿用", "老版本", "暂停", "待确定", "不用")
INNER_KEYWORDS = ("内标", "内 标")
OUTER_KEYWORDS = ("外标", "外 标", "外箱", "外 箱")
DELIVERY_KEYWORDS = ("送货单", "发货清单")
PALLET_KEYWORDS = ("托盘",)

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
    "内外",
    "内标",
    "外标",
    "外箱",
    "送货单",
    "托盘标签",
    "托盘",
    "标签",
    "管装",
    "铝箔袋",
    "屏蔽袋",
)

HEADERS = (
    "状态",
    "客户目录",
    "子目录",
    "匹配编号",
    "标签类型",
    "模板文件名",
    "模板相对路径",
    "模板绝对路径",
    "启用建议",
    "问题",
    "备注",
)


@dataclass
class MappingRow:
    status: str
    customer_dir: str
    sub_dir: str
    match_code: str
    label_type: str
    file_name: str
    relative_path: str
    absolute_path: str
    enabled_suggestion: str
    issue: str
    remark: str = ""

    def as_list(self) -> List[str]:
        return [
            self.status,
            self.customer_dir,
            self.sub_dir,
            self.match_code,
            self.label_type,
            self.file_name,
            self.relative_path,
            self.absolute_path,
            self.enabled_suggestion,
            self.issue,
            self.remark,
        ]


def normalize_text(value: str) -> str:
    return unicodedata.normalize("NFKC", value or "")


def contains_any(value: str, keywords: Iterable[str]) -> bool:
    return any(keyword in value for keyword in keywords)


def detect_label_type(stem: str) -> str:
    text = normalize_text(stem)
    compact = re.sub(r"\s+", "", text)
    has_inner = contains_any(text, INNER_KEYWORDS)
    has_outer = contains_any(text, OUTER_KEYWORDS)
    if "内外标" in compact or "外内标" in compact:
        return "内外标"
    if contains_any(text, DELIVERY_KEYWORDS):
        return "送货单"
    if contains_any(text, PALLET_KEYWORDS):
        return "托盘"
    if has_inner and has_outer:
        return "内外标"
    if has_inner:
        return "内标"
    if has_outer:
        return "外标"
    return "未识别"


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


def scan_templates(template_root: Path) -> List[MappingRow]:
    template_root = template_root.resolve()
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
        relative = path.relative_to(template_root)
        customer_dir = relative.parts[0] if relative.parts else ""
        sub_dir = str(Path(*relative.parts[1:-1])) if len(relative.parts) > 2 else ""
        stem = path.stem
        label_type = detect_label_type(stem)
        match_code = extract_match_code(stem)
        inactive = is_inactive(path)

        issues: List[str] = []
        if inactive:
            issues.append("路径包含旧版/勿用关键字")
        if not match_code:
            issues.append("未识别匹配编号")
        if label_type == "未识别":
            issues.append("未识别标签类型")

        rows.append(
            MappingRow(
                status="待定",
                customer_dir=customer_dir,
                sub_dir=sub_dir,
                match_code=match_code,
                label_type=label_type,
                file_name=path.name,
                relative_path=str(relative),
                absolute_path=str(path.resolve()),
                enabled_suggestion="否" if inactive else "是",
                issue="；".join(issues),
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


def write_csv(rows: List[MappingRow], output_path: Path) -> None:
    with output_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(HEADERS)
        writer.writerows(row.as_list() for row in rows)


def xml_cell(row_idx: int, col_idx: int, value: str, style: int = 0) -> str:
    ref = f"{column_letter(col_idx)}{row_idx}"
    value = escape(str(value or ""))
    style_attr = f' s="{style}"' if style else ""
    return f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{value}</t></is></c>'


def write_minimal_xlsx(rows: List[MappingRow], output_path: Path) -> None:
    data = [list(HEADERS), *[row.as_list() for row in rows]]
    widths = [12, 18, 26, 22, 12, 46, 70, 95, 12, 44, 30]
    row_count = len(data)
    col_count = len(HEADERS)
    last_cell = f"{column_letter(col_count)}{row_count}"

    sheet_rows: List[str] = []
    for row_idx, values in enumerate(data, start=1):
        cells = []
        for col_idx, value in enumerate(values, start=1):
            style = 1 if row_idx == 1 else 0
            if row_idx > 1 and col_idx == 1:
                if value == "自动确认":
                    style = 2
                elif value == "需要确认":
                    style = 3
                elif value == "排除":
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

    rows = scan_templates(template_root)
    write_minimal_xlsx(rows, output)
    if args.csv_output:
        write_csv(rows, Path(args.csv_output))

    counts = Counter(row.status for row in rows)
    print(f"扫描模板数：{len(rows)}")
    print(f"自动确认：{counts['自动确认']}")
    print(f"需要确认：{counts['需要确认']}")
    print(f"排除：{counts['排除']}")
    print(f"已生成：{output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
