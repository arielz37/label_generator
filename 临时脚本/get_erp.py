from __future__ import annotations

import argparse
import zipfile
from collections import OrderedDict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple
from xml.sax.saxutils import escape

import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from erp_runtime_csv import connect_erp, format_csv_value


OUTPUT_DIR = Path(__file__).resolve().parent

# 在这里批量填写客户名称和客户编号；同一个客户有多个编号时用 / 分隔。
# 命令行传参时会临时覆盖这里的设置，客户名称会显示为“命令行输入”。
CUSTOMERS = [
    ("南京华天", "NJHTKJ1759"),
    ("佐藤", "XGZTSS729"),
    ("上海怡凡得/台湾怡凡得/德国怡凡得/菲律宾怡凡得", "SHYFDD894/TWYFDD1165/DGYFDD1166/FLBYFD2260"),
    ("西安奕斯伟", "XAYSW1914"),
    ("西安欣芯", "XAXXCL2326"),
    ("新昇集团", "SHXSBD1497"),
    ("TECNOGLASS", "CLTECN1438"),
    ("科莱恩", "SHKLE1911"),
    ("马来西亚ON", "XJPJCLE220"),
    ("苏州瑞福龙", "UHRFL1304"),
    ("台湾日月光", "RYGBDT1401"),
    ("瑞仪光电", "UHRYGD2123"),
    ("哥伦比亚圣戈班", "GLBYSG1128"),
    ("杭州中欣", "HZZXJY1935"),
    ("丽水中欣", "LSZXJY2226"),
    ("上海阿泰泰珂", "SHASTK2449"),
    ("宁波甬矽半导/电子", "NBYXBD2275/NBYXDZ1718"),
    ("盛合晶微", "JYZXCD1198"),
    ("苏州嘉盛", "UHJSBD1562"),
    ("日月新集团", "KSRYGB770"),
    ("有研亿金", "SDYYYJ2411"),
    ("贝锝包装", "SHBXBZ909"),
    ("合肥晶合", "HFJHJC1399"),
    ("芯哲微电", "SHXZWD790"),
    ("上海蔡司/苏州蔡司", "SHCSGY393/UHCSKJ2059"),
    ("麦斯克", "HNMSK1944"),
    ("江苏芯德", "JSXDBD1948"),
    ("豪威", "SHHWDZ1116"),
    ("西安华羿", "XAHYWD1551"),
    ("杭州士兰集昕/杭州士兰集成/杭州士兰微", "HZSLJX1477/HZSLGD948/HZSLW2410"),
    ("上海伟测/无锡伟测/南京伟测/深圳伟测", "SHWCBD1856/WXWCBD1940/NJWCBD2090/SZWC2387"),
    ("Tech", "FGWTEC1142"),
    ("华润", "WXHRAS086"),
]

HEADERS = (
    "客户名称",
    "客户编号",
    "客户料号",
)

QUERY = """
SELECT
    CUS_NO,
    SUP_PRD_NO
FROM MF_MO
WHERE CUS_NO = ?
  AND SUP_PRD_NO IS NOT NULL
  AND LTRIM(RTRIM(SUP_PRD_NO)) <> ''
ORDER BY SUP_PRD_NO
"""

def split_customer_codes(value: str) -> List[str]:
    return [item.strip() for item in value.split("/") if item.strip()]


def configured_customers() -> List[Tuple[str, str]]:
    result: List[Tuple[str, str]] = []
    for customer_name, codes in CUSTOMERS:
        for customer_code in split_customer_codes(codes):
            result.append((customer_name.strip(), customer_code))
    return result


def group_customer_part_rows(raw_rows: List[Dict[str, Any]], customer_names: Dict[str, str]) -> List[Dict[str, str]]:
    grouped: "OrderedDict[str, Dict[str, str]]" = OrderedDict()
    for row in raw_rows:
        customer_code = format_csv_value(row.get("CUS_NO"))
        part_no = format_csv_value(row.get("SUP_PRD_NO"))
        if not part_no:
            continue

        key = f"{customer_code}\0{part_no}"
        if key not in grouped:
            grouped[key] = {
                "客户名称": customer_names.get(customer_code, ""),
                "客户编号": customer_code,
                "客户料号": part_no,
            }

    return list(grouped.values())


def fetch_customer_part_rows(customer_code: str) -> List[Dict[str, str]]:
    conn = connect_erp()
    try:
        cursor = conn.cursor()
        cursor.execute(QUERY, customer_code)
        columns = [column[0] for column in cursor.description]
        raw_rows = [dict(zip(columns, row)) for row in cursor.fetchall()]
    finally:
        conn.close()

    return group_customer_part_rows(raw_rows, {customer_code: ""})


def fetch_many_customer_part_rows(customers: Iterable[Tuple[str, str]]) -> List[Dict[str, str]]:
    customer_pairs = [(name.strip(), code.strip()) for name, code in customers if code.strip()]
    customer_names = {code: name for name, code in customer_pairs}
    conn = connect_erp()
    try:
        raw_rows: List[Dict[str, Any]] = []
        cursor = conn.cursor()
        for _, customer_code in customer_pairs:
            cursor.execute(QUERY, customer_code)
            columns = [column[0] for column in cursor.description]
            raw_rows.extend(dict(zip(columns, row)) for row in cursor.fetchall())
    finally:
        conn.close()

    return group_customer_part_rows(raw_rows, customer_names)


def column_letter(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def cell_xml(row_idx: int, col_idx: int, value: Any, style: int = 0) -> str:
    ref = f"{column_letter(col_idx)}{row_idx}"
    text = escape(format_csv_value(value))
    style_attr = f' s="{style}"' if style else ""
    return f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{text}</t></is></c>'


def write_xlsx(rows: List[Dict[str, str]], output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    data = [list(HEADERS)]
    data.extend([[row.get(header, "") for header in HEADERS] for row in rows])
    row_count = len(data)
    col_count = len(HEADERS)
    last_cell = f"{column_letter(col_count)}{row_count}"

    sheet_rows: List[str] = []
    for row_idx, values in enumerate(data, start=1):
        cells = [
            cell_xml(row_idx, col_idx, value, style=1 if row_idx == 1 else 0)
            for col_idx, value in enumerate(values, start=1)
        ]
        sheet_rows.append(f'<row r="{row_idx}">{"".join(cells)}</row>')

    widths = [34, 16, 32]
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
  <sheets><sheet name="客户料号" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Arial"/></font>
    <font><b/><sz val="11"/><name val="Arial"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>
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

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as xlsx:
        xlsx.writestr("[Content_Types].xml", content_types_xml)
        xlsx.writestr("_rels/.rels", rels_xml)
        xlsx.writestr("xl/workbook.xml", workbook_xml)
        xlsx.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        xlsx.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        xlsx.writestr("xl/styles.xml", styles_xml)
        xlsx.writestr("docProps/core.xml", core_xml)
        xlsx.writestr("docProps/app.xml", app_xml)

    return output_path


def safe_file_part(value: str) -> str:
    chars = []
    for char in value:
        if char.isalnum() or char in ("-", "_"):
            chars.append(char)
        else:
            chars.append("_")
    return "".join(chars).strip("_") or "customer"


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="按客户编号从 ERP 导出所有客户料号。")
    parser.add_argument("customer_code", nargs="*", help="可选：客户编号，例如 SHHXDZ001。为空时使用脚本顶部 CUSTOMERS。")
    parser.add_argument("--output", default="", help="输出 xlsx 路径，默认写入当前临时脚本目录")
    return parser


def main() -> int:
    args = build_arg_parser().parse_args()
    if args.customer_code:
        customers = [("命令行输入", item.strip()) for item in args.customer_code if item.strip()]
    else:
        customers = configured_customers()

    if not customers:
        print("客户编号不能为空：请填写 CUSTOMERS 或通过命令行传入客户编号。")
        return 1

    rows = fetch_many_customer_part_rows(customers)
    customer_codes = [code for _, code in customers]
    if not rows:
        print(f"未查询到客户料号：{', '.join(customer_codes)}")
        return 1

    output_name = "客户料号_批量导出.xlsx" if len(customer_codes) > 1 else f"{safe_file_part(customer_codes[0])}_客户料号.xlsx"
    output = Path(args.output) if args.output else OUTPUT_DIR / output_name
    write_xlsx(rows, output)

    print("导出成功")
    print(f"客户编号：{', '.join(customer_codes)}")
    print(f"客户料号数量：{len(rows)}")
    print(f"输出文件：{output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
