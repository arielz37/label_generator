from __future__ import annotations

import csv
import re
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple, Union
from xml.etree import ElementTree


PRODUCT_LABEL_DETAILS_FILE = Path(__file__).resolve().parent / "product_label_details.xlsx"


class ProductLabelDetailsError(Exception):
    pass


def normalize_code(value: str) -> str:
    return (value or "").strip().casefold()


def get_column(row: Dict[str, str], names: Tuple[str, ...]) -> str:
    for name in names:
        if name in row:
            return row[name]
    return ""


def xlsx_column_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref)
    if not match:
        return 1
    index = 0
    for char in match.group(1):
        index = index * 26 + ord(char) - ord("A") + 1
    return index


def read_xlsx_first_sheet(path: Path) -> List[Dict[str, str]]:
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path) as workbook:
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in workbook.namelist():
            shared_root = ElementTree.fromstring(workbook.read("xl/sharedStrings.xml"))
            for item in shared_root.findall("main:si", ns):
                shared_strings.append("".join(text.text or "" for text in item.findall(".//main:t", ns)))

        sheet_root = ElementTree.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
        table: List[List[str]] = []
        for row in sheet_root.findall(".//main:sheetData/main:row", ns):
            values_by_col: Dict[int, str] = {}
            for cell in row.findall("main:c", ns):
                col_index = xlsx_column_index(cell.attrib.get("r", "A1"))
                cell_type = cell.attrib.get("t")
                value_node = cell.find("main:v", ns)
                inline_node = cell.find("main:is/main:t", ns)
                if cell_type == "s" and value_node is not None:
                    value = shared_strings[int(value_node.text or "0")]
                elif inline_node is not None:
                    value = inline_node.text or ""
                elif value_node is not None:
                    value = value_node.text or ""
                else:
                    value = ""
                values_by_col[col_index] = value.strip()
            max_col = max(values_by_col, default=0)
            table.append([values_by_col.get(index, "") for index in range(1, max_col + 1)])

    if not table:
        return []

    headers = table[0]
    return [
        {headers[index]: row[index] if index < len(row) else "" for index in range(len(headers))}
        for row in table[1:]
        if any(row)
    ]


def read_detail_rows(details_file: Union[str, Path] = PRODUCT_LABEL_DETAILS_FILE) -> List[Dict[str, str]]:
    path = Path(details_file)
    if not path.exists():
        raise ProductLabelDetailsError(f"产品标签细节文件不存在：{path}")

    if path.suffix.lower() == ".csv":
        with path.open(newline="", encoding="utf-8-sig") as file:
            return list(csv.DictReader(file))

    if path.suffix.lower() == ".xlsx":
        return read_xlsx_first_sheet(path)

    raise ProductLabelDetailsError(f"不支持的产品标签细节文件格式：{path.suffix}")


def read_runtime_csv_first_row(runtime_csv: Union[str, Path]) -> Dict[str, str]:
    path = Path(runtime_csv)
    if not path.exists():
        raise ProductLabelDetailsError(f"runtime CSV 不存在：{path}")

    with path.open(newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        row = next(reader, None)

    if row is None:
        raise ProductLabelDetailsError(f"runtime CSV 没有数据行：{path}")
    return row


def find_product_label_detail_by_runtime_csv(
    runtime_csv: Union[str, Path],
    details_file: Union[str, Path] = PRODUCT_LABEL_DETAILS_FILE,
) -> Dict[str, str]:
    runtime_row = read_runtime_csv_first_row(runtime_csv)
    customer_code = get_column(runtime_row, ("CUS_NO", "CUSTOMER_CODE", "ERP_CUSTOMER_CODE", "客户编号"))
    customer_part_no = get_column(runtime_row, ("SUP_PRD_NO", "CUS_ITEM_NO", "CUSTOMER_PART_NO", "CUST_PN", "客户料号"))
    product_code = get_column(runtime_row, ("MRP_NO", "PRD_NO", "PRODUCT_CODE", "PART_NO", "ITEM_NO", "产品编码", "产品编号"))
    rows = read_detail_rows(details_file)

    customer_rows = [
        row
        for row in rows
        if normalize_code(get_column(row, ("客户编号", "CUS_NO", "CUSTOMER_CODE"))) == normalize_code(customer_code)
    ]
    if not customer_rows:
        return {}

    search_rules = (
        ("客户料号", customer_part_no, ("客户料号", "SUP_PRD_NO", "CUS_ITEM_NO", "CUSTOMER_PART_NO")),
        ("产品编码", product_code, ("产品编码", "MRP_NO", "PRD_NO", "PRODUCT_CODE")),
    )

    for source_name, code, columns in search_rules:
        if not normalize_code(code):
            continue
        matches = [
            row
            for row in customer_rows
            if normalize_code(get_column(row, columns)) == normalize_code(code)
        ]
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            # Duplicate rows with the same shelf life and label settings are harmless for now.
            signatures = {
                (
                    get_column(row, ("保质期（月）", "保质期", "SHELF_LIFE")),
                    get_column(row, ("外标性质（系统标：1，模版标：2）", "OUTER_LABEL_MODE")),
                    get_column(row, ("内标性质（无内标：0，系统标：1，模版标：2）", "INNER_LABEL_MODE")),
                )
                for row in matches
            }
            if len(signatures) == 1:
                return matches[0]
            raise ProductLabelDetailsError(f"{source_name} {code} 找到多条不同产品标签细节。")

    return {}


def find_shelf_life_by_runtime_csv(
    runtime_csv: Union[str, Path],
    details_file: Union[str, Path] = PRODUCT_LABEL_DETAILS_FILE,
) -> str:
    detail = find_product_label_detail_by_runtime_csv(runtime_csv, details_file=details_file)
    return get_column(detail, ("保质期（月）", "保质期", "SHELF_LIFE"))


def build_runtime_options_from_detail(detail: Dict[str, str]) -> Dict[str, str]:
    inner_mode = get_column(detail, ("内标性质（无内标：0，系统标：1，模版标：2）", "INNER_LABEL_MODE"))
    return {
        "SHELF_LIFE": get_column(detail, ("保质期（月）", "保质期", "SHELF_LIFE")),
        "HAS_INNER_LABEL": "N" if normalize_code(inner_mode) == "0" else "Y",
        "INNER_QTY": get_column(detail, ("产品数量（内标）（pcs）", "产品数量（内标）", "INNER_QTY")),
        "OUTER_QTY": get_column(detail, ("产品数量（外标）", "OUTER_QTY")),
        "INNER_LABEL_MODE": inner_mode,
        "OUTER_LABEL_MODE": get_column(detail, ("外标性质（系统标：1，模版标：2）", "OUTER_LABEL_MODE")),
    }
