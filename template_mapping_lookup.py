from __future__ import annotations

import argparse
import csv
import re
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple, Union
from xml.etree import ElementTree


MAPPING_FILE = Path(__file__).resolve().parent / "template_mapping.xlsx"


class TemplateLookupError(Exception):
    pass


def normalize_code(value: str) -> str:
    return (value or "").strip().casefold()


def read_mapping_rows(mapping_file: Union[str, Path] = MAPPING_FILE) -> List[Dict[str, str]]:
    path = Path(mapping_file)
    if not path.exists():
        raise TemplateLookupError(f"映射表不存在：{path}")

    if path.suffix.lower() == ".csv":
        with path.open(newline="", encoding="utf-8-sig") as file:
            return list(csv.DictReader(file))

    if path.suffix.lower() == ".xlsx":
        return read_xlsx_first_sheet(path)

    raise TemplateLookupError(f"不支持的映射表格式：{path.suffix}")


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
                values_by_col[col_index] = value
            max_col = max(values_by_col, default=0)
            table.append([values_by_col.get(index, "") for index in range(1, max_col + 1)])

    if not table:
        return []

    headers = table[0]
    return [
        {headers[index]: row[index] if index < len(row) else "" for index in range(len(headers))}
        for row in table[1:]
    ]


def xlsx_column_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref)
    if not match:
        return 1
    index = 0
    for char in match.group(1):
        index = index * 26 + ord(char) - ord("A") + 1
    return index


def get_column(row: Dict[str, str], names: Tuple[str, ...]) -> str:
    for name in names:
        if name in row:
            return row[name]
    return ""


def read_runtime_csv_first_row(runtime_csv: Union[str, Path]) -> Dict[str, str]:
    path = Path(runtime_csv)
    if not path.exists():
        raise TemplateLookupError(f"runtime CSV 不存在：{path}")

    with path.open(newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        row = next(reader, None)

    if row is None:
        raise TemplateLookupError(f"runtime CSV 没有数据行：{path}")
    return row


def find_template_by_mo_fields(
    mo_no: str,
    customer_code: str,
    customer_part_no: str,
    product_code: str,
    label_type: str,
    mapping_file: Union[str, Path] = MAPPING_FILE,
) -> str:
    return get_column(
        find_template_mapping_by_mo_fields(
            mo_no=mo_no,
            customer_code=customer_code,
            customer_part_no=customer_part_no,
            product_code=product_code,
            label_type=label_type,
            mapping_file=mapping_file,
        ),
        ("模板相对路径", "模版相对路径", "模板路径", "模版路径", "相对路径"),
    )


def find_template_mapping_by_mo_fields(
    mo_no: str,
    customer_code: str,
    customer_part_no: str,
    product_code: str,
    label_type: str,
    mapping_file: Union[str, Path] = MAPPING_FILE,
) -> Dict[str, str]:
    """Return template relative path by ERP fields.

    Search order:
    1. Filter rows by customer code.
    2. Filter rows by label type.
    3. Search mapping code by customer part number.
    4. If not found, search mapping code by product code.
    5. Return the unique matched template relative path.
    """

    rows = read_mapping_rows(mapping_file)
    if not rows:
        raise TemplateLookupError(f"映射表没有数据：{mapping_file}")

    if not any("客户编号" in row or "ERP客户编号" in row for row in rows):
        raise TemplateLookupError("映射表缺少“客户编号”列，请先在 Excel 里补 ERP 客户编号。")

    enabled_rows = [
        row
        for row in rows
        if get_column(row, ("启用建议", "启用")) in ("", "是", "Y", "y", "yes", "YES", "1")
        and get_column(row, ("状态",)) != "排除"
    ]

    customer_rows = [
        row
        for row in enabled_rows
        if normalize_code(get_column(row, ("客户编号", "ERP客户编号"))) == normalize_code(customer_code)
    ]
    if not customer_rows:
        raise TemplateLookupError(f"MO {mo_no} 未找到客户编号对应的模板配置：{customer_code}")

    label_rows = [
        row
        for row in customer_rows
        if normalize_code(get_column(row, ("标签类型",))) == normalize_code(label_type)
    ]
    if not label_rows:
        raise TemplateLookupError(f"MO {mo_no} 未找到标签类型对应的模板配置：{label_type}")

    for source_name, code in (("客户料号", customer_part_no), ("产品编号", product_code)):
        if not normalize_code(code):
            continue

        matches = [
            row
            for row in label_rows
            if normalize_code(get_column(row, ("匹配编号",))) == normalize_code(code)
        ]

        if len(matches) == 1:
            relative_path = get_column(
                matches[0],
                ("模板相对路径", "模版相对路径", "模板路径", "模版路径", "相对路径"),
            )
            if not relative_path:
                raise TemplateLookupError(f"MO {mo_no} 找到模板行，但缺少模板相对路径。")
            return matches[0]

        if len(matches) > 1:
            raise TemplateLookupError(
                f"MO {mo_no} 使用{source_name} {code} 找到多个模板，当前版本要求唯一匹配。"
            )

    raise TemplateLookupError(
        f"MO {mo_no} 未找到模板。客户编号={customer_code}，"
        f"标签类型={label_type}，客户料号={customer_part_no}，产品编号={product_code}"
    )


def find_template_by_runtime_csv(
    runtime_csv: Union[str, Path],
    label_type: str,
    mapping_file: Union[str, Path] = MAPPING_FILE,
) -> str:
    return get_column(
        find_template_mapping_by_runtime_csv(
            runtime_csv=runtime_csv,
            label_type=label_type,
            mapping_file=mapping_file,
        ),
        ("模板相对路径", "模版相对路径", "模板路径", "模版路径", "相对路径"),
    )


def find_template_mapping_by_runtime_csv(
    runtime_csv: Union[str, Path],
    label_type: str,
    mapping_file: Union[str, Path] = MAPPING_FILE,
) -> Dict[str, str]:
    row = read_runtime_csv_first_row(runtime_csv)
    return find_template_mapping_by_mo_fields(
        mo_no=get_column(row, ("MO_NO", "MO批次号")),
        customer_code=get_column(row, ("CUS_NO", "CUSTOMER_CODE", "ERP_CUSTOMER_CODE", "公司编号")),
        customer_part_no=get_column(row, ("SUP_PRD_NO", "CUS_ITEM_NO", "CUSTOMER_PART_NO", "CUST_PN", "客户料号")),
        product_code=get_column(row, ("MRP_NO", "PRD_NO", "PRODUCT_CODE", "PART_NO", "ITEM_NO", "产品编号")),
        label_type=label_type,
        mapping_file=mapping_file,
    )


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="按 MO 字段从模板映射表查找 BarTender 模板。")
    parser.add_argument("--mo-no", required=True, help="MO号")
    parser.add_argument("--customer-code", required=True, help="ERP客户编号")
    parser.add_argument("--customer-part-no", default="", help="客户料号，优先用于匹配")
    parser.add_argument("--product-code", default="", help="产品编号，客户料号找不到时用于匹配")
    parser.add_argument("--label-type", required=True, help="标签类型，例如：内标、外标、内外标、送货单")
    parser.add_argument("--mapping-file", default=str(MAPPING_FILE), help="模板映射表路径，支持 .xlsx 或 .csv")
    return parser


def main() -> int:
    args = build_arg_parser().parse_args()
    try:
        template_path = find_template_by_mo_fields(
            mo_no=args.mo_no,
            customer_code=args.customer_code,
            customer_part_no=args.customer_part_no,
            product_code=args.product_code,
            label_type=args.label_type,
            mapping_file=args.mapping_file,
        )
    except TemplateLookupError as error:
        print("查找失败")
        print(error)
        return 1

    print("查找成功")
    print(f"MO号：{args.mo_no}")
    print(f"客户编号：{args.customer_code}")
    print(f"标签类型：{args.label_type}")
    print(f"客户料号：{args.customer_part_no or '-'}")
    print(f"产品编号：{args.product_code or '-'}")
    print(f"模板相对路径：{template_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
