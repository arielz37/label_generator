from __future__ import annotations

import argparse
import calendar
import csv
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from config import ERP_DATABASE, ERP_DRIVER, ERP_PASSWORD, ERP_SERVER, ERP_USER, require_erp_config


RUNTIME_DIR = Path(__file__).resolve().parent / "runtime_data"

CSV_FIELDS = [
    "MO_NO",
    "CUS_NO",
    "SUP_PRD_NO",
    "MRP_NO",
    "MFG_DATE",
    "EXP_DATE",
    "HAS_INNER_LABEL",
    "INNER_QTY",
    "OUTER_QTY",
    "INNER_LABEL_COUNT",
    "OUTER_LABEL_COUNT",
]

DB_FIELD_ALIASES = {
    "MO_DD": "MFG_DATE",
    "CUS_NO": "CUSTOMER_CODE",
    "CUS_NAME": "CUSTOMER_NAME",
    "SUP_PRD_NO": "SUP_PRD_NO",
    "MRP_NO": "MRP_NO",
    "PRD_NO": "PRODUCT_CODE",
    "PRD_NAME": "PRODUCT_NAME",
}


DEFAULT_MO_QUERY = """
SELECT TOP 1
    MO_NO,
    MO_DD AS MFG_DATE,
    MO_DD AS PROD_DATE,
    SO_NO,
    CUS_NO,
    SUP_PRD_NO,
    MRP_NO,
    QTY
FROM MF_MO
WHERE MO_NO = ?
ORDER BY MO_DD DESC
"""


class RuntimeCsvError(Exception):
    pass


def connect_erp():
    import pyodbc

    require_erp_config()
    return pyodbc.connect(
        f"DRIVER={{{ERP_DRIVER}}};"
        f"SERVER={ERP_SERVER};"
        f"DATABASE={ERP_DATABASE};"
        f"UID={ERP_USER};"
        f"PWD={ERP_PASSWORD};"
        "Encrypt=no;"
        "TrustServerCertificate=yes;"
    )


def format_csv_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, Decimal):
        return format(value, "f").rstrip("0").rstrip(".")
    return str(value).strip()


def parse_decimal(value: Any) -> Optional[Decimal]:
    text = format_csv_value(value)
    if not text:
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def format_decimal(value: Decimal, places: int = 4) -> str:
    quant = Decimal("1").scaleb(-places)
    rounded = value.quantize(quant, rounding=ROUND_HALF_UP)
    return format(rounded, "f").rstrip("0").rstrip(".")


def calculate_label_count(total_qty: Any, per_label_qty: Any) -> str:
    total = parse_decimal(total_qty)
    per_label = parse_decimal(per_label_qty)
    if total is None or per_label is None:
        return ""
    if per_label == 0:
        raise RuntimeCsvError("产品数量不能为 0，无法计算标签张数。")
    return format_decimal(total / per_label, places=4)


def parse_csv_date(value: str) -> Optional[date]:
    value = (value or "").strip()
    if not value:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    return None


def add_months(start_date: date, months: int) -> date:
    month_index = start_date.month - 1 + months
    year = start_date.year + month_index // 12
    month = month_index % 12 + 1
    day = min(start_date.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def calculate_exp_date(mfg_date: str, shelf_life: str) -> str:
    start_date = parse_csv_date(mfg_date)
    shelf_life = (shelf_life or "").strip()
    if not start_date or not shelf_life:
        return ""

    match = re.fullmatch(r"(\d+)\s*(天|日|D|d|day|days)?", shelf_life)
    if match and match.group(2):
        return format_csv_value(datetime.fromordinal(start_date.toordinal() + int(match.group(1))))

    match = re.fullmatch(r"(\d+)\s*(月|个月|M|m|month|months)?", shelf_life)
    if match:
        return format_csv_value(add_months(start_date, int(match.group(1))))

    match = re.fullmatch(r"(\d+)\s*(年|Y|y|year|years)", shelf_life)
    if match:
        return format_csv_value(add_months(start_date, int(match.group(1)) * 12))

    raise RuntimeCsvError(f"无法识别保质期格式：{shelf_life}")


def normalize_db_row(row: Dict[str, Any]) -> Dict[str, str]:
    normalized = {
        "MO_NO": "",
        "SO_NO": "",
        "CUSTOMER_CODE": "",
        "CUSTOMER_NAME": "",
        "CUSTOMER_PART_NO": "",
        "SUP_PRD_NO": "",
        "PRODUCT_CODE": "",
        "MRP_NO": "",
        "PRODUCT_NAME": "",
        "SPEC": "",
        "QTY": "",
        "MFG_DATE": "",
        "PROD_DATE": "",
        "EXP_DATE": "",
        "EXPIRE_DATE": "",
        "BOX_NO": "",
        "CAN_NO": "",
        "SERIAL_NO": "",
        "REMARK": "",
    }
    for key, value in row.items():
        field = DB_FIELD_ALIASES.get(key.upper(), key.upper())
        if field in normalized:
            normalized[field] = format_csv_value(value)

    # Keep common aliases populated so BarTender templates can bind either spelling.
    if normalized["MFG_DATE"] and not normalized["PROD_DATE"]:
        normalized["PROD_DATE"] = normalized["MFG_DATE"]
    if normalized["PROD_DATE"] and not normalized["MFG_DATE"]:
        normalized["MFG_DATE"] = normalized["PROD_DATE"]
    if normalized["EXP_DATE"] and not normalized["EXPIRE_DATE"]:
        normalized["EXPIRE_DATE"] = normalized["EXP_DATE"]
    if normalized["EXPIRE_DATE"] and not normalized["EXP_DATE"]:
        normalized["EXP_DATE"] = normalized["EXPIRE_DATE"]

    return normalized


def first_value(row: Dict[str, str], keys: Tuple[str, ...]) -> str:
    for key in keys:
        value = row.get(key, "")
        if value:
            return value
    return ""


def build_runtime_csv_row(
    normalized_row: Dict[str, str],
    runtime_options: Optional[Dict[str, Any]] = None,
) -> Dict[str, str]:
    options = runtime_options or {}
    qty = first_value(normalized_row, ("QTY",))
    mfg_date = first_value(normalized_row, ("MFG_DATE", "PROD_DATE"))
    exp_date = (
        format_csv_value(options.get("EXP_DATE"))
        or calculate_exp_date(mfg_date, format_csv_value(options.get("SHELF_LIFE")))
        or first_value(normalized_row, ("EXP_DATE", "EXPIRE_DATE"))
    )

    inner_label_count = format_csv_value(options.get("INNER_LABEL_COUNT", "1"))
    outer_label_count = format_csv_value(options.get("OUTER_LABEL_COUNT", "1"))
    has_inner_label = format_csv_value(options.get("HAS_INNER_LABEL")) or ("Y" if inner_label_count != "0" else "N")
    inner_qty = format_csv_value(options.get("INNER_QTY")) or qty
    outer_qty = format_csv_value(options.get("OUTER_QTY")) or qty

    if "INNER_LABEL_COUNT" not in options:
        inner_label_count = "0" if has_inner_label == "N" else calculate_label_count(qty, inner_qty)
    if "OUTER_LABEL_COUNT" not in options:
        outer_label_count = calculate_label_count(qty, outer_qty)

    return {
        "MO_NO": first_value(normalized_row, ("MO_NO",)),
        "CUS_NO": first_value(normalized_row, ("CUSTOMER_CODE",)),
        "SUP_PRD_NO": first_value(normalized_row, ("SUP_PRD_NO", "CUSTOMER_PART_NO")),
        "MRP_NO": first_value(normalized_row, ("MRP_NO", "PRODUCT_CODE")),
        "MFG_DATE": mfg_date,
        "EXP_DATE": exp_date,
        "HAS_INNER_LABEL": has_inner_label,
        "INNER_QTY": inner_qty,
        "OUTER_QTY": outer_qty,
        "INNER_LABEL_COUNT": inner_label_count,
        "OUTER_LABEL_COUNT": outer_label_count,
    }


def fetch_mo_rows(mo_no: str) -> List[Dict[str, Any]]:
    conn = connect_erp()
    try:
        cursor = conn.cursor()
        cursor.execute(DEFAULT_MO_QUERY, mo_no)
        columns = [column[0] for column in cursor.description]
        return [dict(zip(columns, row)) for row in cursor.fetchall()]
    finally:
        conn.close()


def write_runtime_csv(
    rows: List[Dict[str, str]],
    output_path: Union[str, Path],
    fields: Optional[Union[List[str], Tuple[str, ...]]] = None,
) -> Path:
    if not rows:
        raise RuntimeCsvError("没有可写入 CSV 的数据。")

    fieldnames = list(fields or CSV_FIELDS)
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    return path


def generate_runtime_csv_for_mo(
    mo_no: str,
    output_path: Optional[Union[str, Path]] = None,
    fields: Optional[Union[List[str], Tuple[str, ...]]] = None,
    runtime_options: Optional[Dict[str, Any]] = None,
) -> Path:
    db_rows = fetch_mo_rows(mo_no)
    if not db_rows:
        raise RuntimeCsvError(f"未查询到 MO 数据：{mo_no}")

    csv_rows = [
        build_runtime_csv_row(normalize_db_row(row), runtime_options=runtime_options)
        for row in db_rows
    ]
    output = Path(output_path) if output_path else RUNTIME_DIR / "current_label.csv"
    return write_runtime_csv(csv_rows, output, fields=fields)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="从 ERP SQL Server 查询 MO 数据并生成 BarTender runtime CSV。")
    parser.add_argument("--mo-no", required=True, help="MO号")
    parser.add_argument("--output", default="", help="输出 CSV 路径，默认写入 runtime_data/")
    return parser


def main() -> int:
    args = build_arg_parser().parse_args()
    try:
        output = generate_runtime_csv_for_mo(
            mo_no=args.mo_no,
            output_path=args.output or None,
        )
    except RuntimeCsvError as error:
        print("生成失败")
        print(error)
        return 1

    print("生成成功")
    print(f"MO号：{args.mo_no}")
    print(f"CSV路径：{output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
