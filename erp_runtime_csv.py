from __future__ import annotations

import argparse
import calendar
import csv
import re
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from app_paths import PACKAGE_NAME_FILE, RUNTIME_DIR
from config import ERP_DATABASE, ERP_DRIVER, ERP_PASSWORD, ERP_SERVER, ERP_USER, require_erp_config


DEFAULT_SHELF_LIFE_MONTHS = 12
CHINESE_CHAR_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
EDGE_SEPARATOR_CHARS = " \t\r\n-_/,，;；:：|()（）[]【】{}"

RAW_RUNTIME_FIELDS = [
    "MO_NO",
    "SO_NO",
    "CUS_OS_NO",
    "ORDER_NO",
    "CUS_NO",
    "SUP_PRD_NO",
    "MRP_NO",
    "MFG_QTY",
    "MFG_DATE",
    "MFG_DATE_YYMMDD",
    "MFG_DATE_YYYYMMDD",
    "MFG_DATE_DD_MM_YYYY_DOT",
    "MFG_DATE_YYYY_MM_DD_DOT",
    "MFG_DATE_YY_MM_DD_DOT",
    "EXP_DATE",
    "EXP_DATE_YYMMDD",
    "EXP_DATE_YYYYMMDD",
    "EXP_DATE_DD_MM_YYYY_DOT",
    "EXP_DATE_YYYY_MM_DD_DOT",
    "EXP_DATE_YY_MM_DD_DOT",
    "LOT_NO",
    "HAS_INNER_LABEL",
    "INNER_QTY",
    "OUTER_QTY",
    "INNER_LABEL_COUNT",
    "OUTER_LABEL_COUNT",
    "INNER_PACKAGE_PRD_NO",
    "INNER_PACKAGE_NAME",
    "OUTER_PACKAGE_PRD_NO",
    "OUTER_PACKAGE_NAME",
]

LABEL_CSV_FIELDS = [
    "MO_NO",
    "SO_NO",
    "CUS_OS_NO",
    "ORDER_NO",
    "CUS_NO",
    "SUP_PRD_NO",
    "MRP_NO",
    "MFG_QTY",
    "MFG_DATE",
    "MFG_DATE_YYMMDD",
    "MFG_DATE_YYYYMMDD",
    "MFG_DATE_DD_MM_YYYY_DOT",
    "MFG_DATE_YYYY_MM_DD_DOT",
    "MFG_DATE_YY_MM_DD_DOT",
    "EXP_DATE",
    "EXP_DATE_YYMMDD",
    "EXP_DATE_YYYYMMDD",
    "EXP_DATE_DD_MM_YYYY_DOT",
    "EXP_DATE_YYYY_MM_DD_DOT",
    "EXP_DATE_YY_MM_DD_DOT",
    "LOT_NO",
    "HAS_INNER_LABEL",
    "QTY",
    "LABEL_COUNT",
    "LABEL_INDEX",
    "PACKAGE_PRD_NO",
    "PACKAGE_NAME",
]

# Default file shape for runtime_data/current_label.csv. BarTender templates
# should bind these current-label fields, not the internal INNER_/OUTER_ fields.
CSV_FIELDS = LABEL_CSV_FIELDS

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
    CUS_OS_NO,
    CUS_NO,
    SUP_PRD_NO,
    MRP_NO,
    QTY
FROM MF_MO
WHERE MO_NO = ?
ORDER BY MO_DD DESC
"""

PACKAGE_CACHE: Dict[str, Dict[str, str]] = {}

OUTER_PACKAGE_NAME_KEYWORDS = ("箱",)
OUTER_PACKAGE_EXCLUDE_KEYWORDS = ("周转箱",)
CAN_PACKAGE_NAME_KEYWORDS = ("铁桶", "罐")
DEFAULT_INNER_PACKAGE_NAMES = (
    "铁桶",
    "公司自用透明PE袋",
    "双层防潮袋",
    "铝箔袋",
    "自用粉色pe袋",
    "公司自用粉红色静电袋",
)

BOM_PACKAGE_QUERY = """
SELECT
    m.BOM_NO,
    m.PRD_NO,
    t.ITM,
    t.COMPONENT_PRD_NO,
    t.COMPONENT_NAME,
    t.QTY_BAS
FROM (
    SELECT
        t.PRD_NO AS COMPONENT_PRD_NO,
        t.NAME AS COMPONENT_NAME,
        t.QTY_BAS,
        t.ITM,
        t.BOM_NO
    FROM TF_BOM t
) t
INNER JOIN MF_BOM m
    ON m.BOM_NO = t.BOM_NO
WHERE m.PRD_NO = ?
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


def contains_chinese(value: Any) -> bool:
    return bool(CHINESE_CHAR_RE.search(format_csv_value(value)))


def format_non_chinese_csv_value(value: Any) -> str:
    text = format_csv_value(value)
    if not text:
        return ""
    text = CHINESE_CHAR_RE.sub("", text).strip(EDGE_SEPARATOR_CHARS)
    return text


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


def parse_package_name_line(line: str) -> str:
    text = line.strip()
    if not text or text.startswith("#"):
        return ""

    parts = text.split(maxsplit=1)
    if len(parts) == 2:
        return parts[1].strip()
    return text


def load_inner_package_names(path: Optional[Path] = None) -> Tuple[str, ...]:
    path = Path(path) if path is not None else PACKAGE_NAME_FILE
    names: List[str] = []
    if path.exists():
        for line in path.read_text(encoding="utf-8-sig").splitlines():
            name = parse_package_name_line(line)
            if name:
                names.append(name)

    if not names:
        names.extend(DEFAULT_INNER_PACKAGE_NAMES)
    return tuple(dict.fromkeys(names))


def is_can_package_name(package_name: str) -> bool:
    return any(keyword in package_name for keyword in CAN_PACKAGE_NAME_KEYWORDS)


def is_inner_package_name(package_name: str, inner_package_names: Tuple[str, ...]) -> bool:
    name = package_name.strip()
    return is_can_package_name(name) or name in inner_package_names


def is_outer_package_name(package_name: str) -> bool:
    name = package_name.strip()
    return (
        any(keyword in name for keyword in OUTER_PACKAGE_NAME_KEYWORDS)
        and not any(keyword in name for keyword in OUTER_PACKAGE_EXCLUDE_KEYWORDS)
    )


def quantity_sort_value(row: Dict[str, Any]) -> Decimal:
    return parse_decimal(row.get("QTY_BAS")) or Decimal("0")


def item_sort_value(row: Dict[str, Any]) -> Decimal:
    return parse_decimal(row.get("ITM")) or Decimal("0")


def select_inner_package_row(rows: List[Dict[str, Any]], inner_package_names: Tuple[str, ...]) -> Optional[Dict[str, Any]]:
    candidates = [
        row
        for row in rows
        if is_inner_package_name(format_csv_value(row.get("COMPONENT_NAME")), inner_package_names)
    ]
    if not candidates:
        return None

    return sorted(
        candidates,
        key=lambda row: (
            0 if is_can_package_name(format_csv_value(row.get("COMPONENT_NAME"))) else 1,
            -quantity_sort_value(row),
            item_sort_value(row),
        ),
    )[0]


def select_outer_package_row(rows: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    candidates = [
        row
        for row in rows
        if is_outer_package_name(format_csv_value(row.get("COMPONENT_NAME")))
    ]
    if not candidates:
        return None

    return sorted(
        candidates,
        key=lambda row: (
            -quantity_sort_value(row),
            item_sort_value(row),
        ),
    )[0]


def select_package_row_by_name(rows: List[Dict[str, Any]], package_name: str) -> Optional[Dict[str, Any]]:
    name = (package_name or "").strip()
    if not name:
        return None
    candidates = [
        row
        for row in rows
        if format_csv_value(row.get("COMPONENT_NAME")).strip() == name
    ]
    if not candidates:
        return None

    return sorted(
        candidates,
        key=lambda row: (
            -quantity_sort_value(row),
            item_sort_value(row),
        ),
    )[0]


def apply_package_row(result: Dict[str, str], prefix: str, row: Optional[Dict[str, Any]]) -> None:
    if row is None:
        return

    package_qty = format_csv_value(row.get("QTY_BAS"))
    if not package_qty:
        return

    result[f"{prefix}_QTY"] = package_qty
    result[f"{prefix}_PACKAGE_PRD_NO"] = format_csv_value(row.get("COMPONENT_PRD_NO"))
    result[f"{prefix}_PACKAGE_NAME"] = format_csv_value(row.get("COMPONENT_NAME"))


def fetch_bom_material_rows(product_code: str) -> List[Dict[str, Any]]:
    product_code = (product_code or "").strip()
    if not product_code:
        return []

    conn = connect_erp()
    try:
        cursor = conn.cursor()
        cursor.execute(BOM_PACKAGE_QUERY, product_code)
        columns = [column[0] for column in cursor.description]
        return [dict(zip(columns, row)) for row in cursor.fetchall()]
    finally:
        conn.close()


def fetch_bom_material_names(product_code: str) -> List[str]:
    names: List[str] = []
    for row in fetch_bom_material_rows(product_code):
        name = format_csv_value(row.get("COMPONENT_NAME")).strip()
        if name:
            names.append(name)
    return list(dict.fromkeys(names))


def fetch_bom_material_names_for_mo(mo_no: str) -> List[str]:
    db_rows = fetch_mo_rows(mo_no)
    if not db_rows:
        raise RuntimeCsvError(f"未查询到 MO 数据：{mo_no}")
    normalized = normalize_db_row(db_rows[0])
    product_code = first_value(normalized, ("MRP_NO", "PRODUCT_CODE"))
    return fetch_bom_material_names(product_code)


def fetch_bom_pack_quantities(
    product_code: str,
    inner_package_name: str = "",
    outer_package_name: str = "",
) -> Dict[str, str]:
    product_code = (product_code or "").strip()
    if not product_code:
        return {}

    rows = fetch_bom_material_rows(product_code)
    result: Dict[str, str] = {}
    inner_package_names = load_inner_package_names()
    inner_row = (
        select_package_row_by_name(rows, inner_package_name)
        if inner_package_name
        else select_inner_package_row(rows, inner_package_names)
    )
    outer_row = (
        select_package_row_by_name(rows, outer_package_name)
        if outer_package_name
        else select_outer_package_row(rows)
    )

    if inner_package_name and inner_row is None:
        raise RuntimeCsvError(f"BOM 中未找到 UI 指定的内标包装：{inner_package_name}")
    if outer_package_name and outer_row is None:
        raise RuntimeCsvError(f"BOM 中未找到 UI 指定的外标包装：{outer_package_name}")

    apply_package_row(result, "INNER", inner_row)
    apply_package_row(result, "OUTER", outer_row)
    return result


def fetch_pack_quantities(
    product_code: str,
    inner_package_name: str = "",
    outer_package_name: str = "",
) -> Dict[str, str]:
    product_code = (product_code or "").strip()
    if not product_code:
        return {}

    cache_key = "|".join([
        product_code.casefold(),
        (inner_package_name or "").strip().casefold(),
        (outer_package_name or "").strip().casefold(),
    ])
    if cache_key in PACKAGE_CACHE:
        return PACKAGE_CACHE[cache_key].copy()

    result = fetch_bom_pack_quantities(
        product_code,
        inner_package_name=inner_package_name,
        outer_package_name=outer_package_name,
    )
    if result:
        result["PACKAGE_SOURCE"] = "BOM"

    PACKAGE_CACHE[cache_key] = result.copy()
    return result


def calculate_label_count(total_qty: Any, per_label_qty: Any) -> str:
    total = parse_decimal(total_qty)
    per_label = parse_decimal(per_label_qty)
    if total is None or per_label is None:
        return ""
    if per_label == 0:
        raise RuntimeCsvError("产品数量不能为 0，无法计算标签张数。")
    return format_decimal(total / per_label, places=4)


def calculate_remainder_qty(total_qty: Any, per_label_qty: Any) -> str:
    total = parse_decimal(total_qty)
    per_label = parse_decimal(per_label_qty)
    if total is None or per_label is None or per_label == 0:
        return ""

    remainder = total % per_label
    if remainder == 0:
        return ""
    return format_decimal(remainder)


def parse_label_count(value: str) -> int:
    value = (value or "1").strip()
    try:
        return max(int(float(value)), 0)
    except ValueError:
        raise RuntimeCsvError(f"无法识别 LABEL_COUNT：{value}")


def parse_csv_date(value: str) -> Optional[date]:
    value = (value or "").strip()
    if not value:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
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


def shelf_life_months_from_package_name(package_name: str) -> Optional[int]:
    return DEFAULT_SHELF_LIFE_MONTHS


def parse_shelf_life_months(value: Any) -> Optional[int]:
    text = format_csv_value(value)
    if not text:
        return None

    month_match = re.fullmatch(r"(\d+)\s*(个月|月|M|m|month|months)?", text)
    if month_match:
        months = int(month_match.group(1))
    else:
        year_match = re.fullmatch(r"(\d+)\s*(年|Y|y|year|years)", text)
        if not year_match:
            raise RuntimeCsvError(f"无法识别保质期：{text}，请填写例如 12、12个月、1年。")
        months = int(year_match.group(1)) * 12

    if months <= 0:
        raise RuntimeCsvError("保质期必须大于 0。")
    return months


def calculate_exp_date_by_months(mfg_date: str, months: Optional[int]) -> str:
    start_date = parse_csv_date(mfg_date)
    if not start_date or months is None:
        return ""
    return format_csv_value(add_months(start_date, months) - timedelta(days=1))


def format_date_compact(value: str, fmt: str) -> str:
    parsed = parse_csv_date(value)
    if not parsed:
        return ""
    return parsed.strftime(fmt)


def build_lot_no(mfg_date: str, mo_no: str, suffix: str = "SZBD") -> str:
    mfg_compact = format_date_compact(mfg_date, "%y%m%d")
    if not mfg_compact or not mo_no:
        return ""
    return f"{mfg_compact}-{mo_no}-{suffix}"


def calculate_exp_date_by_package(mfg_date: str, package_name: str) -> str:
    return calculate_exp_date_by_months(mfg_date, shelf_life_months_from_package_name(package_name))


def normalize_db_row(row: Dict[str, Any]) -> Dict[str, str]:
    normalized = {
        "MO_NO": "",
        "SO_NO": "",
        "CUS_OS_NO": "",
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


def first_non_chinese_value(row: Dict[str, str], keys: Tuple[str, ...]) -> str:
    for key in keys:
        value = format_non_chinese_csv_value(row.get(key, ""))
        if value:
            return value
    return ""


def build_runtime_csv_row(
    normalized_row: Dict[str, str],
    runtime_options: Optional[Dict[str, Any]] = None,
) -> Dict[str, str]:
    options = runtime_options or {}
    qty = first_value(normalized_row, ("QTY",))
    product_code = first_value(normalized_row, ("MRP_NO", "PRODUCT_CODE"))
    pack_options = fetch_pack_quantities(
        product_code,
        inner_package_name=format_csv_value(options.get("INNER_PACKAGE_NAME_OVERRIDE")),
        outer_package_name=format_csv_value(options.get("OUTER_PACKAGE_NAME_OVERRIDE")),
    )
    mfg_date = first_value(normalized_row, ("MFG_DATE", "PROD_DATE"))
    shelf_life_months = (
        parse_shelf_life_months(options.get("SHELF_LIFE"))
        or shelf_life_months_from_package_name(pack_options.get("INNER_PACKAGE_NAME", ""))
    )
    exp_date = (
        format_csv_value(options.get("EXP_DATE"))
        or calculate_exp_date_by_months(mfg_date, shelf_life_months)
        or first_value(normalized_row, ("EXP_DATE", "EXPIRE_DATE"))
    )

    inner_qty = format_csv_value(options.get("INNER_QTY")) or pack_options.get("INNER_QTY", "")
    outer_qty = format_csv_value(options.get("OUTER_QTY")) or pack_options.get("OUTER_QTY", "")
    if not outer_qty:
        raise RuntimeCsvError(f"BOM 未找到外标包装用量基数，无法生成外标数量：{product_code}")
    has_inner_label = format_csv_value(options.get("HAS_INNER_LABEL")) or ("Y" if inner_qty else "N")
    inner_label_count = format_csv_value(options.get("INNER_LABEL_COUNT", ""))
    outer_label_count = format_csv_value(options.get("OUTER_LABEL_COUNT", ""))

    if "INNER_LABEL_COUNT" not in options:
        inner_label_count = calculate_label_count(qty, inner_qty) if has_inner_label == "Y" and inner_qty else "0"
    if "OUTER_LABEL_COUNT" not in options:
        outer_label_count = calculate_label_count(qty, outer_qty)
    inner_remainder_qty = calculate_remainder_qty(qty, inner_qty) if has_inner_label == "Y" and inner_qty else ""
    outer_remainder_qty = calculate_remainder_qty(qty, outer_qty)

    return {
        "MO_NO": first_non_chinese_value(normalized_row, ("MO_NO",)),
        "SO_NO": first_non_chinese_value(normalized_row, ("SO_NO",)),
        "CUS_OS_NO": first_non_chinese_value(normalized_row, ("CUS_OS_NO",)),
        "ORDER_NO": first_value(normalized_row, ("CUS_OS_NO",)),
        "CUS_NO": first_non_chinese_value(normalized_row, ("CUSTOMER_CODE",)),
        "SUP_PRD_NO": first_non_chinese_value(normalized_row, ("SUP_PRD_NO", "CUSTOMER_PART_NO")),
        "MRP_NO": format_non_chinese_csv_value(product_code),
        "MFG_QTY": qty,
        "MFG_DATE": mfg_date,
        "MFG_DATE_YYMMDD": format_date_compact(mfg_date, "%y%m%d"),
        "MFG_DATE_YYYYMMDD": format_date_compact(mfg_date, "%Y%m%d"),
        "MFG_DATE_DD_MM_YYYY_DOT": format_date_compact(mfg_date, "%d.%m.%Y"),
        "MFG_DATE_YYYY_MM_DD_DOT": format_date_compact(mfg_date, "%Y.%m.%d"),
        "MFG_DATE_YY_MM_DD_DOT": format_date_compact(mfg_date, "%y.%m.%d"),
        "EXP_DATE": exp_date,
        "EXP_DATE_YYMMDD": format_date_compact(exp_date, "%y%m%d"),
        "EXP_DATE_YYYYMMDD": format_date_compact(exp_date, "%Y%m%d"),
        "EXP_DATE_DD_MM_YYYY_DOT": format_date_compact(exp_date, "%d.%m.%Y"),
        "EXP_DATE_YYYY_MM_DD_DOT": format_date_compact(exp_date, "%Y.%m.%d"),
        "EXP_DATE_YY_MM_DD_DOT": format_date_compact(exp_date, "%y.%m.%d"),
        "LOT_NO": build_lot_no(mfg_date, first_value(normalized_row, ("MO_NO",))),
        "HAS_INNER_LABEL": has_inner_label,
        "INNER_QTY": inner_qty,
        "OUTER_QTY": outer_qty,
        "INNER_LABEL_COUNT": inner_label_count,
        "OUTER_LABEL_COUNT": outer_label_count,
        "INNER_REMAINDER_QTY": inner_remainder_qty,
        "OUTER_REMAINDER_QTY": outer_remainder_qty,
        "INNER_PACKAGE_PRD_NO": pack_options.get("INNER_PACKAGE_PRD_NO", ""),
        "INNER_PACKAGE_NAME": pack_options.get("INNER_PACKAGE_NAME", ""),
        "OUTER_PACKAGE_PRD_NO": pack_options.get("OUTER_PACKAGE_PRD_NO", ""),
        "OUTER_PACKAGE_NAME": pack_options.get("OUTER_PACKAGE_NAME", ""),
        "PACKAGE_SOURCE": pack_options.get("PACKAGE_SOURCE", ""),
        "INNER_PACKAGE_MISSING": "",
        "SHELF_LIFE_MONTHS": str(shelf_life_months or ""),
    }


def build_label_csv_row(
    runtime_row: Dict[str, str],
    qty_key: str,
    count_key: str,
    package_prd_key: str,
    package_name_key: str,
    label_index: int,
) -> Dict[str, str]:
    return {
        "MO_NO": runtime_row.get("MO_NO", ""),
        "SO_NO": runtime_row.get("SO_NO", ""),
        "CUS_OS_NO": runtime_row.get("CUS_OS_NO", ""),
        "ORDER_NO": runtime_row.get("ORDER_NO", ""),
        "CUS_NO": runtime_row.get("CUS_NO", ""),
        "SUP_PRD_NO": runtime_row.get("SUP_PRD_NO", ""),
        "MRP_NO": runtime_row.get("MRP_NO", ""),
        "MFG_QTY": runtime_row.get("MFG_QTY", ""),
        "MFG_DATE": runtime_row.get("MFG_DATE", ""),
        "MFG_DATE_YYMMDD": runtime_row.get("MFG_DATE_YYMMDD", ""),
        "MFG_DATE_YYYYMMDD": runtime_row.get("MFG_DATE_YYYYMMDD", ""),
        "MFG_DATE_DD_MM_YYYY_DOT": runtime_row.get("MFG_DATE_DD_MM_YYYY_DOT", ""),
        "MFG_DATE_YYYY_MM_DD_DOT": runtime_row.get("MFG_DATE_YYYY_MM_DD_DOT", ""),
        "MFG_DATE_YY_MM_DD_DOT": runtime_row.get("MFG_DATE_YY_MM_DD_DOT", ""),
        "EXP_DATE": runtime_row.get("EXP_DATE", ""),
        "EXP_DATE_YYMMDD": runtime_row.get("EXP_DATE_YYMMDD", ""),
        "EXP_DATE_YYYYMMDD": runtime_row.get("EXP_DATE_YYYYMMDD", ""),
        "EXP_DATE_DD_MM_YYYY_DOT": runtime_row.get("EXP_DATE_DD_MM_YYYY_DOT", ""),
        "EXP_DATE_YYYY_MM_DD_DOT": runtime_row.get("EXP_DATE_YYYY_MM_DD_DOT", ""),
        "EXP_DATE_YY_MM_DD_DOT": runtime_row.get("EXP_DATE_YY_MM_DD_DOT", ""),
        "LOT_NO": runtime_row.get("LOT_NO", ""),
        "HAS_INNER_LABEL": runtime_row.get("HAS_INNER_LABEL", ""),
        "QTY": runtime_row.get(qty_key, ""),
        "LABEL_COUNT": runtime_row.get(count_key, ""),
        "LABEL_INDEX": str(label_index),
        "PACKAGE_PRD_NO": runtime_row.get(package_prd_key, ""),
        "PACKAGE_NAME": runtime_row.get(package_name_key, ""),
    }


def build_label_csv_rows(runtime_rows: List[Dict[str, str]], label_type: str) -> List[Dict[str, str]]:
    if label_type == "outer":
        qty_key = "OUTER_QTY"
        count_key = "OUTER_LABEL_COUNT"
        package_prd_key = "OUTER_PACKAGE_PRD_NO"
        package_name_key = "OUTER_PACKAGE_NAME"
    elif label_type == "inner":
        qty_key = "INNER_QTY"
        count_key = "INNER_LABEL_COUNT"
        package_prd_key = "INNER_PACKAGE_PRD_NO"
        package_name_key = "INNER_PACKAGE_NAME"
    else:
        raise RuntimeCsvError(f"未知标签 CSV 类型：{label_type}")

    rows = []
    for row in runtime_rows:
        label_count = parse_label_count(row.get(count_key, "1"))
        for label_index in range(1, label_count + 1):
            rows.append(
                build_label_csv_row(
                    row,
                    qty_key=qty_key,
                    count_key=count_key,
                    package_prd_key=package_prd_key,
                    package_name_key=package_name_key,
                    label_index=label_index,
                )
            )
    return rows


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
    label_type: str = "outer",
) -> Path:
    db_rows = fetch_mo_rows(mo_no)
    if not db_rows:
        raise RuntimeCsvError(f"未查询到 MO 数据：{mo_no}")

    csv_rows = [
        build_runtime_csv_row(normalize_db_row(row), runtime_options=runtime_options)
        for row in db_rows
    ]
    label_rows = build_label_csv_rows(csv_rows, label_type)
    output = Path(output_path) if output_path else RUNTIME_DIR / "current_label.csv"
    return write_runtime_csv(label_rows, output, fields=fields or LABEL_CSV_FIELDS)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="从 ERP SQL Server 查询 MO 数据并生成 BarTender runtime CSV。")
    parser.add_argument("--mo-no", required=True, help="MO号")
    parser.add_argument("--output", default="", help="输出 CSV 路径，默认写入 runtime_data/")
    parser.add_argument(
        "--label-type",
        choices=("outer", "inner"),
        default="outer",
        help="生成外标或内标绑定 CSV，默认 outer",
    )
    return parser


def main() -> int:
    args = build_arg_parser().parse_args()
    try:
        output = generate_runtime_csv_for_mo(
            mo_no=args.mo_no,
            output_path=args.output or None,
            label_type=args.label_type,
        )
    except RuntimeCsvError as error:
        print("生成失败")
        print(error)
        return 1

    print("生成成功")
    print(f"MO号：{args.mo_no}")
    print(f"标签类型：{args.label_type}")
    print(f"CSV路径：{output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
