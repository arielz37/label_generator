from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from config import require_erp_config
from erp_runtime_csv import connect_erp, format_csv_value


DEFAULT_MO_NO = "MO26062470"


def quote_name(name: str) -> str:
    return "[" + name.replace("]", "]]") + "]"


def fetch_one(sql: str, *params: Any) -> Tuple[List[str], Dict[str, Any]]:
    conn = connect_erp()
    try:
        cursor = conn.cursor()
        cursor.execute(sql, params)
        row = cursor.fetchone()
        if row is None:
            return [], {}
        columns = [column[0] for column in cursor.description]
        return columns, dict(zip(columns, row))
    finally:
        conn.close()


def fetch_rows(sql: str, *params: Any) -> Tuple[List[str], List[Dict[str, Any]]]:
    conn = connect_erp()
    try:
        cursor = conn.cursor()
        cursor.execute(sql, params)
        columns = [column[0] for column in cursor.description]
        return columns, [dict(zip(columns, row)) for row in cursor.fetchall()]
    finally:
        conn.close()


def print_non_empty_fields(title: str, row: Dict[str, Any], highlight_prefix: str = "722") -> None:
    print(f"\n===== {title} =====")
    if not row:
        print("无数据")
        return

    for key in sorted(row.keys()):
        value = format_csv_value(row.get(key))
        if not value:
            continue
        marker = "  <<< 疑似订单号" if value.startswith(highlight_prefix) else ""
        print(f"{key:<32} {value}{marker}")


def find_tables_with_column(column_name: str, table_like: Iterable[str]) -> List[Dict[str, str]]:
    conditions = " OR ".join("UPPER(TABLE_NAME) LIKE ?" for _ in table_like)
    sql = f"""
SELECT DISTINCT TABLE_SCHEMA, TABLE_NAME
FROM INFORMATION_SCHEMA.COLUMNS
WHERE UPPER(COLUMN_NAME) = ?
  AND ({conditions})
ORDER BY TABLE_NAME
"""
    params = [column_name.upper()]
    params.extend(f"%{keyword.upper()}%" for keyword in table_like)
    columns, rows = fetch_rows(sql, *params)
    return rows


def inspect_related_tables(column_name: str, value: str, table_keywords: Tuple[str, ...], highlight_prefix: str) -> None:
    if not value:
        return
    tables = find_tables_with_column(column_name, table_keywords)
    if not tables:
        print(f"\n没有找到包含 {column_name} 的相关表")
        return

    print(f"\n===== 按 {column_name}={value} 查相关表 =====")
    for table in tables:
        schema = table["TABLE_SCHEMA"]
        table_name = table["TABLE_NAME"]
        sql = (
            f"SELECT TOP 1 * FROM {quote_name(schema)}.{quote_name(table_name)} "
            f"WHERE {quote_name(column_name)} = ?"
        )
        try:
            _, row = fetch_one(sql, value)
        except Exception:
            continue
        if row:
            print_non_empty_fields(f"{schema}.{table_name}", row, highlight_prefix=highlight_prefix)


def inspect_mo(mo_no: str, highlight_prefix: str) -> int:
    require_erp_config()
    _, mo_row = fetch_one("SELECT TOP 1 * FROM MF_MO WHERE MO_NO = ? ORDER BY MO_DD DESC", mo_no)
    print_non_empty_fields(f"MF_MO WHERE MO_NO={mo_no}", mo_row, highlight_prefix=highlight_prefix)
    if not mo_row:
        return 1

    so_no = format_csv_value(mo_row.get("SO_NO"))
    prd_no = format_csv_value(mo_row.get("MRP_NO") or mo_row.get("PRD_NO"))
    cus_no = format_csv_value(mo_row.get("CUS_NO"))

    print("\n===== 关键锚点 =====")
    print(f"MO_NO  = {mo_no}")
    print(f"SO_NO  = {so_no}")
    print(f"PRD_NO = {prd_no}")
    print(f"CUS_NO = {cus_no}")

    inspect_related_tables("SO_NO", so_no, ("SO", "OS", "ORD", "CUS", "TF", "MF"), highlight_prefix)
    inspect_related_tables("MO_NO", mo_no, ("MO", "SO", "OS", "ORD", "TF", "MF"), highlight_prefix)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="展开查看某个 MO 相关表字段，辅助定位客户订单号")
    parser.add_argument("--mo-no", default=DEFAULT_MO_NO, help="MO号，例如 MO26062470")
    parser.add_argument("--prefix", default="722", help="高亮疑似订单号前缀，默认 722")
    args = parser.parse_args()
    return inspect_mo(args.mo_no, args.prefix)


if __name__ == "__main__":
    raise SystemExit(main())
