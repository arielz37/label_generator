from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Sequence


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from erp_runtime_csv import connect_erp, format_csv_value


# 在这里填一个确认存在“包装（一）/包装（二）”资料的产品编码。
PRODUCT_CODE = "TDLAS0010_SHHXDZ"

DESCRIPTION_KEYWORDS = ("包装", "包裝", "箱", "袋", "罐")
COLUMN_KEYWORDS = (
    "PACK",
    "PAK",
    "BOX",
    "CARTON",
    "UNIT",
    "QTY",
    "NUM",
    "PCS",
    "PKG",
)
PREFERRED_TABLE_KEYWORDS = ("PRD", "ITEM", "INV", "GOOD", "BAS")


def run_query(sql: str, params: Sequence[Any] = ()) -> List[Dict[str, str]]:
    conn = connect_erp()
    try:
        cursor = conn.cursor()
        cursor.execute(sql, *params)
        if cursor.description is None:
            return []
        columns = [column[0] for column in cursor.description]
        return [
            {column: format_csv_value(value) for column, value in zip(columns, row)}
            for row in cursor.fetchall()
        ]
    finally:
        conn.close()


def text_width(value: str) -> int:
    return sum(2 if ord(char) > 127 else 1 for char in value)


def print_table(title: str, rows: List[Dict[str, str]], limit: int = 80) -> None:
    print()
    print(title)
    print("=" * text_width(title))
    if not rows:
        print("无结果")
        return

    headers = list(rows[0].keys())
    display_rows = rows[:limit]
    widths = []
    for header in headers:
        width = text_width(header)
        for row in display_rows:
            width = max(width, text_width(row.get(header, "")))
        widths.append(min(width, 48))

    def pad(value: str, width: int) -> str:
        return value + " " * max(width - text_width(value), 0)

    print(" | ".join(pad(header, widths[index]) for index, header in enumerate(headers)))
    print("-+-".join("-" * width for width in widths))
    for row in display_rows:
        print(" | ".join(pad(row.get(header, ""), widths[index]) for index, header in enumerate(headers)))

    if len(rows) > limit:
        print(f"... 仅显示前 {limit} 行，共 {len(rows)} 行")
    else:
        print(f"共 {len(rows)} 行")


def like_any_sql(column_expr: str, keywords: Iterable[str]) -> str:
    return " OR ".join(f"{column_expr} LIKE ?" for _ in keywords)


def like_params(keywords: Iterable[str]) -> List[str]:
    return [f"%{keyword}%" for keyword in keywords]


def find_columns_by_description() -> List[Dict[str, str]]:
    where = like_any_sql("CAST(ep.value AS NVARCHAR(4000))", DESCRIPTION_KEYWORDS)
    sql = f"""
    SELECT TOP 200
        s.name AS 架构,
        t.name AS 表名,
        c.name AS 字段名,
        ty.name AS 类型,
        CAST(ep.value AS NVARCHAR(4000)) AS 字段说明
    FROM sys.tables t
    INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
    INNER JOIN sys.columns c ON t.object_id = c.object_id
    INNER JOIN sys.types ty ON c.user_type_id = ty.user_type_id
    INNER JOIN sys.extended_properties ep
        ON ep.major_id = c.object_id
       AND ep.minor_id = c.column_id
       AND ep.class = 1
    WHERE {where}
    ORDER BY t.name, c.column_id
    """
    return run_query(sql, like_params(DESCRIPTION_KEYWORDS))


def find_columns_by_name() -> List[Dict[str, str]]:
    where = like_any_sql("UPPER(c.name)", COLUMN_KEYWORDS)
    sql = f"""
    SELECT TOP 300
        s.name AS 架构,
        t.name AS 表名,
        c.name AS 字段名,
        ty.name AS 类型,
        c.column_id AS 顺序
    FROM sys.tables t
    INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
    INNER JOIN sys.columns c ON t.object_id = c.object_id
    INNER JOIN sys.types ty ON c.user_type_id = ty.user_type_id
    WHERE {where}
    ORDER BY
        CASE WHEN UPPER(t.name) LIKE '%PRD%' THEN 0 ELSE 1 END,
        t.name,
        c.column_id
    """
    return run_query(sql, like_params(COLUMN_KEYWORDS))


def find_product_tables() -> List[Dict[str, str]]:
    table_where = like_any_sql("UPPER(t.name)", PREFERRED_TABLE_KEYWORDS)
    column_where = like_any_sql("UPPER(c2.name)", COLUMN_KEYWORDS)
    sql = f"""
    SELECT TOP 120
        s.name AS 架构,
        t.name AS 表名,
        COUNT(*) AS 字段数,
        SUM(CASE WHEN UPPER(c.name) = 'PRD_NO' THEN 1 ELSE 0 END) AS 有PRD_NO,
        SUM(CASE WHEN {column_where} THEN 1 ELSE 0 END) AS 疑似包装字段数
    FROM sys.tables t
    INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
    INNER JOIN sys.columns c ON t.object_id = c.object_id
    INNER JOIN sys.columns c2 ON t.object_id = c2.object_id
    WHERE {table_where}
    GROUP BY s.name, t.name
    HAVING SUM(CASE WHEN UPPER(c.name) = 'PRD_NO' THEN 1 ELSE 0 END) > 0
    ORDER BY 疑似包装字段数 DESC, 表名
    """
    params = like_params(COLUMN_KEYWORDS) + like_params(PREFERRED_TABLE_KEYWORDS)
    return run_query(sql, params)


def get_table_columns(schema_name: str, table_name: str) -> List[str]:
    rows = run_query(
        """
        SELECT c.name AS 字段名
        FROM sys.schemas s
        INNER JOIN sys.tables t ON s.schema_id = t.schema_id
        INNER JOIN sys.columns c ON t.object_id = c.object_id
        WHERE s.name = ? AND t.name = ?
        ORDER BY c.column_id
        """,
        (schema_name, table_name),
    )
    return [row["字段名"] for row in rows]


def quote_name(value: str) -> str:
    return "[" + value.replace("]", "]]") + "]"


def fetch_product_rows_from_candidate_tables(product_code: str, tables: List[Dict[str, str]]) -> List[Dict[str, str]]:
    results: List[Dict[str, str]] = []
    for table in tables[:40]:
        schema_name = table["架构"]
        table_name = table["表名"]
        columns = get_table_columns(schema_name, table_name)
        upper_columns = {column.upper(): column for column in columns}
        prd_column = upper_columns.get("PRD_NO")
        if not prd_column:
            continue

        selected_columns = [
            column
            for column in columns
            if column.upper() == "PRD_NO"
            or any(keyword in column.upper() for keyword in COLUMN_KEYWORDS)
        ]
        if not selected_columns:
            selected_columns = columns[:30]

        sql = (
            "SELECT TOP 3 "
            + ", ".join(quote_name(column) for column in selected_columns)
            + f" FROM {quote_name(schema_name)}.{quote_name(table_name)}"
            + f" WHERE {quote_name(prd_column)} = ?"
        )
        rows = run_query(sql, (product_code,))
        for row in rows:
            compact = {
                key: value
                for key, value in row.items()
                if key.upper() == "PRD_NO" or value not in ("", "0", "0.0", "0.00")
            }
            if compact:
                results.append({"表名": f"{schema_name}.{table_name}", **compact})
    return results


def fetch_prdt_pack_field_values(product_code: str) -> List[Dict[str, str]]:
    columns = run_query(
        """
        SELECT
            c.column_id AS 顺序,
            c.name AS 字段名,
            ty.name AS 类型
        FROM sys.columns c
        INNER JOIN sys.types ty ON c.user_type_id = ty.user_type_id
        WHERE c.object_id = OBJECT_ID('dbo.PRDT')
          AND (
                UPPER(c.name) LIKE '%PK%'
             OR UPPER(c.name) LIKE '%PAK%'
             OR UPPER(c.name) LIKE '%PACK%'
             OR UPPER(c.name) LIKE '%UNIT%'
             OR UPPER(c.name) LIKE '%QTY%'
             OR UPPER(c.name) LIKE '%PCS%'
             OR UPPER(c.name) LIKE '%BOX%'
          )
        ORDER BY c.column_id
        """
    )
    if not columns:
        return []

    field_names = [row["字段名"] for row in columns]
    sql = (
        "SELECT TOP 1 "
        + ", ".join(quote_name(field_name) for field_name in field_names)
        + " FROM [dbo].[PRDT] WHERE [PRD_NO] = ?"
    )
    rows = run_query(sql, (product_code,))
    values = rows[0] if rows else {}

    return [
        {
            "顺序": row["顺序"],
            "字段名": row["字段名"],
            "类型": row["类型"],
            "值": values.get(row["字段名"], ""),
        }
        for row in columns
    ]


def fetch_prdt_all_non_empty_values(product_code: str) -> List[Dict[str, str]]:
    columns = get_table_columns("dbo", "PRDT")
    if not columns:
        return []

    sql = (
        "SELECT TOP 1 "
        + ", ".join(quote_name(column) for column in columns)
        + " FROM [dbo].[PRDT] WHERE [PRD_NO] = ?"
    )
    rows = run_query(sql, (product_code,))
    if not rows:
        return []

    values = rows[0]
    return [
        {"字段名": column, "值": values.get(column, "")}
        for column in columns
        if values.get(column, "") not in ("", "0", "0.0", "0.00")
    ]


def fetch_table_non_empty_rows(schema_name: str, table_name: str, product_code: str, max_rows: int = 5) -> List[Dict[str, str]]:
    columns = get_table_columns(schema_name, table_name)
    if not columns:
        return []
    upper_columns = {column.upper(): column for column in columns}
    prd_column = upper_columns.get("PRD_NO")
    if not prd_column:
        return []

    sql = (
        f"SELECT TOP {max_rows} "
        + ", ".join(quote_name(column) for column in columns)
        + f" FROM {quote_name(schema_name)}.{quote_name(table_name)}"
        + f" WHERE {quote_name(prd_column)} = ?"
    )
    rows = run_query(sql, (product_code,))

    result: List[Dict[str, str]] = []
    for row_index, row in enumerate(rows, start=1):
        for column in columns:
            value = row.get(column, "")
            if value in ("", "0", "0.0", "0.00"):
                continue
            result.append({
                "表名": f"{schema_name}.{table_name}",
                "行号": str(row_index),
                "字段名": column,
                "值": value,
            })
    return result


def fetch_pack_detail_table_values(product_code: str) -> List[Dict[str, str]]:
    tables = (
        ("dbo", "PRDT1"),
        ("dbo", "SH_PRDT1"),
        ("dbo", "PRDT1_MX"),
        ("dbo", "SH_PRDT1_MX"),
        ("dbo", "PRDT1_MX_ALL"),
        ("dbo", "SH_PRDT1_MX_ALL"),
        ("dbo", "SPRD1"),
        ("dbo", "SPRD"),
    )
    rows: List[Dict[str, str]] = []
    for schema_name, table_name in tables:
        rows.extend(fetch_table_non_empty_rows(schema_name, table_name, product_code))
    return rows


def fetch_tables_with_prd_no() -> List[Dict[str, str]]:
    return run_query(
        """
        SELECT
            s.name AS 架构,
            t.name AS 表名
        FROM sys.tables t
        INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
        INNER JOIN sys.columns c ON t.object_id = c.object_id
        WHERE UPPER(c.name) = 'PRD_NO'
        ORDER BY
            CASE WHEN UPPER(t.name) IN ('PRDT', 'PRDT1', 'SH_PRDT1', 'PRDT1_MX', 'SH_PRDT1_MX') THEN 0 ELSE 1 END,
            t.name
        """
    )


def fetch_rows_matching_values(product_code: str, needles: Sequence[str], table_limit: int = 180) -> List[Dict[str, str]]:
    tables = fetch_tables_with_prd_no()
    results: List[Dict[str, str]] = []
    normalized_needles = [needle.strip().casefold() for needle in needles if needle.strip()]
    if not normalized_needles:
        return results

    for table in tables[:table_limit]:
        schema_name = table["架构"]
        table_name = table["表名"]
        columns = get_table_columns(schema_name, table_name)
        upper_columns = {column.upper(): column for column in columns}
        prd_column = upper_columns.get("PRD_NO")
        if not prd_column:
            continue

        sql = (
            "SELECT TOP 10 "
            + ", ".join(quote_name(column) for column in columns)
            + f" FROM {quote_name(schema_name)}.{quote_name(table_name)}"
            + f" WHERE {quote_name(prd_column)} = ?"
        )
        try:
            rows = run_query(sql, (product_code,))
        except Exception as error:
            results.append({
                "表名": f"{schema_name}.{table_name}",
                "行号": "-",
                "字段名": "查询失败",
                "值": str(error),
            })
            continue

        for row_index, row in enumerate(rows, start=1):
            for column, value in row.items():
                text = str(value or "").strip()
                if not text:
                    continue
                folded = text.casefold()
                if any(needle in folded for needle in normalized_needles):
                    results.append({
                        "表名": f"{schema_name}.{table_name}",
                        "行号": str(row_index),
                        "字段名": column,
                        "值": text,
                    })
    return results


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="探测 ERP 货品基础资料里的包装字段。")
    parser.add_argument("--product-code", default=PRODUCT_CODE, help="产品编码/品号，默认使用脚本顶部 PRODUCT_CODE")
    parser.add_argument("--limit", type=int, default=80, help="每段最多显示行数")
    parser.add_argument(
        "--search-values",
        default="600,袋,罐,内,包",
        help="额外搜索值，逗号分隔。用于定位内标包装字段，例如 600,袋,罐",
    )
    return parser


def main() -> int:
    args = build_arg_parser().parse_args()

    description_rows = find_columns_by_description()
    print_table("1. 字段说明里包含包装/箱/袋/罐的列", description_rows, limit=args.limit)

    name_rows = find_columns_by_name()
    print_table("2. 字段名疑似包装/单位/数量的列", name_rows, limit=args.limit)

    product_tables = find_product_tables()
    print_table("3. 含 PRD_NO 的货品相关候选表", product_tables, limit=args.limit)

    product_rows = fetch_product_rows_from_candidate_tables(args.product_code, product_tables)
    print_table(f"4. 产品 {args.product_code} 在候选表里的疑似包装字段值", product_rows, limit=args.limit)

    prdt_pack_values = fetch_prdt_pack_field_values(args.product_code)
    print_table(f"5. PRDT 包装/单位/数量相关字段完整值：{args.product_code}", prdt_pack_values, limit=args.limit)

    prdt_non_empty_values = fetch_prdt_all_non_empty_values(args.product_code)
    print_table(f"6. PRDT 非空字段值：{args.product_code}", prdt_non_empty_values, limit=args.limit)

    detail_values = fetch_pack_detail_table_values(args.product_code)
    print_table(f"7. PRDT/货品明细候选表非空字段值：{args.product_code}", detail_values, limit=args.limit)

    search_values = [item.strip() for item in args.search_values.split(",") if item.strip()]
    value_matches = fetch_rows_matching_values(args.product_code, search_values)
    print_table(
        f"8. 所有含 PRD_NO 表中匹配 {search_values} 的字段值：{args.product_code}",
        value_matches,
        limit=args.limit,
    )

    print()
    print("下一步：")
    print("请重点看第 5 段：PK2_QTY 很可能是包装（一）数量，也就是 OUTER_QTY。")
    print("这次结果里 PRDT.PK3_UT / PK3_QTY 是空的，所以请继续看第 7 段是否有包装（二）的单位和数量。")
    print("如果第 7 段还没有，请看第 8 段里 600、袋、罐 分别出现在哪张表和哪个字段。")
    print("确认后就可以把 erp_runtime_csv.py 从 BOM 取数改成从货品基础资料取 OUTER_QTY/INNER_QTY。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
