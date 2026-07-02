from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Tuple

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from config import ERP_DATABASE, require_erp_config
from erp_runtime_csv import connect_erp, format_csv_value


DEFAULT_TARGET_VALUE = "7226003957"
DEFAULT_SAMPLE_LIMIT = 8

STRING_TYPES = (
    "char",
    "nchar",
    "varchar",
    "nvarchar",
    "text",
    "ntext",
)

TABLE_KEYWORDS = (
    "MO",
    "SO",
    "OS",
    "ORD",
    "ORDER",
    "PO",
    "CUS",
)

COLUMN_KEYWORDS = (
    "CUST_PO",
    "CUS_PO",
    "CUST_ORD",
    "CUS_ORD",
    "ORDER",
    "ORD",
    "PO",
    "OS",
    "CUS",
    "SO",
    "NO",
)

STRONG_COLUMN_KEYWORDS = (
    "PO",
    "ORDER",
    "ORD",
    "CUST",
    "CUS",
)


def quote_name(name: str) -> str:
    return "[" + name.replace("]", "]]") + "]"


def fetch_candidate_columns() -> List[Dict[str, str]]:
    conn = connect_erp()
    try:
        cursor = conn.cursor()
        placeholders = ", ".join("?" for _ in STRING_TYPES)
        table_conditions = " OR ".join("UPPER(TABLE_NAME) LIKE ?" for _ in TABLE_KEYWORDS)
        column_conditions = " OR ".join("UPPER(COLUMN_NAME) LIKE ?" for _ in COLUMN_KEYWORDS)
        sql = f"""
SELECT
    TABLE_SCHEMA,
    TABLE_NAME,
    COLUMN_NAME,
    DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE DATA_TYPE IN ({placeholders})
  AND (({table_conditions}) OR ({column_conditions}))
ORDER BY TABLE_NAME, COLUMN_NAME
"""
        params: List[Any] = list(STRING_TYPES)
        params.extend(f"%{keyword}%" for keyword in TABLE_KEYWORDS)
        params.extend(f"%{keyword}%" for keyword in COLUMN_KEYWORDS)
        cursor.execute(sql, params)
        columns = [column[0] for column in cursor.description]
        return [dict(zip(columns, row)) for row in cursor.fetchall()]
    finally:
        conn.close()


def score_column_name(table_name: str, column_name: str) -> Tuple[int, str]:
    text = f"{table_name}.{column_name}".upper()
    score = 0
    reasons = []
    for keyword in STRONG_COLUMN_KEYWORDS:
        if keyword in text:
            score += 8
            reasons.append(f"字段名含 {keyword}")
    if "SO_NO" in text:
        score -= 8
        reasons.append("更像 ERP SO 号")
    if "MO_NO" in text:
        score -= 8
        reasons.append("更像 MO 号")
    if column_name.upper().endswith("_NO") or column_name.upper() == "NO":
        score += 2
        reasons.append("编号字段")
    return score, "；".join(reasons)


def score_sample_value(value: str) -> Tuple[int, str]:
    text = (value or "").strip()
    score = 0
    reasons = []
    if not text:
        return 0, ""

    compact = re.sub(r"[\s\-_/]", "", text)
    if re.fullmatch(r"\d{8,13}", compact):
        score += 20
        reasons.append("8-13位数字")
    elif re.fullmatch(r"[A-Za-z0-9]{8,20}", compact) and re.search(r"\d", compact):
        score += 8
        reasons.append("像订单号的字母数字组合")

    upper = text.upper()
    if upper.startswith("SO"):
        score -= 12
        reasons.append("带 SO 前缀")
    if upper.startswith("MO"):
        score -= 12
        reasons.append("带 MO 前缀")
    if upper.startswith("PO"):
        score += 8
        reasons.append("带 PO 前缀")
    return score, "；".join(reasons)


def fetch_column_samples(schema: str, table: str, column: str, sample_limit: int) -> List[str]:
    conn = connect_erp()
    try:
        cursor = conn.cursor()
        sql = (
            f"SELECT DISTINCT TOP {int(sample_limit)} "
            f"LTRIM(RTRIM(CAST({quote_name(column)} AS NVARCHAR(4000)))) AS SAMPLE_VALUE "
            f"FROM {quote_name(schema)}.{quote_name(table)} "
            f"WHERE {quote_name(column)} IS NOT NULL "
            f"  AND LTRIM(RTRIM(CAST({quote_name(column)} AS NVARCHAR(4000)))) <> '' "
            f"  AND CAST({quote_name(column)} AS NVARCHAR(4000)) LIKE '%[0-9]%' "
            f"ORDER BY SAMPLE_VALUE DESC"
        )
        cursor.execute(sql)
        return [format_csv_value(row[0]) for row in cursor.fetchall()]
    finally:
        conn.close()


def profile_order_number_fields(max_columns: int = 800, sample_limit: int = DEFAULT_SAMPLE_LIMIT) -> List[Dict[str, str]]:
    candidates = fetch_candidate_columns()[:max_columns]
    results: List[Dict[str, str]] = []

    for candidate in candidates:
        schema = candidate["TABLE_SCHEMA"]
        table = candidate["TABLE_NAME"]
        column = candidate["COLUMN_NAME"]
        try:
            samples = fetch_column_samples(schema, table, column, sample_limit)
        except Exception:
            continue
        if not samples:
            continue

        column_score, column_reason = score_column_name(table, column)
        sample_scores = [score_sample_value(value) for value in samples]
        best_sample_score = max((item[0] for item in sample_scores), default=0)
        sample_reasons = [item[1] for item in sample_scores if item[1]]
        total_score = column_score + best_sample_score
        if total_score <= 0:
            continue

        results.append({
            "score": str(total_score),
            "table": f"{schema}.{table}",
            "column": column,
            "samples": " / ".join(samples[:sample_limit]),
            "reason": "；".join(dict.fromkeys([column_reason] + sample_reasons)),
        })

    return sorted(results, key=lambda item: int(item["score"]), reverse=True)


def search_exact_value(target_value: str, max_columns: int = 5000) -> List[Dict[str, str]]:
    candidates = fetch_candidate_columns()[:max_columns]
    matches: List[Dict[str, str]] = []

    conn = connect_erp()
    try:
        cursor = conn.cursor()
        for candidate in candidates:
            schema = candidate["TABLE_SCHEMA"]
            table = candidate["TABLE_NAME"]
            column = candidate["COLUMN_NAME"]
            sql = (
                f"SELECT TOP 3 * FROM {quote_name(schema)}.{quote_name(table)} "
                f"WHERE LTRIM(RTRIM(CAST({quote_name(column)} AS NVARCHAR(4000)))) = ?"
            )
            try:
                cursor.execute(sql, target_value)
                rows = cursor.fetchall()
            except Exception:
                continue
            if not rows:
                continue

            result_columns = [item[0] for item in cursor.description]
            for row in rows:
                row_dict = dict(zip(result_columns, row))
                preview = {
                    key: format_csv_value(value)
                    for key, value in row_dict.items()
                    if format_csv_value(value)
                }
                matches.append({
                    "table": f"{schema}.{table}",
                    "column": column,
                    "preview": str(preview)[:1000],
                })
    finally:
        conn.close()

    return matches


def print_candidate_columns(limit: int = 120) -> None:
    columns = fetch_candidate_columns()
    print(f"数据库：{ERP_DATABASE}")
    print(f"疑似字段数量：{len(columns)}")
    for item in columns[:limit]:
        print(f"{item['TABLE_SCHEMA']}.{item['TABLE_NAME']}.{item['COLUMN_NAME']} ({item['DATA_TYPE']})")
    if len(columns) > limit:
        print(f"... 仅显示前 {limit} 个")


def main() -> int:
    parser = argparse.ArgumentParser(description="查找 ERP 里的客户订单号字段")
    parser.add_argument("--value", default=DEFAULT_TARGET_VALUE, help="要搜索的客户订单号，例如 7226003957")
    parser.add_argument("--exact", action="store_true", help="精确搜索 --value 指定的订单号")
    parser.add_argument("--list-columns", action="store_true", help="只列出疑似字段，不抽样")
    parser.add_argument("--max-columns", type=int, default=800, help="最多抽样多少个疑似字段")
    parser.add_argument("--sample-limit", type=int, default=DEFAULT_SAMPLE_LIMIT, help="每个字段最多显示几个样本值")
    args = parser.parse_args()

    require_erp_config()
    if args.list_columns:
        print_candidate_columns()
        return 0

    if not args.exact:
        print("扫描疑似客户订单号字段")
        print("判断标准：字段名像 PO/ORDER/ORD/CUS，样本值像 8-13 位数字，且不像 SO/MO 号。")
        results = profile_order_number_fields(max_columns=args.max_columns, sample_limit=args.sample_limit)
        if not results:
            print("未找到明显疑似字段。可以加大 --max-columns，或运行 --list-columns 查看候选字段。")
            return 1
        for index, item in enumerate(results[:80], start=1):
            print(f"\n[{index}] score={item['score']} {item['table']}.{item['column']}")
            print(f"样本：{item['samples']}")
            print(f"原因：{item['reason']}")
        if len(results) > 80:
            print(f"\n... 仅显示前 80 个，共 {len(results)} 个疑似字段")
        return 0

    print(f"搜索客户订单号：{args.value}")
    matches = search_exact_value(args.value)
    if not matches:
        print("未找到精确匹配。可以先运行 --list-columns 看疑似字段，再缩小范围。")
        return 1

    for index, match in enumerate(matches, start=1):
        print(f"\n[{index}] {match['table']}.{match['column']}")
        print(match["preview"])
    print(f"\n共找到 {len(matches)} 处匹配")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
