from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any, Dict, Sequence


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from erp_runtime_csv import connect_erp, format_csv_value


# 在这里填你要验证的公司成品编码。
PRODUCT_CODE = "TDLAS0010_SHHXDZ"

# 内标数量来源料件：按中文名称检索。
# 规则：有“铁桶”或“罐”优先；否则从指定袋类名称里选择 QTY_BAS 最大的。
INNER_TANK_KEYWORD = "铁桶"
INNER_CAN_KEYWORD = "罐"
INNER_BAG_NAMES = (
    "公司自用透明PE袋",
    "双层防潮袋",
    "铝箔袋",
    "自用粉色pe袋",
    "公司自用粉红色静电袋",
)

# 外标数量来源料件：优先用中文名称匹配，避免每个纸箱料号都手工维护。
# 不建议只填“箱”，容易匹配到“周转箱”；先用“纸箱”更稳。
OUTER_NAME_KEYWORDS = (
    "纸箱",
)

MAX_OUTER_KEYWORDS = 5

# 默认 SQL：按成品料号查包装料件在 BOM 明细中的基数。
# 这里暂时把 TF_BOM.QTY_BAS 当作“包装内产品数量”候选值来验证。
SQL = """
WITH BomRows AS (
    SELECT
        m.BOM_NO,
        m.PRD_NO,
        m.NAME AS BOM_NAME,
        m.SPC,
        m.CUS_NO,
        m.QTY AS BOM_QTY,
        m.UNIT AS BOM_UNIT,
        t.ITM,
        t.PRD_NO AS COMPONENT_PRD_NO,
        t.NAME AS COMPONENT_NAME,
        t.UNIT AS COMPONENT_UNIT,
        t.QTY AS COMPONENT_QTY,
        t.QTY_BAS,
        t.REM
    FROM MF_BOM m
    INNER JOIN TF_BOM t
        ON m.BOM_NO = t.BOM_NO
    WHERE m.PRD_NO = ?
),
InnerCandidates AS (
    SELECT
        *,
        ROW_NUMBER() OVER (
            ORDER BY
                CASE
                    WHEN ISNULL(COMPONENT_NAME, '') LIKE '%铁桶%' THEN 0
                    WHEN ISNULL(COMPONENT_NAME, '') LIKE '%罐%' THEN 0
                    ELSE 1
                END,
                QTY_BAS DESC,
                ITM
        ) AS RN
    FROM BomRows
    WHERE ISNULL(COMPONENT_NAME, '') LIKE '%铁桶%'
       OR ISNULL(COMPONENT_NAME, '') LIKE '%罐%'
       OR LTRIM(RTRIM(ISNULL(COMPONENT_NAME, ''))) IN (
            '公司自用透明PE袋',
            '双层防潮袋',
            '铝箔袋',
            '自用粉色pe袋',
            '公司自用粉红色静电袋'
       )
),
OuterCandidates AS (
    SELECT
        *,
        ROW_NUMBER() OVER (
            ORDER BY
                QTY_BAS DESC,
                ITM
        ) AS RN
    FROM BomRows
    WHERE (
        COMPONENT_NAME LIKE ?
        OR COMPONENT_NAME LIKE ?
        OR COMPONENT_NAME LIKE ?
        OR COMPONENT_NAME LIKE ?
        OR COMPONENT_NAME LIKE ?
    )
    AND ISNULL(COMPONENT_NAME, '') NOT LIKE '%周转箱%'
)
SELECT
    '内标'              AS 对应标签,
    BOM_NO             AS BOM编号,
    PRD_NO             AS 成品料号,
    BOM_NAME           AS BOM名称,
    SPC                AS 成品规格,
    CUS_NO             AS 客户编号,
    BOM_QTY            AS 成品基准数量,
    BOM_UNIT           AS 成品单位,
    ITM                AS 明细行号,
    COMPONENT_PRD_NO   AS 子件料号,
    COMPONENT_NAME     AS 子件名称,
    COMPONENT_UNIT     AS 子件单位,
    COMPONENT_QTY      AS 子件用量,
    QTY_BAS            AS 用量基数,
    QTY_BAS            AS 建议包装数量,
    REM                AS 明细备注
FROM InnerCandidates
WHERE RN = 1

UNION ALL

SELECT
    '外标'              AS 对应标签,
    BOM_NO             AS BOM编号,
    PRD_NO             AS 成品料号,
    BOM_NAME           AS BOM名称,
    SPC                AS 成品规格,
    CUS_NO             AS 客户编号,
    BOM_QTY            AS 成品基准数量,
    BOM_UNIT           AS 成品单位,
    ITM                AS 明细行号,
    COMPONENT_PRD_NO   AS 子件料号,
    COMPONENT_NAME     AS 子件名称,
    COMPONENT_UNIT     AS 子件单位,
    COMPONENT_QTY      AS 子件用量,
    QTY_BAS            AS 用量基数,
    QTY_BAS            AS 建议包装数量,
    REM                AS 明细备注
FROM OuterCandidates
WHERE RN = 1

ORDER BY 对应标签, 明细行号
"""

# 如果 SQL 里有 ? 参数，在这里按顺序填写。
PADDED_OUTER_NAME_PATTERNS = tuple(f"%{keyword}%" for keyword in OUTER_NAME_KEYWORDS[:MAX_OUTER_KEYWORDS]) + ("",) * max(
    MAX_OUTER_KEYWORDS - len(OUTER_NAME_KEYWORDS),
    0,
)
PARAMS: Sequence[Any] = (
    PRODUCT_CODE,
    *PADDED_OUTER_NAME_PATTERNS,
)


def run_query(sql: str, params: Sequence[Any]) -> list[Dict[str, str]]:
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
    width = 0
    for char in value:
        width += 2 if ord(char) > 127 else 1
    return width


def pad_text(value: str, width: int) -> str:
    return value + " " * max(width - text_width(value), 0)


def print_table(rows: list[Dict[str, str]], limit: int = 100) -> None:
    if not rows:
        print("无查询结果")
        return

    headers = list(rows[0].keys())
    display_rows = rows[:limit]
    widths = []
    for header in headers:
        max_width = text_width(header)
        for row in display_rows:
            max_width = max(max_width, text_width(row.get(header, "")))
        widths.append(min(max_width, 60))

    header_line = " | ".join(pad_text(header, widths[index]) for index, header in enumerate(headers))
    separator = "-+-".join("-" * width for width in widths)
    print(header_line)
    print(separator)

    for row in display_rows:
        print(" | ".join(pad_text(row.get(header, ""), widths[index]) for index, header in enumerate(headers)))

    if len(rows) > limit:
        print(f"... 仅显示前 {limit} 行，共 {len(rows)} 行")
    else:
        print(f"共 {len(rows)} 行")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="连接 ERP 并执行脚本顶部 SQL，结果直接显示在终端。")
    parser.add_argument("--limit", type=int, default=100, help="终端最多显示多少行，默认 100")
    return parser


def main() -> int:
    args = build_arg_parser().parse_args()
    rows = run_query(SQL, PARAMS)
    print_table(rows, limit=max(args.limit, 1))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
