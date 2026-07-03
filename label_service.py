from __future__ import annotations

import subprocess
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_FLOOR
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from bartender_label_runner import BarTenderRunnerError, OUTPUT_DIR as JOB_OUTPUT_DIR, generate_final_label_job
from bartender_preview_runner import (
    OUTPUT_DIR as PREVIEW_OUTPUT_DIR,
    cleanup_preview_cache,
    clear_directory,
    generate_preview_image,
)
from app_paths import RUNTIME_DIR
from config import BARTENDER_EXE, BARTENDER_PRINTER
from erp_runtime_csv import (
    DEFAULT_SHELF_LIFE_MONTHS,
    RuntimeCsvError,
    build_runtime_csv_row,
    fetch_mo_rows,
    normalize_db_row,
    shelf_life_months_from_package_name,
    write_runtime_csv,
)
from template_mapping_lookup import (
    TemplateLookupError,
    find_template_mapping_by_runtime_csv,
    get_column,
)


RUNTIME_CSV = RUNTIME_DIR / "current_label.csv"
PRINT_CSV_DIR = JOB_OUTPUT_DIR / "print_csv"
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
TEMPLATE_PATH_COLUMNS = ("模板相对路径", "模版相对路径", "模板路径", "模版路径", "相对路径")
INNER_LABEL_TRUE_VALUES = {"y", "yes", "1", "true", "是", "有"}
TEMPLATE_MISSING_HINT = "当前可能应该使用系统标，或 template_mapping.xlsx 还未录入对应模板标。"


class LabelServiceError(Exception):
    pass


def safe_segment(value: str) -> str:
    chars = []
    for char in value:
        if char.isalnum() or char in ("-", "_"):
            chars.append(char)
        else:
            chars.append("_")
    return "".join(chars).strip("_") or "label"


def build_runtime_rows_for_mo(mo_no: str, runtime_options: Optional[Dict[str, str]] = None) -> List[Dict[str, str]]:
    db_rows = fetch_mo_rows(mo_no)
    if not db_rows:
        raise RuntimeCsvError(f"未查询到 MO 数据：{mo_no}")
    return [
        build_runtime_csv_row(normalize_db_row(row), runtime_options=runtime_options)
        for row in db_rows
    ]


def parse_label_count(value: str) -> int:
    value = (value or "").strip()
    if not value:
        return 1
    try:
        count = Decimal(value)
        return max(int(count.to_integral_value(rounding=ROUND_FLOOR)), 0)
    except (InvalidOperation, ValueError):
        raise LabelServiceError(f"无法识别 LABEL_COUNT：{value}")


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


def label_csv_keys(label_type: str) -> Tuple[str, str, str, str]:
    if label_type == "outer":
        return "OUTER_QTY", "OUTER_LABEL_COUNT", "OUTER_PACKAGE_PRD_NO", "OUTER_PACKAGE_NAME"
    elif label_type == "inner":
        return "INNER_QTY", "INNER_LABEL_COUNT", "INNER_PACKAGE_PRD_NO", "INNER_PACKAGE_NAME"
    raise RuntimeCsvError(f"未知标签 CSV 类型：{label_type}")


def label_remainder_key(label_type: str) -> str:
    if label_type == "outer":
        return "OUTER_REMAINDER_QTY"
    elif label_type == "inner":
        return "INNER_REMAINDER_QTY"
    raise RuntimeCsvError(f"未知标签 CSV 类型：{label_type}")


def build_label_csv_rows(runtime_rows: List[Dict[str, str]], label_type: str) -> List[Dict[str, str]]:
    qty_key, count_key, package_prd_key, package_name_key = label_csv_keys(label_type)
    remainder_key = label_remainder_key(label_type)
    rows = []
    for row in runtime_rows:
        label_count = parse_label_count(row.get(count_key, "1"))
        remainder_qty = row.get(remainder_key, "")
        total_label_count = label_count + (1 if remainder_qty else 0)
        for label_index in range(1, label_count + 1):
            label_row = build_label_csv_row(
                row,
                qty_key=qty_key,
                count_key=count_key,
                package_prd_key=package_prd_key,
                package_name_key=package_name_key,
                label_index=label_index,
            )
            label_row["LABEL_COUNT"] = str(total_label_count)
            rows.append(label_row)

        if remainder_qty:
            label_row = build_label_csv_row(
                row,
                qty_key=qty_key,
                count_key=count_key,
                package_prd_key=package_prd_key,
                package_name_key=package_name_key,
                label_index=label_count + 1,
            )
            label_row["QTY"] = remainder_qty
            label_row["LABEL_COUNT"] = str(total_label_count)
            rows.append(label_row)

    return rows


def has_label_rows(runtime_rows: List[Dict[str, str]], label_type: str) -> bool:
    return bool(build_label_csv_rows(runtime_rows, label_type))


def write_lookup_csv(runtime_rows: List[Dict[str, str]]) -> Path:
    if not runtime_rows:
        raise RuntimeCsvError("没有可写入 CSV 的数据。")
    row = build_label_csv_row(
        runtime_rows[0],
        qty_key="OUTER_QTY",
        count_key="OUTER_LABEL_COUNT",
        package_prd_key="OUTER_PACKAGE_PRD_NO",
        package_name_key="OUTER_PACKAGE_NAME",
        label_index=1,
    )
    return write_runtime_csv([row], RUNTIME_CSV, fields=LABEL_CSV_FIELDS)


def write_label_csv(
    runtime_rows: List[Dict[str, str]],
    label_type: str,
    output_path: Optional[Path] = None,
    row_limit: Optional[int] = None,
) -> Path:
    rows = build_label_csv_rows(runtime_rows, label_type)
    if row_limit is not None:
        if row_limit < 1 or row_limit > len(rows):
            label_name = "外标" if label_type == "outer" else "内标"
            raise LabelServiceError(f"{label_name}打印张数必须在 1 到 {len(rows)} 之间")
        rows = rows[:row_limit]
    return write_runtime_csv(rows, output_path or RUNTIME_CSV, fields=LABEL_CSV_FIELDS)


def write_print_csv_snapshot(
    runtime_rows: List[Dict[str, str]],
    label_type: str,
    mo_no: str,
    row_limit: Optional[int] = None,
) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_dir = PRINT_CSV_DIR / safe_segment(mo_no)
    csv_path = csv_dir / f"{timestamp}_{safe_segment(label_type)}.csv"
    suffix = 1
    while csv_path.exists():
        csv_path = csv_dir / f"{timestamp}_{safe_segment(label_type)}_{suffix}.csv"
        suffix += 1
    return write_label_csv(runtime_rows, label_type, output_path=csv_path, row_limit=row_limit)


def find_template_path_by_label_types(csv_path: Path, label_types: Iterable[str]) -> str:
    last_error = None
    for label_type in label_types:
        try:
            mapping = find_template_mapping_by_runtime_csv(csv_path, label_type=label_type)
        except TemplateLookupError as error:
            last_error = error
            continue

        template_path = get_column(mapping, TEMPLATE_PATH_COLUMNS)
        if template_path:
            return template_path

    if last_error:
        raise last_error
    raise TemplateLookupError("未找到标签类型对应的模板")


def find_optional_template_path_by_label_types(csv_path: Path, label_types: Iterable[str]) -> str:
    try:
        return find_template_path_by_label_types(csv_path, label_types)
    except TemplateLookupError:
        return ""


def normalize_template_override(template_path: str) -> str:
    value = (template_path or "").strip()
    if not value:
        return ""
    return value


def find_template_for_mo(
    mo_no: str,
    label_type: str,
    outer_template_override: str = "",
    inner_template_override: str = "",
) -> str:
    prepared = prepare_runtime(mo_no)
    runtime_rows: List[Dict[str, str]] = prepared["runtime_rows"]  # type: ignore[assignment]
    csv_path: Path = prepared["csv_path"]  # type: ignore[assignment]
    runtime_row = runtime_rows[0]

    if label_type == "outer":
        override = normalize_template_override(outer_template_override)
        if override:
            return override
        try:
            return find_template_path_by_label_types(csv_path, ("外标", "内外标"))
        except TemplateLookupError as error:
            raise template_missing_error(runtime_row, ("外标", "内外标"), error) from error

    if label_type == "inner":
        override = normalize_template_override(inner_template_override)
        if override:
            return override
        try:
            return find_template_path_by_label_types(csv_path, ("内标", "内外标"))
        except TemplateLookupError as error:
            raise template_missing_error(runtime_row, ("内标", "内外标"), error) from error

    raise LabelServiceError(f"未知模板类型：{label_type}")


def template_lookup_context(runtime_row: Dict[str, str]) -> str:
    return (
        f"客户编号={runtime_row.get('CUS_NO', '-') or '-'}，"
        f"客户料号={runtime_row.get('SUP_PRD_NO', '-') or '-'}，"
        f"产品编号={runtime_row.get('MRP_NO', '-') or '-'}"
    )


def template_missing_error(runtime_row: Dict[str, str], label_types: Iterable[str], error: Exception) -> LabelServiceError:
    label_text = "/".join(label_types)
    return LabelServiceError(
        f"未找到{label_text}模板标配置。\n"
        f"{template_lookup_context(runtime_row)}\n"
        f"{TEMPLATE_MISSING_HINT}\n"
        f"原始错误：{error}"
    )


def build_missing_inner_notice(runtime_row: Dict[str, str]) -> str:
    return (
        f"未找到内标/内外标模板标配置，当前按无内标处理。"
        f"{TEMPLATE_MISSING_HINT}"
    )


def build_missing_inner_qty_notice(runtime_row: Dict[str, str]) -> str:
    return (
        "检测不到内标包装数量：ERP/BOM 未匹配到内标包装用量基数。"
        "内标模板已显示，但暂不能生成内标预览或打印内标。"
    )


def set_inner_label_enabled(runtime_rows: List[Dict[str, str]], enabled: bool) -> None:
    value = "Y" if enabled else "N"
    for row in runtime_rows:
        row["HAS_INNER_LABEL"] = value
        if not enabled:
            row["INNER_LABEL_COUNT"] = "0"


def prepare_runtime(
    mo_no: str,
    shelf_life: str = "",
    inner_package_name: str = "",
    outer_package_name: str = "",
) -> Dict[str, object]:
    runtime_options = {}
    if shelf_life:
        runtime_options["SHELF_LIFE"] = shelf_life
    if inner_package_name:
        runtime_options["INNER_PACKAGE_NAME_OVERRIDE"] = inner_package_name
    if outer_package_name:
        runtime_options["OUTER_PACKAGE_NAME_OVERRIDE"] = outer_package_name
    runtime_rows = build_runtime_rows_for_mo(mo_no, runtime_options=runtime_options)
    csv_path = write_lookup_csv(runtime_rows)

    return {
        "runtime_rows": runtime_rows,
        "csv_path": csv_path,
    }


def build_common_result(mo_no: str, runtime_row: Dict[str, str]) -> Dict[str, object]:
    shelf_life_months = runtime_row.get("SHELF_LIFE_MONTHS") or shelf_life_months_from_package_name(
        runtime_row.get("INNER_PACKAGE_NAME", "")
    )
    remainder_notices = []
    if runtime_row.get("OUTER_REMAINDER_QTY"):
        outer_count = parse_label_count(runtime_row.get("OUTER_LABEL_COUNT", "0"))
        if outer_count == 0:
            remainder_notices.append(
                f"外标本次无整标，将生成 1 张零数箱标签，QTY={runtime_row.get('OUTER_REMAINDER_QTY')}。请人工额外打印零数箱标签。"
            )
        else:
            remainder_notices.append(
                f"外标数量不是整箱：整标 {outer_count} 张，另生成 1 张零数箱标签，QTY={runtime_row.get('OUTER_REMAINDER_QTY')}。请人工额外打印零数箱标签。"
            )
    if runtime_row.get("INNER_REMAINDER_QTY"):
        inner_count = parse_label_count(runtime_row.get("INNER_LABEL_COUNT", "0"))
        if inner_count == 0:
            remainder_notices.append(
                f"内标本次无整标，将生成 1 张零数袋标签，QTY={runtime_row.get('INNER_REMAINDER_QTY')}。"
            )
        else:
            remainder_notices.append(
                f"内标数量不是整袋：整标 {inner_count} 张，另生成 1 张零数袋标签，QTY={runtime_row.get('INNER_REMAINDER_QTY')}。"
            )

    return {
        "mo_no": mo_no,
        "customer_code": runtime_row.get("CUS_NO", ""),
        "customer_part_no": runtime_row.get("SUP_PRD_NO", ""),
        "product_code": runtime_row.get("MRP_NO", ""),
        "mfg_date": runtime_row.get("MFG_DATE", ""),
        "exp_date": runtime_row.get("EXP_DATE", ""),
        "has_inner_label": runtime_row.get("HAS_INNER_LABEL", "").strip().casefold() in INNER_LABEL_TRUE_VALUES,
        "shelf_life": f"{shelf_life_months}个月" if shelf_life_months else "",
        "notices": [],
        "remainder_notices": remainder_notices,
    }


def preview_urls_from_paths(paths: List[Path]) -> List[str]:
    urls = []
    root = JOB_OUTPUT_DIR.resolve()
    for path in paths:
        resolved = Path(path).resolve()
        try:
            relative = resolved.relative_to(root)
        except ValueError:
            relative = resolved.name
        urls.append("/files/" + str(relative).replace("\\", "/"))
    return urls


def generate_label_preview(
    mo_no: str,
    shelf_life: str = "",
    printer_name: str = BARTENDER_PRINTER,
    outer_template_override: str = "",
    inner_template_override: str = "",
    inner_package_name: str = "",
    outer_package_name: str = "",
) -> Dict[str, object]:
    prepared = prepare_runtime(
        mo_no,
        shelf_life=shelf_life,
        inner_package_name=inner_package_name,
        outer_package_name=outer_package_name,
    )
    runtime_rows: List[Dict[str, str]] = prepared["runtime_rows"]  # type: ignore[assignment]
    csv_path: Path = prepared["csv_path"]  # type: ignore[assignment]
    runtime_row = runtime_rows[0]

    outer_template = normalize_template_override(outer_template_override)
    if not outer_template:
        try:
            outer_template = find_template_path_by_label_types(csv_path, ("外标", "内外标"))
        except TemplateLookupError as error:
            raise template_missing_error(runtime_row, ("外标", "内外标"), error) from error

    has_inner_package = bool(runtime_row.get("INNER_QTY"))
    inner_missing_notice = ""
    inner_qty_missing_notice = ""
    inner_template = normalize_template_override(inner_template_override)
    if not inner_template:
        try:
            inner_template = find_template_path_by_label_types(csv_path, ("内标", "内外标"))
        except TemplateLookupError:
            inner_template = ""
            if has_inner_package:
                inner_missing_notice = build_missing_inner_notice(runtime_row)
    if inner_template and not has_inner_package:
        inner_qty_missing_notice = build_missing_inner_qty_notice(runtime_row)
    if not has_inner_package:
        runtime_row["INNER_LABEL_COUNT"] = "0"
    if not inner_template:
        inner_template = ""
    has_inner = bool(inner_template) and has_inner_package
    set_inner_label_enabled(runtime_rows, has_inner)

    result = build_common_result(mo_no, runtime_row)
    result["shelf_life_input"] = shelf_life
    result["shelf_life_source"] = "手动输入" if shelf_life else f"默认一年（{DEFAULT_SHELF_LIFE_MONTHS}个月）"
    result["printer_name"] = printer_name
    result["outer_template_override"] = outer_template_override
    result["inner_template_override"] = inner_template_override
    result["inner_package_override"] = inner_package_name
    result["outer_package_override"] = outer_package_name
    if inner_missing_notice:
        result["notices"].append(inner_missing_notice)  # type: ignore[union-attr]
    if inner_qty_missing_notice:
        result["notices"].append(inner_qty_missing_notice)  # type: ignore[union-attr]

    preview_root = PREVIEW_OUTPUT_DIR / safe_segment(mo_no)
    clear_directory(preview_root)

    outer_paths: List[Path] = []
    if has_label_rows(runtime_rows, "outer"):
        csv_path = write_label_csv(runtime_rows, "outer")
        outer_paths = generate_preview_image(
            template_path=outer_template,
            csv_path=csv_path,
            label_name="outer",
            output_dir=preview_root / "outer",
            bartender_exe=BARTENDER_EXE,
            printer_name=printer_name,
            run=True,
        )
    result["outer"] = {
        "template": outer_template,
        "qty": runtime_row.get("OUTER_QTY", ""),
        "label_count": runtime_row.get("OUTER_LABEL_COUNT", ""),
        "remainder_qty": runtime_row.get("OUTER_REMAINDER_QTY", ""),
        "zero_label_qty": runtime_row.get("OUTER_REMAINDER_QTY", ""),
        "total_label_count": str(len(build_label_csv_rows(runtime_rows, "outer"))),
        "package_prd_no": runtime_row.get("OUTER_PACKAGE_PRD_NO", ""),
        "package_name": runtime_row.get("OUTER_PACKAGE_NAME", ""),
        "preview_images": preview_urls_from_paths(outer_paths),
    }

    if inner_template:
        inner_paths: List[Path] = []
        if has_inner and has_label_rows(runtime_rows, "inner"):
            csv_path = write_label_csv(runtime_rows, "inner")
            inner_paths = generate_preview_image(
                template_path=inner_template,
                csv_path=csv_path,
                label_name="inner",
                output_dir=preview_root / "inner",
                bartender_exe=BARTENDER_EXE,
                printer_name=printer_name,
                run=True,
            )
        result["inner"] = {
            "template": inner_template,
            "qty": runtime_row.get("INNER_QTY", ""),
            "label_count": runtime_row.get("INNER_LABEL_COUNT", ""),
            "remainder_qty": runtime_row.get("INNER_REMAINDER_QTY", ""),
            "zero_label_qty": runtime_row.get("INNER_REMAINDER_QTY", ""),
            "total_label_count": str(len(build_label_csv_rows(runtime_rows, "inner"))),
            "package_prd_no": runtime_row.get("INNER_PACKAGE_PRD_NO", ""),
            "package_name": runtime_row.get("INNER_PACKAGE_NAME", ""),
            "preview_images": preview_urls_from_paths(inner_paths),
            "can_print": has_inner,
        }
    else:
        result["inner"] = None

    cleanup_preview_cache()
    return result


def print_labels(
    mo_no: str,
    label_types: Iterable[str],
    printer_name: str = BARTENDER_PRINTER,
    shelf_life: str = "",
    outer_template_override: str = "",
    inner_template_override: str = "",
    inner_package_name: str = "",
    outer_package_name: str = "",
    print_row_limits: Optional[Dict[str, int]] = None,
) -> Dict[str, object]:
    requested = [item for item in label_types if item in ("outer", "inner")]
    if not requested:
        raise LabelServiceError("请选择要打印的标签类型")

    prepared = prepare_runtime(
        mo_no,
        shelf_life=shelf_life,
        inner_package_name=inner_package_name,
        outer_package_name=outer_package_name,
    )
    runtime_rows: List[Dict[str, str]] = prepared["runtime_rows"]  # type: ignore[assignment]
    csv_path: Path = prepared["csv_path"]  # type: ignore[assignment]
    runtime_row = runtime_rows[0]

    outer_template = normalize_template_override(outer_template_override)
    if not outer_template:
        try:
            outer_template = find_template_path_by_label_types(csv_path, ("外标", "内外标"))
        except TemplateLookupError as error:
            raise template_missing_error(runtime_row, ("外标", "内外标"), error) from error

    if runtime_row.get("INNER_QTY"):
        inner_template = normalize_template_override(inner_template_override)
        if not inner_template:
            try:
                inner_template = find_template_path_by_label_types(csv_path, ("内标", "内外标"))
            except TemplateLookupError:
                inner_template = ""
    else:
        inner_template = ""
    has_inner = bool(inner_template) and bool(runtime_row.get("INNER_QTY"))
    set_inner_label_enabled(runtime_rows, has_inner)

    printed = []
    for label_type in requested:
        if label_type == "inner" and not has_inner:
            continue
        if not has_label_rows(runtime_rows, label_type):
            continue

        row_limit = (print_row_limits or {}).get(label_type)
        write_label_csv(runtime_rows, label_type, row_limit=row_limit)
        csv_path = write_print_csv_snapshot(runtime_rows, label_type, mo_no, row_limit=row_limit)
        template = outer_template if label_type == "outer" else inner_template
        manifest_path = generate_final_label_job(
            template_path=template,
            csv_path=csv_path,
            bartender_exe=BARTENDER_EXE,
            printer_name=printer_name,
            print_now=True,
            copies=1,
            run=True,
        )
        printed.append({
            "type": label_type,
            "manifest": str(manifest_path),
            "template": template,
            "row_limit": row_limit or "",
        })

    result = build_common_result(mo_no, runtime_row)
    result["printer_name"] = printer_name
    result["inner_package_override"] = inner_package_name
    result["outer_package_override"] = outer_package_name
    result["printed"] = printed
    return result


def service_error_message(error: Exception) -> str:
    if isinstance(error, subprocess.CalledProcessError):
        return f"BarTender 执行失败，退出码：{error.returncode}"
    return str(error)
