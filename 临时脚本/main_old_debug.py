from __future__ import annotations

import subprocess

from bartender_label_runner import BarTenderRunnerError, generate_final_label_job
from bartender_preview_runner import generate_preview_image
from config import BARTENDER_EXE, ERP_DATABASE, ERP_DRIVER, ERP_PASSWORD, ERP_SERVER, ERP_USER, require_erp_config
from erp_runtime_csv import (
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


TEST_MO_NO = "MO26061617"
TEST_CSV_OUTPUT = "runtime_data/current_label.csv"
LABEL_CSV_FIELDS = [
    "MO_NO",
    "CUS_NO",
    "SUP_PRD_NO",
    "MRP_NO",
    "MFG_DATE",
    "EXP_DATE",
    "HAS_INNER_LABEL",
    "QTY",
    "LABEL_COUNT",
    "PACKAGE_PRD_NO",
    "PACKAGE_NAME",
]
TEST_RUNTIME_OPTIONS = {
    # Override ERP-calculated fields here if needed.
    # Example: {"EXP_DATE": "2027-06-18"}
}
PREVIEW_BARTENDER_EXE = BARTENDER_EXE
PRINTER_NAME = ""
PRINT_NOW = True
RUN_BARTENDER = False
GENERATE_PREVIEW = True
TEMPLATE_PATH_COLUMNS = ("模板相对路径", "模版相对路径", "模板路径", "模版路径", "相对路径")
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


def build_runtime_rows_for_mo(mo_no: str, runtime_options=None):
    db_rows = fetch_mo_rows(mo_no)
    if not db_rows:
        raise RuntimeCsvError(f"未查询到 MO 数据：{mo_no}")
    return [
        build_runtime_csv_row(normalize_db_row(row), runtime_options=runtime_options)
        for row in db_rows
    ]


def build_label_csv_row(runtime_row, qty_key: str, count_key: str, package_prd_key: str, package_name_key: str):
    return {
        "MO_NO": runtime_row.get("MO_NO", ""),
        "CUS_NO": runtime_row.get("CUS_NO", ""),
        "SUP_PRD_NO": runtime_row.get("SUP_PRD_NO", ""),
        "MRP_NO": runtime_row.get("MRP_NO", ""),
        "MFG_DATE": runtime_row.get("MFG_DATE", ""),
        "EXP_DATE": runtime_row.get("EXP_DATE", ""),
        "HAS_INNER_LABEL": runtime_row.get("HAS_INNER_LABEL", ""),
        "QTY": runtime_row.get(qty_key, ""),
        "LABEL_COUNT": runtime_row.get(count_key, ""),
        "PACKAGE_PRD_NO": runtime_row.get(package_prd_key, ""),
        "PACKAGE_NAME": runtime_row.get(package_name_key, ""),
    }


def write_label_csv(runtime_rows, label_type: str):
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

    label_rows = [
        build_label_csv_row(
            row,
            qty_key=qty_key,
            count_key=count_key,
            package_prd_key=package_prd_key,
            package_name_key=package_name_key,
        )
        for row in runtime_rows
    ]
    return write_runtime_csv(label_rows, TEST_CSV_OUTPUT, fields=LABEL_CSV_FIELDS)


def find_template_path_by_label_types(csv_path, label_types) -> str:
    last_error = None
    for label_type in label_types:
        try:
            template_mapping = find_template_mapping_by_runtime_csv(csv_path, label_type=label_type)
        except TemplateLookupError as error:
            last_error = error
            continue

        template_path = get_column(template_mapping, TEMPLATE_PATH_COLUMNS)
        if template_path:
            return template_path

    if last_error:
        raise last_error
    raise TemplateLookupError(f"未找到标签类型对应的模板：{', '.join(label_types)}")


def find_optional_template_path_by_label_types(csv_path, label_types) -> str:
    try:
        return find_template_path_by_label_types(csv_path, label_types)
    except TemplateLookupError:
        return ""


def set_inner_label_enabled(runtime_rows, enabled: bool) -> None:
    value = "Y" if enabled else "N"
    for row in runtime_rows:
        row["HAS_INNER_LABEL"] = value
        if not enabled:
            row["INNER_LABEL_COUNT"] = "0"


def generate_label_job(label_name: str, template_path: str, csv_path) -> None:
    if GENERATE_PREVIEW:
        preview_paths = generate_preview_image(
            template_path=template_path,
            csv_path=csv_path,
            label_name=label_name,
            bartender_exe=PREVIEW_BARTENDER_EXE,
            run=True,
        )
        for preview_path in preview_paths:
            print(f"{label_name}预览图：{preview_path}")

    manifest_path = generate_final_label_job(
        template_path=template_path,
        csv_path=csv_path,
        bartender_exe=BARTENDER_EXE,
        printer_name=PRINTER_NAME,
        print_now=PRINT_NOW,
        run=RUN_BARTENDER,
    )
    print(f"{label_name}打印任务：{manifest_path}")


def main() -> None:
    if TEST_MO_NO:
        try:
            runtime_rows = build_runtime_rows_for_mo(TEST_MO_NO, runtime_options=TEST_RUNTIME_OPTIONS)
            csv_path = write_label_csv(runtime_rows, "outer")
        except RuntimeCsvError as error:
            print("CSV生成失败")
            print(error)
            return

        print("CSV生成成功")
        print(f"MO号：{TEST_MO_NO}")
        print(f"CSV路径：{csv_path}")

        try:
            outer_template_path = find_template_path_by_label_types(csv_path, ("外标", "内外标"))
            inner_template_path = find_optional_template_path_by_label_types(csv_path, ("内标", "内外标"))
            has_inner_label = bool(inner_template_path)
            set_inner_label_enabled(runtime_rows, has_inner_label)
            inner_template_path = inner_template_path or "无"
        except TemplateLookupError as error:
            print("模板查找失败")
            print(error)
            return

        print("模板查找成功")
        shelf_life_months = shelf_life_months_from_package_name(runtime_rows[0].get("INNER_PACKAGE_NAME", ""))
        print(f"保质期：{shelf_life_months}个月" if shelf_life_months else "保质期：-")
        print(f"外标模板：{outer_template_path}")
        print(f"内标模板：{inner_template_path}")

        try:
            csv_path = write_label_csv(runtime_rows, "outer")
            print(f"已写入外标CSV：{csv_path}")
            generate_label_job("外标", outer_template_path, csv_path)

            if has_inner_label:
                csv_path = write_label_csv(runtime_rows, "inner")
                print(f"已写入内标CSV：{csv_path}")
                generate_label_job("内标", inner_template_path, csv_path)
        except (RuntimeCsvError, BarTenderRunnerError, subprocess.CalledProcessError, FileNotFoundError) as error:
            print("BarTender打印任务生成失败")
            print(error)
            return

        if RUN_BARTENDER:
            print("BarTender打印命令已执行")
        else:
            print("当前为 dry-run，只生成打印任务文件；current_label.csv 会停留在最后一次写入的标签数据。")
            print("确认无误后将 RUN_BARTENDER 改为 True，可按外标 -> 内标顺序直接打印。")
        return

    # Set TEST_MO_NO to run the full label flow. Without it, show recent ERP rows.
    conn = connect_erp()
    print("连接成功")

    cursor = conn.cursor()
    cursor.execute("""
    SELECT TOP 10
        MO_NO,
        MO_DD,
        SO_NO,
        QTY,
        UNIT
    FROM MF_MO
    ORDER BY MO_DD DESC
    """)

    for row in cursor.fetchall():
        print(row)

    conn.close()


if __name__ == "__main__":
    main()
