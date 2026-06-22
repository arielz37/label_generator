from __future__ import annotations

from erp_runtime_csv import RuntimeCsvError, generate_runtime_csv_for_mo
from product_label_details_lookup import (
    ProductLabelDetailsError,
    build_runtime_options_from_detail,
    find_product_label_detail_by_runtime_csv,
)
from template_mapping_lookup import (
    TemplateLookupError,
    find_template_mapping_by_runtime_csv,
    get_column,
)
from config import ERP_DATABASE, ERP_DRIVER, ERP_PASSWORD, ERP_SERVER, ERP_USER, require_erp_config


TEST_MO_NO = ""
TEST_LABEL_TYPE = "内标"
TEST_CSV_OUTPUT = "runtime_data/current_label.csv"
TEST_CSV_FIELDS = [
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
TEST_RUNTIME_OPTIONS = {
    # 可在这里覆盖 product_label_details.xlsx 算出的字段。
    # 例如：{"EXP_DATE": "2027-06-18"}
}


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


def main() -> None:
    if TEST_MO_NO:
        try:
            csv_path = generate_runtime_csv_for_mo(
                mo_no=TEST_MO_NO,
                output_path=TEST_CSV_OUTPUT,
                fields=TEST_CSV_FIELDS,
                runtime_options=TEST_RUNTIME_OPTIONS,
            )
        except RuntimeCsvError as error:
            print("CSV生成失败")
            print(error)
            return

        print("CSV生成成功")
        print(f"MO号：{TEST_MO_NO}")
        print(f"标签类型：{TEST_LABEL_TYPE or '-'}")
        print(f"CSV路径：{csv_path}")

        product_detail = {}
        try:
            product_detail = find_product_label_detail_by_runtime_csv(csv_path)
        except ProductLabelDetailsError as error:
            print("产品标签细节查找失败")
            print(error)
            return

        if product_detail:
            runtime_options = {
                **build_runtime_options_from_detail(product_detail),
                **TEST_RUNTIME_OPTIONS,
            }
            try:
                csv_path = generate_runtime_csv_for_mo(
                    mo_no=TEST_MO_NO,
                    output_path=TEST_CSV_OUTPUT,
                    fields=TEST_CSV_FIELDS,
                    runtime_options=runtime_options,
                )
            except RuntimeCsvError as error:
                print("CSV生成失败")
                print(error)
                return

        try:
            template_mapping = find_template_mapping_by_runtime_csv(csv_path, label_type=TEST_LABEL_TYPE)
        except TemplateLookupError as error:
            print("模板查找失败")
            print(error)
            return

        template_path = get_column(
            template_mapping,
            ("模板相对路径", "模版相对路径", "模板路径", "模版路径", "相对路径"),
        )
        print("模板查找成功")
        print(f"保质期：{get_column(product_detail, ('保质期（月）', '保质期', 'SHELF_LIFE')) or '-'}")
        print(f"模板相对路径：{template_path}")
        return

    # 设置 TEST_MO_NO 后：
    # 1. 从 ERP 生成 MO runtime CSV
    # 2. 从 runtime CSV 读取客户编号、客户料号、产品编号
    # 3. 使用 TEST_LABEL_TYPE 到 template_mapping.xlsx 查找模板相对路径
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
