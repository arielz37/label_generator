# Label Generator Backend

This project builds the backend flow for generating BarTender label data from ERP manufacturing orders.

Current scope:

1. Query ERP data by `MO_NO`.
2. Generate `runtime_data/current_label.csv`.
3. Read product label rules from `product_label_details.xlsx`.
4. Calculate expiration date, label quantities, and label counts.
5. Locate the required `.btw` template from `template_mapping.xlsx`.
6. Generate a BarTender job manifest for later printing.

## Requirements

Python 3.8 is the target runtime.

Install Python dependencies:

```bash
pip install -r requirements.txt
```

ERP access also requires Microsoft ODBC Driver 17 for SQL Server to be installed on the machine.

Set ERP connection environment variables before running ERP queries:

```bash
export ERP_SERVER="your_sql_server"
export ERP_DATABASE="DB_BD01"
export ERP_USER="your_user"
export ERP_PASSWORD="your_password"
```

Optional:

```bash
export ERP_DRIVER="ODBC Driver 17 for SQL Server"
```

## Main Files

- `test_erp.py`  
  Main test entry. Set `TEST_MO_NO` and run the backend flow.

- `erp_runtime_csv.py`  
  Queries ERP `MF_MO` data and generates runtime CSV.

- `product_label_details_lookup.py`  
  Reads `product_label_details.xlsx` for product-level rules:
  shelf life, inner/outer label mode, inner/outer quantity.

- `template_mapping_lookup.py`  
  Reads `template_mapping.xlsx` and locates the `.btw` template.

- `generate_template_mapping.py`  
  Scans `Templates/` and generates an initial template mapping workbook.

- `bartender_label_runner.py`  
  Builds a dry-run BarTender job manifest from a `.btw` template and runtime CSV.

## Runtime CSV Fields

The current runtime CSV fields are:

```text
MO_NO
CUS_NO
SUP_PRD_NO
MRP_NO
MFG_DATE
EXP_DATE
HAS_INNER_LABEL
INNER_QTY
OUTER_QTY
INNER_LABEL_COUNT
OUTER_LABEL_COUNT
```

`SUP_PRD_NO` is the customer part number.  
`MRP_NO` is the product number.

## Test Flow

Edit `test_erp.py`:

```python
TEST_MO_NO = "your_mo_no"
TEST_LABEL_TYPE = "内标"
```

Then run:

```bash
python test_erp.py
```

If `TEST_MO_NO` is empty, `test_erp.py` only queries the latest 10 rows from `MF_MO`.

If `TEST_MO_NO` is set, the flow is:

```text
MO_NO
-> ERP query
-> runtime CSV
-> product_label_details.xlsx
-> calculated runtime CSV
-> template_mapping.xlsx
-> template path
```

## Mapping Workbooks

### `product_label_details.xlsx`

Product-level label rules. Lookup key:

```text
客户编号 + 客户料号
```

Main columns:

```text
客户编号
客户料号
产品编码
保质期（月）
外标性质（系统标：1，模版标：2）
内标性质（无内标：0，系统标：1，模版标：2）
产品数量（内标）（pcs）
产品数量（外标）
```

### `template_mapping.xlsx`

Template file mapping. Lookup key:

```text
客户编号 + 匹配编号 + 标签类型
```

Main output:

```text
模版相对路径
```

## BarTender Notes

Python does not decide where fields appear on the label. BarTender templates must bind their objects to the runtime CSV field names.

Example:

```text
MFG_DATE -> production date text object
EXP_DATE -> expiration date text object
SUP_PRD_NO -> customer part number text object
```

Current BarTender runner is dry-run by default:

```bash
python bartender_label_runner.py --template "H环旭-（订单号码20开头）/内外 标59-359808-01 300.btw" --csv runtime_data/current_label.csv
```

It writes a job manifest under `final_labels/`.
