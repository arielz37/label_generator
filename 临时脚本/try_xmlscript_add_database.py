from __future__ import annotations

import argparse
import csv
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
DEFAULT_TEMPLATE = PROJECT_ROOT / "Templates" / "Amelco" / "BBHIC560-AD（HIC-0560-CF-BA3 ）  内标是系统表.btw"
DEFAULT_RUNTIME_CSV = PROJECT_ROOT / "runtime_data" / "current_label.csv"
TRIAL_ROOT = PROJECT_ROOT / "临时脚本" / "xmlscript_database_trials"

CSV_FIELDS = [
    "MO_NO",
    "CUS_NO",
    "SUP_PRD_NO",
    "MRP_NO",
    "MFG_DATE",
    "EXP_DATE",
    "HAS_INNER_LABEL",
    "QTY",
    "LABEL_COUNT",
    "LABEL_INDEX",
    "PACKAGE_PRD_NO",
    "PACKAGE_NAME",
]


class TrialError(Exception):
    pass


def find_bartender_exe() -> str:
    try:
        from config import BARTENDER_EXE
    except Exception:
        BARTENDER_EXE = "bartend.exe"
    if BARTENDER_EXE and Path(BARTENDER_EXE).exists():
        return BARTENDER_EXE
    for candidate in (
        Path(r"C:\Program Files\Seagull\BarTender Suite\bartend.exe"),
        Path(r"C:\Program Files (x86)\Seagull\BarTender Suite\bartend.exe"),
    ):
        if candidate.exists():
            return str(candidate)
    return BARTENDER_EXE


def ensure_runtime_csv(path: Path) -> None:
    if path.exists():
        return

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerow({field: "" for field in CSV_FIELDS})


def build_xmlscript(template_path: Path, csv_path: Path, recordset_name: str, variant: str) -> str:
    delimitation = "btDelimMixedQuoteAndComma" if variant != "comma" else "btDelimComma"
    add_if_none = "true" if variant != "no_add_flag" else "false"
    return f'''<?xml version="1.0" encoding="utf-8"?>
<XMLScript Version="2.0">
  <Command Name="AddCsvDatabase">
    <FormatSetup>
      <Format CloseAtEndOfJob="true" SaveAtEndOfJob="true">{escape(str(template_path))}</Format>
      <RecordSet Name="{escape(recordset_name)}" Type="btTextFile" AddIfNone="{add_if_none}">
        <FileName>{escape(str(csv_path))}</FileName>
        <Delimitation>{delimitation}</Delimitation>
        <UseFieldNamesFromFirstRecord>true</UseFieldNamesFromFirstRecord>
      </RecordSet>
    </FormatSetup>
  </Command>
</XMLScript>
'''


def database_count(template_path: Path) -> int:
    try:
        import win32com.client
    except ModuleNotFoundError as error:
        raise TrialError("缺少 pywin32，请先运行：pip install pywin32") from error

    app = win32com.client.Dispatch("BarTender.Application")
    format_doc = None
    try:
        app.Visible = False
        format_doc = app.Formats.Open(str(template_path), False, "")
        databases = format_doc.Databases
        try:
            return int(databases.Count)
        except Exception:
            return 0
    finally:
        if format_doc is not None:
            try:
                format_doc.Close(0)
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass


def run_trial(template_path: Path, csv_path: Path, variant: str, visible: bool) -> int:
    if not template_path.exists():
        raise TrialError(f"模板不存在：{template_path}")
    if not csv_path.exists():
        raise TrialError(f"runtime CSV 不存在：{csv_path}")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    trial_dir = TRIAL_ROOT / timestamp
    trial_dir.mkdir(parents=True, exist_ok=True)
    trial_template = trial_dir / template_path.name
    shutil.copy2(template_path, trial_template)

    script_path = trial_dir / f"add_database_{variant}.xml"
    script_path.write_text(
        build_xmlscript(
            template_path=trial_template,
            csv_path=csv_path,
            recordset_name="current_label",
            variant=variant,
        ),
        encoding="utf-8",
    )

    before_count = database_count(trial_template)
    bartender_exe = find_bartender_exe()
    command = [bartender_exe, f"/XMLSCRIPT={script_path}", "/X"]
    if visible:
        command.remove("/X")

    print(f"模板副本：{trial_template}")
    print(f"XMLScript：{script_path}")
    print(f"BarTender：{bartender_exe}")
    print(f"执行前数据库连接数：{before_count}")
    subprocess.run(command, check=True)
    after_count = database_count(trial_template)
    print(f"执行后数据库连接数：{after_count}")
    return 0 if after_count > before_count else 2


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="实验：用 BarTender XMLScript 给模板副本新增 CSV 数据库连接。")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="测试模板路径；脚本会复制副本，不改原文件")
    parser.add_argument("--csv", default=str(DEFAULT_RUNTIME_CSV), help="runtime CSV 路径")
    parser.add_argument("--create-csv", action="store_true", help="CSV 不存在时创建只有表头的测试 CSV")
    parser.add_argument(
        "--variant",
        choices=("mixed", "comma", "no_add_flag"),
        default="mixed",
        help="XMLScript 变体，默认 mixed",
    )
    parser.add_argument("--visible", action="store_true", help="执行 XMLScript 后不加 /X，保留 BarTender 窗口便于查看")
    return parser


def main() -> int:
    args = build_arg_parser().parse_args()
    template_path = Path(args.template).resolve()
    csv_path = Path(args.csv).resolve()

    if args.create_csv:
        ensure_runtime_csv(csv_path)

    try:
        return run_trial(template_path, csv_path, args.variant, args.visible)
    except subprocess.CalledProcessError as error:
        print(f"BarTender XMLScript 执行失败，退出码：{error.returncode}")
        return error.returncode or 1
    except FileNotFoundError as error:
        print(f"找不到 BarTender 可执行文件：{error.filename}")
        print("请设置环境变量 BARTENDER_EXE，或确认 BarTender 安装路径。")
        return 1
    except TrialError as error:
        print(error)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
