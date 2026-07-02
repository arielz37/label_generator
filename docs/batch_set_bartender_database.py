from __future__ import annotations

import argparse
import csv
import shutil
import subprocess
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable, List, Optional, Tuple
from xml.sax.saxutils import escape

from app_paths import LOG_DIR, PROJECT_ROOT, RUNTIME_DIR, TEMPLATE_ROOT

TOOL_DIR = Path(__file__).resolve().parent
LOG_RETENTION_DAYS = 7
DEFAULT_TEMPLATE_ROOT = TEMPLATE_ROOT
DEFAULT_RUNTIME_CSV = RUNTIME_DIR / "current_label.csv"
DEFAULT_REPORT = LOG_DIR / "batch_set_bartender_database_report.csv"
SCRIPT_DIR = LOG_DIR / "batch_set_bartender_database_xmlscripts"
XMLSCRIPT_TIMEOUT_SECONDS = 60

CSV_FIELDS = [
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
    "EXP_DATE",
    "EXP_DATE_YYMMDD",
    "EXP_DATE_YYYYMMDD",
    "LOT_NO",
    "HAS_INNER_LABEL",
    "QTY",
    "LABEL_COUNT",
    "LABEL_INDEX",
    "PACKAGE_PRD_NO",
    "PACKAGE_NAME",
]


def cleanup_old_logs() -> None:
    if not LOG_DIR.exists():
        return
    cutoff = datetime.now() - timedelta(days=LOG_RETENTION_DAYS)
    for path in LOG_DIR.iterdir():
        if not path.is_file():
            continue
        try:
            modified_at = datetime.fromtimestamp(path.stat().st_mtime)
        except OSError:
            continue
        if modified_at < cutoff:
            try:
                path.unlink()
            except OSError:
                pass

CONNECTION_COLLECTION_NAMES = (
    "DatabaseConnections",
    "Databases",
)

CONNECTION_PATH_PROPERTIES = (
    "FileName",
    "FilePath",
    "Path",
    "TextFileName",
    "TextFile",
    "DatabaseName",
)

NEW_CONNECTION_METHODS = (
    "AddTextFile",
    "AddTextFileConnection",
    "Add",
    "New",
)

FIRST_ROW_FIELD_NAME_PROPERTIES = (
    "UseFirstRowAsFieldNames",
    "UseFieldNamesFromFirstRecord",
    "FirstRowFieldNames",
    "FieldNamesInFirstRow",
)

INSPECT_NAMES = (
    "Databases",
    "DatabaseConnections",
    "Database",
    "DatabaseSetup",
    "TextFile",
    "ODBC",
    "Add",
    "Set",
    "File",
    "Name",
)

PROBE_MEMBER_NAMES = (
    "Formats",
    "Open",
    "Save",
    "SaveAs",
    "Close",
    "Databases",
    "DatabaseConnections",
    "DatabaseConnection",
    "DatabaseSetup",
    "Database",
    "Connections",
    "Add",
    "AddTextFile",
    "AddTextFileConnection",
    "AddDatabase",
    "AddDatabaseConnection",
    "New",
    "Create",
    "Insert",
    "Item",
    "Count",
    "FileName",
    "FilePath",
    "Path",
    "TextFileName",
    "TextFile",
    "DatabaseName",
    "UseFirstRowAsFieldNames",
    "UseFieldNamesFromFirstRecord",
    "SetFileName",
    "SetTextFileName",
    "SetDatabaseName",
    "QueryPrompts",
    "NamedSubStrings",
    "SubStrings",
)


class BatchBindError(Exception):
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


def iter_templates(template_root: Path, limit: int = 0) -> Iterable[Path]:
    files = sorted(
        {
            path
            for pattern in ("*.btw", "*.btW", "*.BTW")
            for path in template_root.rglob(pattern)
            if path.is_file()
        },
        key=lambda path: str(path).casefold(),
    )
    if limit > 0:
        files = files[:limit]
    return files


def dispatch_bartender(visible: bool):
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except ModuleNotFoundError as error:
        raise BatchBindError("缺少 pywin32，请先运行：pip install pywin32") from error
    except Exception:
        # The thread may already have a compatible COM apartment.
        pass

    try:
        import win32com.client
    except ModuleNotFoundError as error:
        raise BatchBindError("缺少 pywin32，请先运行：pip install pywin32") from error

    try:
        app = win32com.client.gencache.EnsureDispatch("BarTender.Application")
    except Exception:
        app = win32com.client.Dispatch("BarTender.Application")
    try:
        app.Visible = visible
    except Exception:
        pass
    return app


def open_format(app, template_path: Path):
    errors: List[str] = []
    open_attempts = (
        lambda: app.Formats.Open(str(template_path), False, ""),
        lambda: app.Formats.Open(str(template_path), False, "", False),
        lambda: app.Formats.Open(str(template_path), False),
        lambda: app.Formats.Open(str(template_path)),
    )
    for attempt in open_attempts:
        try:
            return attempt()
        except Exception as error:
            errors.append(str(error))
    raise BatchBindError("BarTender 无法打开模板：" + " | ".join(errors[-2:]))


def collection_count(collection) -> int:
    for name in ("Count", "count", "Length", "length"):
        try:
            value = getattr(collection, name)
            if callable(value):
                value = value()
            return int(value)
        except Exception:
            pass
    return 0


def collection_item(collection, index: int):
    attempts = (
        lambda: collection.Item(index),
        lambda: collection.Item(index + 1),
        lambda: collection[index],
        lambda: collection[index + 1],
    )
    for attempt in attempts:
        try:
            return attempt()
        except Exception:
            pass
    return None


def get_connection_collection(format_doc):
    for name in CONNECTION_COLLECTION_NAMES:
        try:
            collection = getattr(format_doc, name)
            if collection is not None:
                return name, collection
        except Exception:
            pass
    return "", None


def get_database_count(app, template_path: Path) -> int:
    format_doc = None
    try:
        format_doc = open_format(app, template_path)
        _collection_name, collection = get_connection_collection(format_doc)
        return collection_count(collection) if collection is not None else 0
    finally:
        if format_doc is not None:
            close_format(format_doc, save_changes=False)


def set_connection_property(connection, csv_path: Path) -> Tuple[bool, str]:
    csv_text = str(csv_path)
    for prop_name in CONNECTION_PATH_PROPERTIES:
        try:
            setattr(connection, prop_name, csv_text)
            return True, f"设置属性 {prop_name}"
        except Exception:
            pass

    method_attempts = (
        ("SetFileName", lambda: connection.SetFileName(csv_text)),
        ("SetTextFileName", lambda: connection.SetTextFileName(csv_text)),
        ("SetDatabaseName", lambda: connection.SetDatabaseName(csv_text)),
    )
    for method_name, attempt in method_attempts:
        try:
            attempt()
            return True, f"调用方法 {method_name}"
        except Exception:
            pass

    return False, "未找到可写入 CSV 路径的数据库连接属性"


def configure_connection_options(connection) -> List[str]:
    messages: List[str] = []
    for prop_name in FIRST_ROW_FIELD_NAME_PROPERTIES:
        try:
            setattr(connection, prop_name, True)
            messages.append(f"{prop_name}=True")
        except Exception:
            pass
    return messages


def try_add_connection(target, csv_path: Path) -> Tuple[bool, str]:
    csv_text = str(csv_path)
    errors: List[str] = []
    for method_name in NEW_CONNECTION_METHODS:
        try:
            method = getattr(target, method_name)
        except Exception:
            continue

        attempts = (
            lambda: method(csv_text),
            lambda: method(csv_text, True),
            lambda: method("Text File", csv_text),
            lambda: method("Text", csv_text),
            lambda: method(),
        )
        for attempt in attempts:
            try:
                connection = attempt()
                detail = f"新增连接方法 {method_name}"
                if connection is not None:
                    path_ok, path_message = set_connection_property(connection, csv_path)
                    option_messages = configure_connection_options(connection)
                    detail = "；".join([detail, path_message, *option_messages])
                    if not path_ok:
                        return False, detail
                return True, detail
            except Exception as error:
                errors.append(f"{method_name}: {error}")
    return False, "未找到可用的新增文本数据库连接方法" + ("；" + "；".join(errors[-3:]) if errors else "")


def add_database_connection(format_doc, csv_path: Path) -> Tuple[bool, str]:
    messages: List[str] = []
    for collection_name in CONNECTION_COLLECTION_NAMES:
        try:
            collection = getattr(format_doc, collection_name)
        except Exception:
            collection = None

        if collection is not None:
            ok, message = try_add_connection(collection, csv_path)
            messages.append(f"{collection_name}：{message}")
            if ok:
                return True, f"{collection_name}：{message}"

    ok, message = try_add_connection(format_doc, csv_path)
    messages.append(f"Format：{message}")
    if ok:
        return True, f"Format：{message}"

    return False, "；".join(messages)


def set_database_path(format_doc, csv_path: Path) -> Tuple[bool, str]:
    collection_name, collection = get_connection_collection(format_doc)
    if collection is None:
        ok, message = add_database_connection(format_doc, csv_path)
        return ok, "模板没有暴露 DatabaseConnections/Databases 集合；" + message

    count = collection_count(collection)
    if count <= 0:
        return (
            False,
            f"{collection_name} 集合为空；当前 BarTender COM 接口只暴露 Count/Item/QueryPrompts，"
            "不支持从脚本新增数据库连接。请先在模板中手动添加一次文本数据库连接，"
            "之后脚本可以批量更新已有连接的 CSV 路径。",
        )

    messages: List[str] = []
    for index in range(count):
        connection = collection_item(collection, index)
        if connection is None:
            messages.append(f"第 {index + 1} 个连接读取失败")
            continue

        ok, message = set_connection_property(connection, csv_path)
        messages.append(f"第 {index + 1} 个连接：{message}")
        if ok:
            return True, f"{collection_name}；" + "；".join(messages)

    return False, f"{collection_name}；" + "；".join(messages)


def build_add_database_xml(
    template_path: Path,
    csv_path: Path,
    recordset_name: str = "current_label",
    printer_name: str = "",
) -> str:
    printer_xml = ""
    if printer_name:
        printer_xml = f"""
      <PrintSetup>
        <Printer>{escape(printer_name)}</Printer>
      </PrintSetup>"""
    return f'''<?xml version="1.0" encoding="utf-8"?>
<XMLScript Version="2.0">
  <Command Name="AddCsvDatabase">
    <FormatSetup>
      <Format CloseAtEndOfJob="true" SaveAtEndOfJob="true">{escape(str(template_path))}</Format>{printer_xml}
      <RecordSet Name="{escape(recordset_name)}" Type="btTextFile" AddIfNone="true">
        <FileName>{escape(str(csv_path))}</FileName>
        <Delimitation>btDelimMixedQuoteAndComma</Delimitation>
        <UseFieldNamesFromFirstRecord>true</UseFieldNamesFromFirstRecord>
      </RecordSet>
    </FormatSetup>
  </Command>
</XMLScript>
'''


def add_database_with_xmlscript(template_path: Path, csv_path: Path, printer_name: str = "") -> Tuple[bool, str]:
    SCRIPT_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    script_path = SCRIPT_DIR / f"{timestamp}_{template_path.stem[:40]}_add_database.xml"
    script_path.write_text(build_add_database_xml(template_path, csv_path, printer_name=printer_name), encoding="utf-8")

    command = [find_bartender_exe(), f"/XMLSCRIPT={script_path}", "/X"]
    try:
        process = subprocess.Popen(command)
        try:
            process.wait(timeout=XMLSCRIPT_TIMEOUT_SECONDS)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait()
            raise
        if process.returncode != 0:
            raise subprocess.CalledProcessError(process.returncode, command)
    except FileNotFoundError as error:
        return False, f"找不到 BarTender 可执行文件：{error.filename}"
    except subprocess.TimeoutExpired:
        return False, f"XMLScript 执行超时：超过 {XMLSCRIPT_TIMEOUT_SECONDS} 秒。模板可能弹出 BarTender 确认窗口或无法自动保存，脚本：{script_path}"
    except subprocess.CalledProcessError as error:
        return False, f"XMLScript 执行失败，退出码：{error.returncode}，脚本：{script_path}"

    return True, f"XMLScript 新增数据库连接：{script_path}"


def save_format(format_doc) -> None:
    attempts = (
        lambda: format_doc.Save(),
        lambda: format_doc.SaveAs(str(format_doc.FullName), True),
    )
    errors: List[str] = []
    for attempt in attempts:
        try:
            attempt()
            return
        except Exception as error:
            errors.append(str(error))
    raise BatchBindError("模板保存失败：" + " | ".join(errors))


def close_format(format_doc, save_changes: bool = False) -> None:
    attempts = (
        lambda: format_doc.Close(1 if save_changes else 0),
        lambda: format_doc.Close(save_changes),
        lambda: format_doc.Close(),
    )
    for attempt in attempts:
        try:
            attempt()
            return
        except Exception:
            pass


def backup_template(template_path: Path, backup_root: Path, template_root: Path) -> Path:
    relative = template_path.relative_to(template_root)
    target = backup_root / relative
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, target)
    return target


def process_template(
    app,
    template_path: Path,
    template_root: Path,
    csv_path: Path,
    save: bool,
    backup_root: Optional[Path],
    printer_name: str = "",
    xmlscript_only: bool = False,
) -> dict:
    if xmlscript_only:
        if not save:
            return result_row(
                template_path,
                template_root,
                "dry-run成功",
                "XMLScript-only 模式；保存模式会直接写入 CSV 数据库连接。",
                "",
            )
        backup_path = str(backup_template(template_path, backup_root, template_root)) if backup_root is not None else ""
        ok, message = add_database_with_xmlscript(template_path, csv_path, printer_name=printer_name)
        if not ok:
            return result_row(template_path, template_root, "失败", message, backup_path)
        return result_row(template_path, template_root, "已保存", f"{message}；未使用 COM 二次验证", backup_path)

    initial_count_error = ""
    try:
        initial_count = get_database_count(app, template_path)
    except Exception as error:
        if not save:
            return result_row(
                template_path,
                template_root,
                "失败",
                f"无法用 COM 打开模板检查数据库连接数：{error}",
                "",
            )
        initial_count = 0
        initial_count_error = str(error)

    if not save and initial_count <= 0:
        return result_row(
            template_path,
            template_root,
            "dry-run成功",
            "模板当前无数据库连接；保存模式会用 XMLScript 新增 CSV 数据库连接。",
            "",
        )

    if save and initial_count <= 0:
        backup_path = str(backup_template(template_path, backup_root, template_root)) if backup_root is not None else ""
        ok, message = add_database_with_xmlscript(template_path, csv_path, printer_name=printer_name)
        if not ok:
            detail = message
            if initial_count_error:
                detail = f"COM 打开失败，改用 XMLScript；{message}；COM 错误：{initial_count_error}"
            return result_row(template_path, template_root, "失败", detail, backup_path)
        try:
            after_count = get_database_count(app, template_path)
        except Exception as error:
            detail = f"{message}；COM 无法再次打开模板验证数据库连接数：{error}"
            if initial_count_error:
                detail += f"；初始 COM 错误：{initial_count_error}"
            return result_row(template_path, template_root, "已保存", detail, backup_path)
        if after_count <= initial_count:
            return result_row(template_path, template_root, "失败", f"{message}；但数据库连接数仍为 {after_count}", backup_path)
        return result_row(template_path, template_root, "已保存", f"{message}；数据库连接数 {initial_count} -> {after_count}", backup_path)

    format_doc = None
    try:
        format_doc = open_format(app, template_path)
        ok, detail = set_database_path(format_doc, csv_path)
        if not ok:
            return result_row(template_path, template_root, "失败", detail, "")

        backup_path = ""
        if save:
            if backup_root is not None:
                backup_path = str(backup_template(template_path, backup_root, template_root))
            save_format(format_doc)
            status = "已保存"
        else:
            status = "dry-run成功"

        return result_row(template_path, template_root, status, detail, backup_path)
    except Exception as error:
        return result_row(template_path, template_root, "失败", str(error), "")
    finally:
        if format_doc is not None:
            close_format(format_doc, save_changes=False)


def scan_database_count(app, template_path: Path, template_root: Path) -> dict:
    format_doc = None
    try:
        format_doc = open_format(app, template_path)
        _collection_name, collection = get_connection_collection(format_doc)
        count = collection_count(collection) if collection is not None else 0
        status = "已有数据库连接" if count > 0 else "无数据库连接"
        return result_row(template_path, template_root, status, f"数据库连接数：{count}", "")
    except Exception as error:
        return result_row(template_path, template_root, "失败", str(error), "")
    finally:
        if format_doc is not None:
            close_format(format_doc, save_changes=False)


def result_row(template_path: Path, template_root: Path, status: str, message: str, backup_path: str) -> dict:
    try:
        relative = str(template_path.relative_to(template_root))
    except ValueError:
        relative = str(template_path)
    return {
        "模板相对路径": relative,
        "状态": status,
        "说明": message,
        "备份路径": backup_path,
    }


def write_report(rows: List[dict], report_path: Path) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with report_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=("模板相对路径", "状态", "说明", "备份路径"))
        writer.writeheader()
        writer.writerows(rows)


def run_batch_database_setup(
    template_root: Path = DEFAULT_TEMPLATE_ROOT,
    csv_path: Path = DEFAULT_RUNTIME_CSV,
    report_path: Path = DEFAULT_REPORT,
    save: bool = True,
    limit: int = 0,
    visible: bool = False,
    create_csv: bool = True,
    backup: bool = True,
    printer_name: str = "",
    xmlscript_only: bool = False,
    progress_callback=None,
) -> dict:
    template_root = Path(template_root).resolve()
    csv_path = Path(csv_path).resolve()
    report_path = Path(report_path).resolve()
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    cleanup_old_logs()

    if not template_root.exists():
        raise BatchBindError(f"模板根目录不存在：{template_root}")
    if create_csv:
        ensure_runtime_csv(csv_path)
    if not csv_path.exists():
        raise BatchBindError(f"runtime CSV 不存在：{csv_path}")

    app = None if xmlscript_only else dispatch_bartender(visible=visible)
    backup_root = None
    if save and backup:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_root = PROJECT_ROOT / "template_backups" / f"before_database_bind_{timestamp}"

    rows: List[dict] = []
    templates = list(iter_templates(template_root, limit=limit))
    try:
        for index, template_path in enumerate(templates, start=1):
            if progress_callback:
                progress_callback(index, len(templates), template_path)
            rows.append(
                process_template(
                    app=app,
                    template_path=template_path,
                    template_root=template_root,
                    csv_path=csv_path,
                    save=save,
                    backup_root=backup_root,
                    printer_name=printer_name,
                    xmlscript_only=xmlscript_only,
                )
            )
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass

    write_report(rows, report_path)
    success_count = sum(1 for row in rows if row["状态"] in ("dry-run成功", "已保存"))
    fail_count = len(rows) - success_count
    return {
        "rows": rows,
        "total": len(rows),
        "success": success_count,
        "fail": fail_count,
        "report": str(report_path),
        "backup_root": str(backup_root) if backup_root else "",
        "saved": save,
        "printer_name": printer_name,
        "xmlscript_only": xmlscript_only,
    }


def public_members(value) -> List[str]:
    names = set()
    try:
        names.update(name for name in dir(value) if not name.startswith("_"))
    except Exception:
        pass

    try:
        type_info = value._oleobj_.GetTypeInfo()
        type_attr = type_info.GetTypeAttr()
        func_count = type_attr[6]
        var_count = type_attr[7]
        for index in range(func_count):
            func_desc = type_info.GetFuncDesc(index)
            memid = func_desc[0]
            for name in type_info.GetNames(memid):
                if name:
                    names.add(name)
        for index in range(var_count):
            var_desc = type_info.GetVarDesc(index)
            memid = var_desc[0]
            for name in type_info.GetNames(memid):
                if name:
                    names.add(name)
    except Exception as error:
        if not names:
            return [f"<type info failed: {error}>"]

    return sorted(names)


def interesting_members(value) -> List[str]:
    names = public_members(value)
    matched = [
        name
        for name in names
        if any(keyword.lower() in name.lower() for keyword in INSPECT_NAMES)
    ]
    return matched or names[:100]


def probe_members(value) -> List[str]:
    rows = []
    try:
        oleobj = value._oleobj_
    except Exception as error:
        return [f"<no _oleobj_: {error}>"]

    for name in PROBE_MEMBER_NAMES:
        try:
            dispid = oleobj.GetIDsOfNames(name)
            rows.append(f"{name}: OK dispid={dispid}")
        except Exception as error:
            rows.append(f"{name}: NO {error}")
    return rows


def write_inspection(app, template_path: Path, report_path: Path) -> Path:
    inspect_path = report_path.with_name(report_path.stem + "_inspect.txt")
    lines: List[str] = []
    format_doc = None
    try:
        lines.append(f"Template: {template_path}")
        lines.append("")
        lines.append("[Application]")
        lines.extend(interesting_members(app))
        lines.append("")
        lines.append("[Application probe]")
        lines.extend(probe_members(app))
        lines.append("")

        format_doc = open_format(app, template_path)
        lines.append("[Format]")
        lines.extend(interesting_members(format_doc))
        lines.append("")
        lines.append("[Format probe]")
        lines.extend(probe_members(format_doc))
        lines.append("")

        for collection_name in CONNECTION_COLLECTION_NAMES:
            lines.append(f"[Format.{collection_name}]")
            try:
                collection = getattr(format_doc, collection_name)
                lines.append(f"count={collection_count(collection)}")
                lines.extend(interesting_members(collection))
                lines.append("")
                lines.append(f"[Format.{collection_name} probe]")
                lines.extend(probe_members(collection))
                first = collection_item(collection, 0)
                if first is not None:
                    lines.append("")
                    lines.append(f"[Format.{collection_name}.Item(1)]")
                    lines.extend(interesting_members(first))
                    lines.append("")
                    lines.append(f"[Format.{collection_name}.Item(1) probe]")
                    lines.extend(probe_members(first))
            except Exception as error:
                lines.append(f"<failed: {error}>")
            lines.append("")
    finally:
        if format_doc is not None:
            close_format(format_doc, save_changes=False)

    inspect_path.write_text("\n".join(lines), encoding="utf-8")
    return inspect_path


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="批量给 BarTender 模板设置 runtime CSV 数据库路径。")
    parser.add_argument("--template-root", default=str(DEFAULT_TEMPLATE_ROOT), help="模板根目录，默认 Templates/")
    parser.add_argument("--csv", default=str(DEFAULT_RUNTIME_CSV), help="要绑定的 runtime CSV 路径")
    parser.add_argument("--report", default=str(DEFAULT_REPORT), help="处理报告 CSV 路径")
    parser.add_argument("--save", action="store_true", help="实际保存模板；不加时只做 dry-run")
    parser.add_argument("--limit", type=int, default=0, help="只处理前 N 个模板，方便先小范围测试")
    parser.add_argument("--visible", action="store_true", help="显示 BarTender 窗口，便于调试")
    parser.add_argument("--printer", default="", help="打开并保存模板时指定 BarTender 打印机名称")
    parser.add_argument("--xmlscript-only", action="store_true", help="只用 XMLScript 处理模板，不使用 COM 预检查/验证")
    parser.add_argument("--no-backup", action="store_true", help="保存时不备份原模板")
    parser.add_argument("--create-csv", action="store_true", help="如果 runtime CSV 不存在，则创建一个只有表头的测试 CSV")
    parser.add_argument("--inspect-com", action="store_true", help="只打开第一个模板并导出 COM 成员列表，便于适配 BarTender 版本")
    parser.add_argument("--scan-db-counts", action="store_true", help="只扫描每个模板已有数据库连接数量，不修改模板")
    return parser


def main() -> int:
    args = build_arg_parser().parse_args()
    template_root = Path(args.template_root).resolve()
    csv_path = Path(args.csv).resolve()
    report_path = Path(args.report).resolve()
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    cleanup_old_logs()

    if not template_root.exists():
        print(f"模板根目录不存在：{template_root}")
        return 1
    if args.create_csv:
        ensure_runtime_csv(csv_path)
    if not csv_path.exists():
        print(f"runtime CSV 不存在：{csv_path}")
        print("可以先生成一次预览，或加 --create-csv 创建测试 CSV。")
        return 1

    app = None
    if not args.xmlscript_only:
        try:
            app = dispatch_bartender(visible=args.visible)
        except BatchBindError as error:
            print(error)
            return 1

    backup_root = None
    if args.save and not args.no_backup:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_root = PROJECT_ROOT / "template_backups" / f"before_database_bind_{timestamp}"

    rows: List[dict] = []
    templates = list(iter_templates(template_root, limit=args.limit))
    if args.inspect_com:
        if app is None:
            print("--inspect-com 需要 COM，不能和 --xmlscript-only 一起使用。")
            return 1
        if not templates:
            print("没有找到模板。")
            return 1
        try:
            inspect_path = write_inspection(app, templates[0], report_path)
            print(f"COM 检查报告：{inspect_path}")
            return 0
        finally:
            try:
                app.Quit()
            except Exception:
                pass

    if args.scan_db_counts:
        if app is None:
            print("--scan-db-counts 需要 COM，不能和 --xmlscript-only 一起使用。")
            return 1
        try:
            for index, template_path in enumerate(templates, start=1):
                print(f"[{index}/{len(templates)}] {template_path.relative_to(template_root)}")
                rows.append(scan_database_count(app, template_path, template_root))
        finally:
            try:
                app.Quit()
            except Exception:
                pass
        write_report(rows, report_path)
        print(f"扫描完成，报告：{report_path}")
        return 0

    try:
        for index, template_path in enumerate(templates, start=1):
            print(f"[{index}/{len(templates)}] {template_path.relative_to(template_root)}")
            rows.append(
                process_template(
                    app=app,
                    template_path=template_path,
                    template_root=template_root,
                    csv_path=csv_path,
                    save=args.save,
                    backup_root=backup_root,
                    printer_name=args.printer,
                    xmlscript_only=args.xmlscript_only,
                )
            )
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass

    write_report(rows, report_path)
    success_count = sum(1 for row in rows if row["状态"] in ("dry-run成功", "已保存"))
    fail_count = len(rows) - success_count
    print(f"处理完成：成功 {success_count}，失败 {fail_count}")
    print(f"报告：{report_path}")
    if not args.save:
        print("当前是 dry-run，没有保存模板。确认报告无误后加 --save。")
    return 0 if fail_count == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
