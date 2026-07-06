from __future__ import annotations

import os
import queue
import re
import shutil
import subprocess
import threading
from datetime import datetime, timedelta
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ModuleNotFoundError as error:
    raise SystemExit(
        "当前 Python 环境缺少 tkinter，无法启动桌面 GUI。\n"
        "tkinter 不能通过 pip install tkinter 安装；请安装完整 Python 3.8，"
        "并在安装时勾选 Tcl/Tk and IDLE，然后重新创建 .venv。"
    ) from error

from bartender_label_runner import OUTPUT_DIR, resolve_template_path
import app_paths
from app_paths import APP_SETTINGS_FILE, DEFAULT_PACKAGE_NAME_FILE, DEFAULT_TEMPLATE_MAPPING_FILE, DEFAULT_TEMPLATE_ROOT, LOG_DIR, PACKAGE_NAME_FILE, PROJECT_ROOT, TEMPLATE_MAPPING_FILE, TEMPLATE_ROOT, save_path_settings
from config import BARTENDER_EXE, BARTENDER_PRINTER, ERP_DATABASE, ERP_DRIVER, ERP_PASSWORD, ERP_SERVER, ERP_USER
from docs.batch_set_bartender_database import run_batch_database_setup
from erp_runtime_csv import fetch_bom_material_names_for_mo
from generate_template_mapping import update_template_mapping_from_directory
from label_service import find_template_for_mo, generate_label_preview, print_labels


APP_TITLE = "标签生成工作台"
LOG_RETENTION_DAYS = 7
MO_NO_PATTERN = re.compile(r"^MO\d{8}$")
SHELF_LIFE_DEFAULT_TEXT = "默认12个月"
PRINTER_DEFAULT_TEXT = "使用模板默认打印机"
PACKAGE_AUTO_TEXT = "自动判断"
COMPANY_PATH_SETTINGS_FILE = PROJECT_ROOT / "公司目录配置.txt"
COMPANY_PATH_SETTING_KEYS = {
    "template_root": ("模板搜索目录", "template_root"),
    "package_name_file": ("BOM料号名汇总", "BOM料号汇总", "package_name_file"),
    "template_mapping_file": ("模板映射表", "模板映射标", "template_mapping_file"),
}


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


def safe_log_segment(value: str) -> str:
    chars = []
    for char in value:
        if char.isalnum() or char in ("-", "_"):
            chars.append(char)
        else:
            chars.append("_")
    return "".join(chars).strip("_") or "work_order"


def is_valid_mo_no(value: str) -> bool:
    return bool(MO_NO_PATTERN.fullmatch((value or "").strip()))


SHELF_LIFE_OPTIONS = ("半年", "一年", "一年半", "两年", "3年", "5年", "其他")
SHELF_LIFE_MONTHS = {
    "半年": "6",
    "一年": "12",
    "一年半": "18",
    "两年": "24",
    "3年": "36",
    "5年": "60",
}


def list_windows_printers() -> list[str]:
    try:
        import win32print
    except Exception:
        return []

    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    printers = []
    try:
        for printer in win32print.EnumPrinters(flags):
            if len(printer) > 2 and printer[2]:
                printers.append(str(printer[2]))
    except Exception:
        return []
    return sorted(set(printers), key=str.casefold)


def get_windows_default_printer() -> str:
    try:
        import win32print
    except Exception:
        return ""

    try:
        return str(win32print.GetDefaultPrinter()).strip()
    except Exception:
        return ""


def is_virtual_printer_name(printer_name: str) -> bool:
    text = (printer_name or "").casefold()
    virtual_keywords = (
        "microsoft print to pdf",
        "microsoft xps",
        "onenote",
        "fax",
        "pdf",
    )
    return any(keyword in text for keyword in virtual_keywords)


def url_to_file_path(url: str) -> Path:
    prefix = "/files/"
    if not url.startswith(prefix):
        return Path(url)
    return OUTPUT_DIR / url[len(prefix):].replace("/", "\\")


class LabelGeneratorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1380x780")
        self.minsize(1280, 700)

        self.result = None
        self.current_label_type = "outer"
        self.current_image_index = 0
        self.current_photo = None
        self.template_lookup_mo = ""
        self.material_lookup_mo = ""
        self.worker_queue: "queue.Queue[tuple[str, object]]" = queue.Queue()
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        cleanup_old_logs()
        self.log_path = LOG_DIR / f"desktop_app_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        self.tools_menu = None
        self.template_db_menu_index = None
        self.mapping_menu_index = None
        self.environment_check_menu_index = None
        self.settings_menu_index = None
        self.entry_placeholders = {}
        self.placeholder_texts = set()
        self.warned_virtual_printers = set()

        self.create_menu()
        self.create_widgets()
        self.after(100, self.poll_worker_queue)

    def configure_styles(self) -> None:
        style = ttk.Style(self)
        style.configure("Primary.TButton", font=("", 11, "bold"), padding=(14, 8))

    def add_entry_placeholder(self, entry: ttk.Entry, variable: tk.StringVar, placeholder: str) -> None:
        self.entry_placeholders[entry] = (variable, placeholder)
        self.placeholder_texts.add(placeholder)

        def show_placeholder(_event=None) -> None:
            if not variable.get().strip():
                variable.set(placeholder)
                self.set_entry_foreground(entry, "#888888")

        def hide_placeholder(_event=None) -> None:
            if variable.get().strip() == placeholder:
                variable.set("")
                self.set_entry_foreground(entry, "#000000")

        entry.bind("<FocusIn>", hide_placeholder, add="+")
        entry.bind("<FocusOut>", show_placeholder, add="+")
        show_placeholder()

    def entry_value(self, variable: tk.StringVar) -> str:
        value = variable.get().strip()
        if value in self.placeholder_texts:
            return ""
        return value

    def refresh_entry_placeholder(self, entry: ttk.Entry) -> None:
        variable, placeholder = self.entry_placeholders.get(entry, (None, ""))
        if variable is not None and not variable.get().strip():
            variable.set(placeholder)
            self.set_entry_foreground(entry, "#888888")

    def set_entry_foreground(self, entry: ttk.Entry, color: str) -> None:
        try:
            entry.configure(foreground=color)
        except tk.TclError:
            pass

    def create_menu(self) -> None:
        menu_bar = tk.Menu(self)
        self.tools_menu = tk.Menu(menu_bar, tearoff=0)
        self.tools_menu.add_command(label="模板数据库设置", command=self.start_template_database_setup)
        self.template_db_menu_index = self.tools_menu.index("end")
        self.tools_menu.add_command(label="更新模板映射", command=self.start_template_mapping_update)
        self.mapping_menu_index = self.tools_menu.index("end")
        self.tools_menu.add_separator()
        self.tools_menu.add_command(label="环境自检", command=self.start_environment_check)
        self.environment_check_menu_index = self.tools_menu.index("end")
        self.tools_menu.add_command(label="更改参数和路径", command=self.open_settings_dialog)
        self.settings_menu_index = self.tools_menu.index("end")
        menu_bar.add_cascade(label="工具", menu=self.tools_menu)
        self.config(menu=menu_bar)

    def create_widgets(self) -> None:
        self.configure_styles()
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self, padding=12)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        top.columnconfigure(3, weight=0)
        top.columnconfigure(5, weight=1)
        top.columnconfigure(8, weight=1)

        ttk.Label(top, text="MO号").grid(row=0, column=0, padx=(0, 8))
        self.mo_var = tk.StringVar(value="MO26061617")
        self.mo_entry = ttk.Entry(top, textvariable=self.mo_var)
        self.mo_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.mo_entry.bind("<Return>", self.on_mo_template_lookup)
        self.mo_entry.bind("<FocusOut>", self.on_mo_template_lookup)
        self.mo_var.trace_add("write", self.on_mo_changed)

        ttk.Label(top, text="打印机").grid(row=1, column=0, padx=(0, 8), pady=(8, 0))
        printer_values = list_windows_printers()
        if PRINTER_DEFAULT_TEXT not in printer_values:
            printer_values = [PRINTER_DEFAULT_TEXT] + printer_values
        self.printer_var = tk.StringVar(value=BARTENDER_PRINTER or PRINTER_DEFAULT_TEXT)
        self.printer_combo = ttk.Combobox(
            top,
            textvariable=self.printer_var,
            values=printer_values,
            state="normal",
        )
        self.printer_combo.grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=(8, 0))

        ttk.Label(top, text="外标模板").grid(row=1, column=2, padx=(0, 8), pady=(8, 0))
        self.outer_template_var = tk.StringVar(value="")
        self.outer_template_entry = ttk.Entry(top, textvariable=self.outer_template_var)
        self.outer_template_entry.grid(row=1, column=3, columnspan=3, sticky="ew", padx=(0, 8), pady=(8, 0))
        self.outer_template_button = ttk.Button(
            top,
            text="修改外标",
            command=lambda: self.choose_template_file(self.outer_template_var),
        )
        self.outer_template_button.grid(row=1, column=6, padx=(0, 8), pady=(8, 0))

        ttk.Label(top, text="内标模板").grid(row=1, column=7, padx=(0, 8), pady=(8, 0))
        self.inner_template_var = tk.StringVar(value="")
        self.inner_template_entry = ttk.Entry(top, textvariable=self.inner_template_var)
        self.inner_template_entry.grid(row=1, column=8, columnspan=3, sticky="ew", padx=(0, 8), pady=(8, 0))
        self.inner_template_button = ttk.Button(
            top,
            text="修改内标",
            command=lambda: self.choose_template_file(self.inner_template_var),
        )
        self.inner_template_button.grid(row=1, column=11, padx=(0, 0), pady=(8, 0))

        self.open_outer_template_button = ttk.Button(
            top,
            text="打开外标模板",
            command=lambda: self.open_template_file("outer"),
        )
        self.open_outer_template_button.grid(row=3, column=3, padx=(0, 8), pady=(8, 0), sticky="w")
        self.open_inner_template_button = ttk.Button(
            top,
            text="打开内标模板",
            command=lambda: self.open_template_file("inner"),
        )
        self.open_inner_template_button.grid(row=3, column=8, padx=(0, 8), pady=(8, 0), sticky="w")

        self.outer_package_var = tk.StringVar(value=PACKAGE_AUTO_TEXT)
        ttk.Label(top, text="外标包装").grid(row=2, column=3, sticky="e", padx=(0, 8), pady=(8, 0))
        self.outer_package_combo = ttk.Combobox(
            top,
            textvariable=self.outer_package_var,
            values=(PACKAGE_AUTO_TEXT,),
            state="readonly",
        )
        self.outer_package_combo.grid(row=2, column=4, columnspan=3, sticky="ew", padx=(0, 8), pady=(8, 0))

        self.inner_package_var = tk.StringVar(value=PACKAGE_AUTO_TEXT)
        ttk.Label(top, text="内标包装").grid(row=2, column=8, sticky="e", padx=(0, 8), pady=(8, 0))
        self.inner_package_combo = ttk.Combobox(
            top,
            textvariable=self.inner_package_var,
            values=(PACKAGE_AUTO_TEXT,),
            state="readonly",
        )
        self.inner_package_combo.grid(row=2, column=9, columnspan=3, sticky="ew", padx=(0, 8), pady=(8, 0))

        self.outer_print_count_var = tk.StringVar(value="")
        self.outer_print_count_entry = ttk.Entry(top, textvariable=self.outer_print_count_var, width=6)

        self.inner_print_count_var = tk.StringVar(value="")
        self.inner_print_count_entry = ttk.Entry(top, textvariable=self.inner_print_count_var, width=6)

        ttk.Label(top, text="保质期").grid(row=0, column=2, sticky="e", padx=(0, 8))
        self.shelf_life_var = tk.StringVar(value=SHELF_LIFE_DEFAULT_TEXT)
        self.shelf_life_combo = ttk.Combobox(
            top,
            textvariable=self.shelf_life_var,
            values=(SHELF_LIFE_DEFAULT_TEXT,) + SHELF_LIFE_OPTIONS,
            width=8,
            state="readonly",
        )
        self.shelf_life_combo.grid(row=0, column=3, sticky="w", padx=(0, 8))
        self.shelf_life_combo.bind("<<ComboboxSelected>>", self.on_shelf_life_selected)
        self.custom_shelf_life_var = tk.StringVar(value="")
        self.custom_shelf_life_entry = ttk.Entry(top, textvariable=self.custom_shelf_life_var, width=8, state="disabled")
        self.custom_shelf_life_entry.grid(row=0, column=4, sticky="w", padx=(0, 4))
        self.custom_shelf_life_unit = ttk.Label(top, text="月")
        self.custom_shelf_life_unit.grid(row=0, column=5, sticky="w", padx=(0, 8))
        self.preview_button = ttk.Button(top, text="生成预览", command=self.start_preview, style="Primary.TButton")
        self.preview_button.grid(row=2, column=0, columnspan=2, sticky="ew", padx=(0, 8), pady=(8, 0))
        ttk.Label(top, text="外标张数").grid(row=0, column=8, sticky="e", padx=(0, 4))
        self.outer_print_count_entry.grid(row=0, column=9, padx=(0, 8))
        self.print_outer_button = ttk.Button(top, text="打印外标", command=lambda: self.start_print(["outer"]))
        self.print_outer_button.grid(row=0, column=10, padx=(0, 8))
        ttk.Label(top, text="内标张数").grid(row=0, column=11, sticky="e", padx=(0, 4))
        self.inner_print_count_entry.grid(row=0, column=12, padx=(0, 8))
        self.print_inner_button = ttk.Button(top, text="打印内标", command=lambda: self.start_print(["inner"]))
        self.print_inner_button.grid(row=0, column=13, padx=(0, 8))
        self.print_all_button = ttk.Button(top, text="全部打印", command=self.print_all)
        self.print_all_button.grid(row=0, column=14)
        body = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        body.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))

        left = ttk.Frame(body, padding=10)
        right = ttk.Frame(body, padding=10)
        body.add(left, weight=4)
        body.add(right, weight=2)

        left.rowconfigure(2, weight=1)
        left.columnconfigure(0, weight=1)

        tabs = ttk.Frame(left)
        tabs.grid(row=0, column=0, sticky="ew")
        self.outer_tab = ttk.Button(tabs, text="外标预览", command=lambda: self.switch_label_type("outer"))
        self.outer_tab.grid(row=0, column=0, padx=(0, 8))
        self.inner_tab = ttk.Button(tabs, text="内标预览", command=lambda: self.switch_label_type("inner"))
        self.inner_tab.grid(row=0, column=1)

        nav = ttk.Frame(left)
        nav.grid(row=1, column=0, sticky="ew", pady=8)
        nav.columnconfigure(1, weight=1)
        self.prev_button = ttk.Button(nav, text="上一页", command=self.prev_image)
        self.prev_button.grid(row=0, column=0)
        self.counter_var = tk.StringVar(value="0 / 0")
        ttk.Label(nav, textvariable=self.counter_var, anchor="center").grid(row=0, column=1)
        self.next_button = ttk.Button(nav, text="下一页", command=self.next_image)
        self.next_button.grid(row=0, column=2)

        self.preview_frame = ttk.Frame(left, relief="sunken")
        self.preview_frame.grid(row=2, column=0, sticky="nsew")
        self.preview_frame.rowconfigure(0, weight=1)
        self.preview_frame.columnconfigure(0, weight=1)
        self.preview_label = ttk.Label(self.preview_frame, text="生成预览后会显示标签图片", anchor="center")
        self.preview_label.grid(row=0, column=0, sticky="nsew")

        right.columnconfigure(0, weight=1)
        ttk.Label(right, text="工单信息", font=("", 12, "bold")).grid(row=0, column=0, sticky="w")
        self.info_text = tk.Text(right, height=15, wrap="word")
        self.info_text.grid(row=1, column=0, sticky="ew", pady=(8, 14))
        self.info_text.tag_configure("warning", foreground="#b00020", font=("", 10, "bold"))
        self.info_text.configure(state="disabled")

        ttk.Label(right, text="运行日志", font=("", 12, "bold")).grid(row=2, column=0, sticky="w")
        self.log_text = tk.Text(right, wrap="word")
        self.log_text.grid(row=3, column=0, sticky="nsew", pady=(8, 0))
        right.rowconfigure(3, weight=1)

        self.status_var = tk.StringVar(value="等待输入 MO 号")
        status = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w", padding=6)
        status.grid(row=2, column=0, sticky="ew")

        self.configure_entry_placeholders()
        self.update_action_state()

    def configure_entry_placeholders(self) -> None:
        self.add_entry_placeholder(self.outer_template_entry, self.outer_template_var, "自动匹配外标模板")
        self.add_entry_placeholder(self.inner_template_entry, self.inner_template_var, "自动匹配内标模板")
        self.add_entry_placeholder(self.outer_print_count_entry, self.outer_print_count_var, "全部")
        self.add_entry_placeholder(self.inner_print_count_entry, self.inner_print_count_var, "全部")

    def log(self, message: str) -> None:
        timestamped = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        with self.log_path.open("a", encoding="utf-8") as file:
            file.write(timestamped + "\n")

    def on_shelf_life_selected(self, _event=None) -> None:
        if self.shelf_life_var.get() == "其他":
            self.custom_shelf_life_entry.configure(state="normal")
            self.custom_shelf_life_entry.focus_set()
        else:
            self.custom_shelf_life_var.set("")
            self.refresh_entry_placeholder(self.custom_shelf_life_entry)
            self.custom_shelf_life_entry.configure(state="disabled")

    def on_mo_changed(self, *_args) -> None:
        mo_no = self.mo_var.get().strip()
        if self.template_lookup_mo and mo_no != self.template_lookup_mo:
            self.outer_template_var.set("")
            self.inner_template_var.set("")
            self.refresh_entry_placeholder(self.outer_template_entry)
            self.refresh_entry_placeholder(self.inner_template_entry)
            self.template_lookup_mo = ""
        if self.material_lookup_mo and mo_no != self.material_lookup_mo:
            self.reset_package_options()

    def on_mo_template_lookup(self, _event=None) -> None:
        self.populate_template_fields(show_errors=False)

    def populate_template_fields(self, show_errors: bool = False) -> bool:
        mo_no = self.mo_var.get().strip()
        if not mo_no:
            return False
        self.populate_package_options(mo_no)
        if self.template_lookup_mo == mo_no and (self.entry_value(self.outer_template_var) or self.entry_value(self.inner_template_var)):
            return True

        try:
            outer_template = find_template_for_mo(mo_no, "outer")
            self.outer_template_var.set(outer_template)
            self.set_entry_foreground(self.outer_template_entry, "#000000")
        except Exception as error:
            if show_errors:
                messagebox.showerror(APP_TITLE, str(error))
            else:
                self.status_var.set("未自动识别到外标模板")
                self.log(f"未自动识别到外标模板：{error}")
            return False

        try:
            inner_template = find_template_for_mo(mo_no, "inner")
        except Exception as error:
            inner_template = ""
            self.log(f"未自动识别到内标模板：{error}")
        self.inner_template_var.set(inner_template)
        if inner_template:
            self.set_entry_foreground(self.inner_template_entry, "#000000")
        else:
            self.refresh_entry_placeholder(self.inner_template_entry)
        self.template_lookup_mo = mo_no
        self.status_var.set("已自动识别模板路径")
        self.log(f"自动识别外标模板：{outer_template}")
        if inner_template:
            self.log(f"自动识别内标模板：{inner_template}")
        return True

    def reset_package_options(self) -> None:
        self.outer_package_combo.configure(values=(PACKAGE_AUTO_TEXT,))
        self.inner_package_combo.configure(values=(PACKAGE_AUTO_TEXT,))
        self.outer_package_var.set(PACKAGE_AUTO_TEXT)
        self.inner_package_var.set(PACKAGE_AUTO_TEXT)
        self.material_lookup_mo = ""

    def populate_package_options(self, mo_no: str) -> None:
        if self.material_lookup_mo == mo_no:
            return
        try:
            names = fetch_bom_material_names_for_mo(mo_no)
        except Exception as error:
            self.reset_package_options()
            self.log(f"未加载BOM原材料品名：{error}")
            return

        values = (PACKAGE_AUTO_TEXT,) + tuple(names)
        current_outer = self.outer_package_var.get().strip()
        current_inner = self.inner_package_var.get().strip()
        self.outer_package_combo.configure(values=values)
        self.inner_package_combo.configure(values=values)
        self.outer_package_var.set(current_outer if current_outer in values else PACKAGE_AUTO_TEXT)
        self.inner_package_var.set(current_inner if current_inner in values else PACKAGE_AUTO_TEXT)
        self.material_lookup_mo = mo_no
        self.log(f"已加载BOM原材料品名：{len(names)} 项")

    def choose_template_file(self, target_var: tk.StringVar) -> None:
        path_text = filedialog.askopenfilename(
            title="选择 BarTender 模板",
            initialdir=str(TEMPLATE_ROOT),
            filetypes=(("BarTender 模板", "*.btw *.btW *.BTW"), ("所有文件", "*.*")),
        )
        if not path_text:
            return
        path = Path(path_text).resolve()
        try:
            value = str(path.relative_to(TEMPLATE_ROOT))
        except ValueError:
            value = str(path)
        target_var.set(value)
        if target_var is self.outer_template_var:
            self.set_entry_foreground(self.outer_template_entry, "#000000")
        elif target_var is self.inner_template_var:
            self.set_entry_foreground(self.inner_template_entry, "#000000")

    def open_settings_dialog(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("更改参数和路径")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)

        template_root_var = tk.StringVar(value=str(TEMPLATE_ROOT))
        package_name_file_var = tk.StringVar(value=str(PACKAGE_NAME_FILE))
        template_mapping_file_var = tk.StringVar(value=str(TEMPLATE_MAPPING_FILE))

        ttk.Label(frame, text="模板搜索目录").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
        template_entry = ttk.Entry(frame, textvariable=template_root_var, width=72)
        template_entry.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        ttk.Button(
            frame,
            text="选择目录",
            command=lambda: self.choose_settings_directory(template_root_var),
        ).grid(row=0, column=2, padx=(8, 0), pady=(0, 8))

        ttk.Label(frame, text="BOM料号名汇总").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
        package_entry = ttk.Entry(frame, textvariable=package_name_file_var, width=72)
        package_entry.grid(row=1, column=1, sticky="ew", pady=(0, 8))
        ttk.Button(
            frame,
            text="选择文件",
            command=lambda: self.choose_settings_file(package_name_file_var),
        ).grid(row=1, column=2, padx=(8, 0), pady=(0, 8))

        ttk.Label(frame, text="模板映射表").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
        mapping_entry = ttk.Entry(frame, textvariable=template_mapping_file_var, width=72)
        mapping_entry.grid(row=2, column=1, sticky="ew", pady=(0, 8))
        ttk.Button(
            frame,
            text="选择文件",
            command=lambda: self.choose_template_mapping_file(template_mapping_file_var),
        ).grid(row=2, column=2, padx=(8, 0), pady=(0, 8))

        hint = (
            f"配置会保存到：{APP_SETTINGS_FILE}\n"
            "保存后立即影响模板查找、模板映射更新和BOM内包材名称匹配。"
        )
        ttk.Label(frame, text=hint, foreground="#555555").grid(row=3, column=0, columnspan=3, sticky="w", pady=(2, 12))

        buttons = ttk.Frame(frame)
        buttons.grid(row=4, column=0, columnspan=3, sticky="e")
        ttk.Button(
            buttons,
            text="一键配置公司目录",
            command=lambda: self.apply_company_path_preset(
                dialog,
                template_root_var,
                package_name_file_var,
                template_mapping_file_var,
            ),
        ).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(
            buttons,
            text="恢复默认",
            command=lambda: self.reset_settings_paths(template_root_var, package_name_file_var, template_mapping_file_var),
        ).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(buttons, text="取消", command=dialog.destroy).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(
            buttons,
            text="保存",
            command=lambda: self.save_settings_paths(
                dialog,
                template_root_var,
                package_name_file_var,
                template_mapping_file_var,
            ),
        ).grid(row=0, column=3)

        template_entry.focus_set()
        self.wait_window(dialog)

    def choose_settings_directory(self, target_var: tk.StringVar) -> None:
        initial = target_var.get().strip() or str(TEMPLATE_ROOT)
        selected = filedialog.askdirectory(title="选择模板搜索目录", initialdir=initial)
        if selected:
            target_var.set(str(Path(selected).resolve()))

    def choose_settings_file(self, target_var: tk.StringVar) -> None:
        initial_path = Path(target_var.get().strip() or str(PACKAGE_NAME_FILE))
        selected = filedialog.askopenfilename(
            title="选择BOM料号名汇总文件",
            initialdir=str(initial_path.parent if initial_path.parent.exists() else PROJECT_ROOT),
            filetypes=(("文本文件", "*.txt"), ("所有文件", "*.*")),
        )
        if selected:
            target_var.set(str(Path(selected).resolve()))

    def choose_template_mapping_file(self, target_var: tk.StringVar) -> None:
        initial_path = Path(target_var.get().strip() or str(TEMPLATE_MAPPING_FILE))
        selected = filedialog.askopenfilename(
            title="选择模板映射表",
            initialdir=str(initial_path.parent if initial_path.parent.exists() else PROJECT_ROOT),
            filetypes=(("Excel/CSV", "*.xlsx *.csv"), ("Excel 文件", "*.xlsx"), ("CSV 文件", "*.csv"), ("所有文件", "*.*")),
        )
        if selected:
            target_var.set(str(Path(selected).resolve()))

    def reset_settings_paths(
        self,
        template_root_var: tk.StringVar,
        package_name_file_var: tk.StringVar,
        template_mapping_file_var: tk.StringVar,
    ) -> None:
        template_root_var.set(str(DEFAULT_TEMPLATE_ROOT))
        package_name_file_var.set(str(DEFAULT_PACKAGE_NAME_FILE))
        template_mapping_file_var.set(str(DEFAULT_TEMPLATE_MAPPING_FILE))

    def load_company_path_preset(self) -> dict[str, str]:
        if not COMPANY_PATH_SETTINGS_FILE.exists():
            raise FileNotFoundError(f"公司目录配置文件不存在：{COMPANY_PATH_SETTINGS_FILE}")

        values = {}
        with COMPANY_PATH_SETTINGS_FILE.open("r", encoding="utf-8") as file:
            for line in file:
                text = line.strip()
                if not text or text.startswith("#") or "=" not in text:
                    continue
                key, value = [part.strip() for part in text.split("=", 1)]
                if not value:
                    continue
                for target_key, aliases in COMPANY_PATH_SETTING_KEYS.items():
                    if key in aliases:
                        values[target_key] = value
                        break

        missing = [
            aliases[0]
            for target_key, aliases in COMPANY_PATH_SETTING_KEYS.items()
            if target_key not in values
        ]
        if missing:
            raise ValueError("公司目录配置缺少字段：" + "、".join(missing))
        return values

    def apply_company_path_preset(
        self,
        dialog: tk.Toplevel,
        template_root_var: tk.StringVar,
        package_name_file_var: tk.StringVar,
        template_mapping_file_var: tk.StringVar,
    ) -> None:
        try:
            values = self.load_company_path_preset()
        except (OSError, ValueError) as error:
            messagebox.showwarning(APP_TITLE, str(error), parent=dialog)
            return

        template_root_var.set(values["template_root"])
        package_name_file_var.set(values["package_name_file"])
        template_mapping_file_var.set(values["template_mapping_file"])
        self.log(f"已读取公司目录配置：{COMPANY_PATH_SETTINGS_FILE}")

    def normalize_settings_path(self, value: str) -> Path:
        path = Path(value.strip())
        if not path.is_absolute():
            path = PROJECT_ROOT / path
        return path.resolve()

    def save_settings_paths(
        self,
        dialog: tk.Toplevel,
        template_root_var: tk.StringVar,
        package_name_file_var: tk.StringVar,
        template_mapping_file_var: tk.StringVar,
    ) -> None:
        try:
            template_root = self.normalize_settings_path(template_root_var.get())
            package_name_file = self.normalize_settings_path(package_name_file_var.get())
            template_mapping_file = self.normalize_settings_path(template_mapping_file_var.get())
        except OSError as error:
            messagebox.showwarning(APP_TITLE, f"路径无效：{error}", parent=dialog)
            return

        if not template_root.exists() or not template_root.is_dir():
            messagebox.showwarning(APP_TITLE, f"模板搜索目录不存在：\n{template_root}", parent=dialog)
            return
        if not package_name_file.exists() or not package_name_file.is_file():
            messagebox.showwarning(APP_TITLE, f"BOM料号名汇总文件不存在：\n{package_name_file}", parent=dialog)
            return
        if not template_mapping_file.exists() or not template_mapping_file.is_file():
            messagebox.showwarning(APP_TITLE, f"模板映射表不存在：\n{template_mapping_file}", parent=dialog)
            return
        if template_mapping_file.suffix.lower() not in (".xlsx", ".csv"):
            messagebox.showwarning(APP_TITLE, "模板映射表只支持 .xlsx 或 .csv 文件。", parent=dialog)
            return

        try:
            save_path_settings(template_root, package_name_file, template_mapping_file)
            self.apply_path_settings(template_root, package_name_file, template_mapping_file)
        except OSError as error:
            messagebox.showerror(APP_TITLE, f"保存配置失败：\n{error}", parent=dialog)
            return

        self.log(f"已更新模板搜索目录：{template_root}")
        self.log(f"已更新BOM料号名汇总：{package_name_file}")
        self.log(f"已更新模板映射表：{template_mapping_file}")
        self.status_var.set("参数和路径已更新")
        messagebox.showinfo(APP_TITLE, "参数和路径已保存。", parent=dialog)
        dialog.destroy()

    def apply_path_settings(self, template_root: Path, package_name_file: Path, template_mapping_file: Path) -> None:
        global TEMPLATE_ROOT, PACKAGE_NAME_FILE, TEMPLATE_MAPPING_FILE
        TEMPLATE_ROOT = template_root.resolve()
        PACKAGE_NAME_FILE = package_name_file.resolve()
        TEMPLATE_MAPPING_FILE = template_mapping_file.resolve()
        app_paths.TEMPLATE_ROOT = TEMPLATE_ROOT
        app_paths.PACKAGE_NAME_FILE = PACKAGE_NAME_FILE
        app_paths.TEMPLATE_MAPPING_FILE = TEMPLATE_MAPPING_FILE

        import bartender_label_runner
        import bartender_preview_runner
        import docs.batch_set_bartender_database as batch_database
        import erp_runtime_csv
        import generate_template_mapping
        import template_mapping_lookup

        bartender_label_runner.TEMPLATE_ROOT = TEMPLATE_ROOT
        bartender_preview_runner.TEMPLATE_ROOT = TEMPLATE_ROOT
        batch_database.DEFAULT_TEMPLATE_ROOT = TEMPLATE_ROOT
        erp_runtime_csv.PACKAGE_NAME_FILE = PACKAGE_NAME_FILE
        erp_runtime_csv.PACKAGE_CACHE.clear()
        generate_template_mapping.TEMPLATE_ROOT = TEMPLATE_ROOT
        generate_template_mapping.DEFAULT_TEMPLATE_ROOT = TEMPLATE_ROOT
        generate_template_mapping.TEMPLATE_MAPPING_FILE = TEMPLATE_MAPPING_FILE
        generate_template_mapping.DEFAULT_OUTPUT = TEMPLATE_MAPPING_FILE
        template_mapping_lookup.app_paths.TEMPLATE_MAPPING_FILE = TEMPLATE_MAPPING_FILE

    def get_template_overrides(self) -> tuple[str, str]:
        return self.entry_value(self.outer_template_var), self.entry_value(self.inner_template_var)

    def get_package_overrides(self) -> tuple[str, str]:
        outer_package = self.outer_package_var.get().strip()
        inner_package = self.inner_package_var.get().strip()
        if outer_package == PACKAGE_AUTO_TEXT:
            outer_package = ""
        if inner_package == PACKAGE_AUTO_TEXT:
            inner_package = ""
        return inner_package, outer_package

    def get_selected_printer(self) -> str:
        printer_name = self.printer_var.get().strip()
        if printer_name == PRINTER_DEFAULT_TEXT:
            return ""
        if printer_name and is_virtual_printer_name(printer_name):
            message = (
                f"当前选择的是虚拟打印机：{printer_name}。"
                "PDF/XPS/OneNote/Fax 这类虚拟打印机可能导致 BarTender 预览或打印结果异常。"
            )
            self.log(message)
            if printer_name not in self.warned_virtual_printers:
                self.warned_virtual_printers.add(printer_name)
                messagebox.showwarning(APP_TITLE, message)
        return printer_name

    def open_template_file(self, label_type: str) -> None:
        mo_no = self.mo_var.get().strip()
        if not mo_no:
            messagebox.showwarning(APP_TITLE, "请输入 MO 号")
            return
        outer_template, inner_template = self.get_template_overrides()
        try:
            template = find_template_for_mo(
                mo_no,
                label_type,
                outer_template_override=outer_template,
                inner_template_override=inner_template,
            )
            if label_type == "outer":
                self.outer_template_var.set(template)
                self.set_entry_foreground(self.outer_template_entry, "#000000")
            else:
                self.inner_template_var.set(template)
                self.set_entry_foreground(self.inner_template_entry, "#000000")
            self.template_lookup_mo = mo_no
            template_path = resolve_template_path(template)
            if not template_path.exists():
                messagebox.showerror(APP_TITLE, f"模板文件不存在：\n{template_path}")
                return
            subprocess.Popen([BARTENDER_EXE, str(template_path)])
        except FileNotFoundError:
            messagebox.showerror(APP_TITLE, f"找不到 BarTender：\n{BARTENDER_EXE}")
            return
        except Exception as error:
            messagebox.showerror(APP_TITLE, str(error))
            return

        label_text = "外标" if label_type == "outer" else "内标"
        self.log(f"打开{label_text}模板：{template_path}")

    def get_shelf_life_value(self) -> str:
        selected = self.shelf_life_var.get().strip()
        if not selected or selected == SHELF_LIFE_DEFAULT_TEXT:
            return ""
        if selected == "其他":
            months = self.entry_value(self.custom_shelf_life_var)
            if not months:
                raise ValueError("请选择“其他”时，请输入保质期月份。")
            if not months.isdigit() or int(months) <= 0:
                raise ValueError("保质期月份必须是大于 0 的整数。")
            return months
        return SHELF_LIFE_MONTHS[selected]

    def set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        if not busy:
            self.preview_button.configure(text="生成预览")
        self.preview_button.configure(state=state)
        inner_result = (self.result or {}).get("inner") or {}
        inner_can_print = bool(inner_result.get("can_print"))
        self.print_outer_button.configure(state=state if self.result else "disabled")
        self.print_inner_button.configure(state=state if inner_can_print else "disabled")
        self.print_all_button.configure(state=state if self.result else "disabled")
        if self.tools_menu is not None:
            if self.template_db_menu_index is not None:
                self.tools_menu.entryconfig(self.template_db_menu_index, state=state)
            if self.mapping_menu_index is not None:
                self.tools_menu.entryconfig(self.mapping_menu_index, state=state)
            if self.environment_check_menu_index is not None:
                self.tools_menu.entryconfig(self.environment_check_menu_index, state=state)
            if self.settings_menu_index is not None:
                self.tools_menu.entryconfig(self.settings_menu_index, state=state)
        self.printer_combo.configure(state="disabled" if busy else "normal")
        self.outer_package_combo.configure(state="disabled" if busy else "readonly")
        self.inner_package_combo.configure(state="disabled" if busy else "readonly")
        template_state = "disabled" if busy else "normal"
        self.outer_template_entry.configure(state=template_state)
        self.inner_template_entry.configure(state=template_state)
        self.outer_template_button.configure(state=template_state)
        self.inner_template_button.configure(state=template_state)
        self.open_outer_template_button.configure(state=template_state)
        self.open_inner_template_button.configure(state=template_state)
        self.outer_print_count_entry.configure(state=template_state)
        self.inner_print_count_entry.configure(state=template_state)
        self.shelf_life_combo.configure(state="disabled" if busy else "readonly")
        if busy:
            self.custom_shelf_life_entry.configure(state="disabled")
        else:
            self.on_shelf_life_selected()

    def start_preview(self) -> None:
        mo_no = self.mo_var.get().strip()
        try:
            shelf_life = self.get_shelf_life_value()
        except ValueError as error:
            messagebox.showwarning(APP_TITLE, str(error))
            return
        if not mo_no:
            messagebox.showwarning(APP_TITLE, "请输入 MO 号")
            return

        if not is_valid_mo_no(mo_no):
            message = f"MO号格式错误：{mo_no}。正确格式示例：MO88888888"
            self.log(message)
            messagebox.showwarning(APP_TITLE, message)
            return

        try:
            printer_name = self.get_selected_printer()
        except ValueError as error:
            messagebox.showwarning(APP_TITLE, str(error))
            return
        self.set_busy(True)
        self.preview_button.configure(text="预览生成中...")
        self.status_var.set("正在生成预览...")
        outer_template, inner_template = self.get_template_overrides()
        if not outer_template and not inner_template:
            if not self.populate_template_fields(show_errors=True):
                self.set_busy(False)
                self.update_action_state()
                return
            outer_template, inner_template = self.get_template_overrides()
        else:
            self.populate_package_options(mo_no)
        inner_package_name, outer_package_name = self.get_package_overrides()
        self.log(f"开始生成预览：{mo_no}")
        if printer_name:
            self.log(f"预览打印机：{printer_name}")
        else:
            self.log(f"预览打印机：{PRINTER_DEFAULT_TEXT}")
        if outer_template:
            self.log(f"UI指定外标模板：{outer_template}")
        if inner_template:
            self.log(f"UI指定内标模板：{inner_template}")
        if outer_package_name:
            self.log(f"UI指定外标包装：{outer_package_name}")
        if inner_package_name:
            self.log(f"UI指定内标包装：{inner_package_name}")
        threading.Thread(
            target=self.preview_worker,
            args=(mo_no, shelf_life, printer_name, outer_template, inner_template, inner_package_name, outer_package_name),
            daemon=True,
        ).start()

    def preview_worker(
        self,
        mo_no: str,
        shelf_life: str,
        printer_name: str,
        outer_template: str,
        inner_template: str,
        inner_package_name: str,
        outer_package_name: str,
    ) -> None:
        try:
            self.worker_queue.put((
                "preview_result",
                generate_label_preview(
                    mo_no,
                    shelf_life=shelf_life,
                    printer_name=printer_name,
                    outer_template_override=outer_template,
                    inner_template_override=inner_template,
                    inner_package_name=inner_package_name,
                    outer_package_name=outer_package_name,
                ),
            ))
        except Exception as error:
            self.worker_queue.put(("error", error))

    def get_label_total_count(self, label_type: str) -> int:
        if not self.result:
            return 0
        label_result = self.result.get(label_type)
        if not label_result:
            return 0
        try:
            return int(str(label_result.get("total_label_count", "") or "0"))
        except (TypeError, ValueError):
            return 0

    def parse_print_row_limit(self, label_type: str, value: str):
        text = (value or "").strip()
        if not text:
            return None
        label_name = "外标" if label_type == "outer" else "内标"
        max_count = self.get_label_total_count(label_type)
        if not text.isdigit():
            raise ValueError(f"{label_name}打印张数必须是 1 到 {max_count} 的整数")
        count = int(text)
        if count < 1 or count > max_count:
            raise ValueError(f"{label_name}打印张数超出范围：请输入 1 到 {max_count}")
        return count

    def get_print_row_limits(self, label_types):
        limits = {}
        if "outer" in label_types:
            outer_limit = self.parse_print_row_limit("outer", self.entry_value(self.outer_print_count_var))
            if outer_limit is not None:
                limits["outer"] = outer_limit
        if "inner" in label_types:
            inner_limit = self.parse_print_row_limit("inner", self.entry_value(self.inner_print_count_var))
            if inner_limit is not None:
                limits["inner"] = inner_limit
        return limits

    def describe_print_row_limits(self, label_types, print_row_limits) -> str:
        descriptions = []
        for label_type in label_types:
            label_name = "外标" if label_type == "outer" else "内标"
            max_count = self.get_label_total_count(label_type)
            limit = print_row_limits.get(label_type)
            if limit is None:
                descriptions.append(f"{label_name}：全部 {max_count} 张")
            else:
                descriptions.append(f"{label_name}：仅打印前 {limit} 张（共 {max_count} 张）")
        return "\n".join(descriptions)

    def start_print(self, label_types) -> None:
        if not self.result:
            return
        mo_no = self.result["mo_no"]
        shelf_life = self.result.get("shelf_life_input", "")
        try:
            current_shelf_life = self.get_shelf_life_value()
        except ValueError as error:
            messagebox.showwarning(APP_TITLE, str(error))
            return
        if current_shelf_life != shelf_life:
            messagebox.showwarning(APP_TITLE, "保质期已修改，请重新生成预览后再打印。")
            return
        try:
            printer_name = self.get_selected_printer()
        except ValueError as error:
            messagebox.showwarning(APP_TITLE, str(error))
            return
        if printer_name != self.result.get("printer_name", ""):
            messagebox.showwarning(APP_TITLE, "打印机已修改，请重新生成预览后再打印。")
            return
        outer_template, inner_template = self.get_template_overrides()
        inner_package_name, outer_package_name = self.get_package_overrides()
        if (
            outer_template != self.result.get("outer_template_override", "")
            or inner_template != self.result.get("inner_template_override", "")
        ):
            messagebox.showwarning(APP_TITLE, "指定模板已修改，请重新生成预览后再打印。")
            return
        if (
            inner_package_name != self.result.get("inner_package_override", "")
            or outer_package_name != self.result.get("outer_package_override", "")
        ):
            messagebox.showwarning(APP_TITLE, "指定包装已修改，请重新生成预览后再打印。")
            return
        if not messagebox.askyesno(APP_TITLE, f"确认打印 {mo_no} 的 {', '.join(label_types)} 标签？"):
            return
        self.set_busy(True)
        self.status_var.set("正在调用 BarTender 打印...")
        self.log(f"开始打印：{', '.join(label_types)}")
        if printer_name:
            self.log(f"打印机：{printer_name}")
        else:
            self.log(f"打印机：{PRINTER_DEFAULT_TEXT}")
        if outer_package_name:
            self.log(f"打印外标包装：{outer_package_name}")
        if inner_package_name:
            self.log(f"打印内标包装：{inner_package_name}")
        try:
            print_row_limits = self.get_print_row_limits(label_types)
        except ValueError as error:
            messagebox.showwarning(APP_TITLE, str(error))
            self.set_busy(False)
            self.update_action_state()
            return
        limit_description = self.describe_print_row_limits(label_types, print_row_limits)
        if limit_description:
            self.log(limit_description)
        threading.Thread(
            target=self.print_worker,
            args=(
                mo_no,
                label_types,
                shelf_life,
                printer_name,
                outer_template,
                inner_template,
                inner_package_name,
                outer_package_name,
                print_row_limits,
            ),
            daemon=True,
        ).start()

    def print_worker(
        self,
        mo_no: str,
        label_types,
        shelf_life: str,
        printer_name: str,
        outer_template: str,
        inner_template: str,
        inner_package_name: str,
        outer_package_name: str,
        print_row_limits,
    ) -> None:
        try:
            self.worker_queue.put((
                "print_result",
                print_labels(
                    mo_no,
                    label_types,
                    printer_name=printer_name,
                    shelf_life=shelf_life,
                    outer_template_override=outer_template,
                    inner_template_override=inner_template,
                    inner_package_name=inner_package_name,
                    outer_package_name=outer_package_name,
                    print_row_limits=print_row_limits,
                ),
            ))
        except Exception as error:
            self.worker_queue.put(("error", error))

    def print_all(self) -> None:
        if not self.result:
            return
        inner = self.result.get("inner") or {}
        self.start_print(["outer", "inner"] if inner.get("can_print") else ["outer"])

    def start_template_mapping_update(self) -> None:
        mapping_file = TEMPLATE_MAPPING_FILE
        selected_dir = filedialog.askdirectory(
            title="选择要扫描并追加到模板映射表的模板目录",
            initialdir=str(TEMPLATE_ROOT),
        )
        if not selected_dir:
            return
        template_root = Path(selected_dir)
        if not messagebox.askyesno(
            APP_TITLE,
            "确认扫描所选目录并追加新模板到当前模板映射表吗？\n\n"
            f"扫描目录：{template_root}\n"
            f"映射表：{mapping_file}\n\n"
            "已有数据不会被修改，只会在表格下方新增未出现过的模板路径。",
        ):
            return

        self.set_busy(True)
        self.status_var.set("正在更新模板映射...")
        self.log(f"开始更新模板映射：{template_root}")
        self.log(f"模板映射写入文件：{mapping_file}")
        threading.Thread(target=self.template_mapping_update_worker, args=(template_root,), daemon=True).start()

    def template_mapping_update_worker(self, template_root: Path) -> None:
        try:
            result = update_template_mapping_from_directory(template_root)
            self.worker_queue.put(("template_mapping_result", result))
        except Exception as error:
            self.worker_queue.put(("error", error))

    def start_template_database_setup(self) -> None:
        selected_dir = filedialog.askdirectory(
            title="选择要设置数据库连接的模板目录",
            initialdir=str(TEMPLATE_ROOT),
        )
        if not selected_dir:
            return
        template_root = Path(selected_dir)
        template_count = len(
            {
                path.resolve()
                for pattern in ("*.btw", "*.btW", "*.BTW")
                for path in template_root.rglob(pattern)
                if path.is_file()
            }
        )
        if template_count <= 0:
            messagebox.showwarning(APP_TITLE, "所选目录下没有找到 .btw 模板文件。")
            return
        try:
            printer_name = self.get_selected_printer()
        except ValueError as error:
            messagebox.showwarning(APP_TITLE, str(error))
            return
        if not printer_name and not messagebox.askyesno(
            APP_TITLE,
            "当前设置为使用模板默认打印机。\n\n"
            "BarTender 打开并保存模板时，可能沿用模板内保存的打印机设置。\n"
            "仍然继续吗？",
        ):
            return

        if not messagebox.askyesno(
            APP_TITLE,
            "确认要批量给所选目录下的 BarTender 模板设置数据库连接吗？\n\n"
            f"目录：{template_root}\n"
            f"模板数量：{template_count}\n\n"
            f"打印机：{printer_name or PRINTER_DEFAULT_TEXT}\n\n"
            "程序会保存模板，并在 template_backups/ 下备份原文件。\n"
            "运行前请关闭 BarTender 中已打开的模板。",
        ):
            return

        self.set_busy(True)
        self.status_var.set("正在批量设置模板数据库连接...")
        self.log(f"开始批量设置模板数据库连接：{template_root}")
        if printer_name:
            self.log(f"模板数据库设置使用打印机：{printer_name}")
        threading.Thread(target=self.template_database_setup_worker, args=(template_root, printer_name), daemon=True).start()

    def template_database_setup_worker(self, template_root: Path, printer_name: str) -> None:
        def progress(index, total, template_path):
            self.worker_queue.put(("template_db_progress", (index, total, template_path)))

        try:
            result = run_batch_database_setup(
                template_root=template_root,
                save=True,
                create_csv=True,
                printer_name=printer_name,
                xmlscript_only=True,
                progress_callback=progress,
            )
            self.worker_queue.put(("template_db_result", result))
        except Exception as error:
            self.worker_queue.put(("error", error))

    def start_environment_check(self) -> None:
        try:
            printer_name = self.get_selected_printer()
        except ValueError as error:
            messagebox.showwarning(APP_TITLE, str(error))
            return
        self.set_busy(True)
        self.status_var.set("正在执行环境自检...")
        self.log("开始环境自检")
        threading.Thread(target=self.environment_check_worker, args=(printer_name,), daemon=True).start()

    def environment_check_worker(self, printer_name: str) -> None:
        try:
            result = self.run_environment_check(printer_name=printer_name)
            self.worker_queue.put(("environment_check_result", result))
        except Exception as error:
            self.worker_queue.put(("error", error))

    def add_check_result(self, rows: list, status: str, item: str, detail: str) -> None:
        rows.append({"status": status, "item": item, "detail": detail})

    def check_write_access(self, rows: list, path: Path, item: str) -> None:
        try:
            path.mkdir(parents=True, exist_ok=True)
            probe = path / f".write_test_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.tmp"
            probe.write_text("ok", encoding="utf-8")
            probe.unlink(missing_ok=True)
            self.add_check_result(rows, "OK", item, f"可写：{path}")
        except Exception as error:
            self.add_check_result(rows, "FAIL", item, f"不可写：{path}；{error}")

    def resolved_bartender_exe(self) -> str:
        path = Path(BARTENDER_EXE)
        if path.exists():
            return str(path)
        found = shutil.which(BARTENDER_EXE)
        return found or BARTENDER_EXE

    def get_file_version(self, path: str) -> str:
        try:
            import win32api

            info = win32api.GetFileVersionInfo(path, "\\")
            ms = info["FileVersionMS"]
            ls = info["FileVersionLS"]
            return ".".join(
                str(part)
                for part in (
                    ms >> 16,
                    ms & 0xFFFF,
                    ls >> 16,
                    ls & 0xFFFF,
                )
            )
        except Exception:
            return ""

    def run_environment_check(self, printer_name: str = "") -> dict:
        rows = []

        self.add_check_result(rows, "OK", "程序目录", str(PROJECT_ROOT))
        self.check_write_access(rows, LOG_DIR, "日志目录写入")
        self.check_write_access(rows, PROJECT_ROOT / "runtime_data", "runtime_data 写入")

        bartender_exe = self.resolved_bartender_exe()
        if Path(bartender_exe).exists():
            version = self.get_file_version(bartender_exe)
            detail = f"{bartender_exe}" + (f"；版本 {version}" if version else "")
            self.add_check_result(rows, "OK", "BarTender", detail)
        else:
            self.add_check_result(rows, "FAIL", "BarTender", f"未找到 bartend.exe：{BARTENDER_EXE}")

        try:
            printers = list_windows_printers()
            if printers:
                detail = f"已安装 {len(printers)} 个打印机"
                if printer_name:
                    detail += f"；当前选择：{printer_name}"
                    if printer_name not in printers:
                        self.add_check_result(rows, "WARN", "打印机", f"{detail}；当前选择不在 Windows 打印机列表中")
                    else:
                        self.add_check_result(rows, "OK", "打印机", detail)
                else:
                    self.add_check_result(rows, "OK", "打印机", detail + "；使用模板默认打印机")
            else:
                self.add_check_result(rows, "WARN", "打印机", "Windows 未返回已安装打印机列表")
        except Exception as error:
            self.add_check_result(rows, "FAIL", "打印机", str(error))

        if TEMPLATE_ROOT.exists() and TEMPLATE_ROOT.is_dir():
            try:
                has_template = any(TEMPLATE_ROOT.rglob("*.btw"))
                detail = f"可访问：{TEMPLATE_ROOT}"
                if not has_template:
                    self.add_check_result(rows, "WARN", "模板目录", detail + "；未发现 .btw 模板")
                else:
                    self.add_check_result(rows, "OK", "模板目录", detail)
            except Exception as error:
                self.add_check_result(rows, "FAIL", "模板目录", f"无法扫描：{TEMPLATE_ROOT}；{error}")
        else:
            self.add_check_result(rows, "FAIL", "模板目录", f"不存在或不可访问：{TEMPLATE_ROOT}")

        if TEMPLATE_MAPPING_FILE.exists() and TEMPLATE_MAPPING_FILE.is_file():
            try:
                from template_mapping_lookup import read_mapping_rows

                mapping_rows = read_mapping_rows(TEMPLATE_MAPPING_FILE)
                self.add_check_result(rows, "OK", "模板映射表", f"可读取 {len(mapping_rows)} 行：{TEMPLATE_MAPPING_FILE}")
            except Exception as error:
                self.add_check_result(rows, "FAIL", "模板映射表", f"读取失败：{TEMPLATE_MAPPING_FILE}；{error}")
        else:
            self.add_check_result(rows, "FAIL", "模板映射表", f"不存在或不可访问：{TEMPLATE_MAPPING_FILE}")

        if PACKAGE_NAME_FILE.exists() and PACKAGE_NAME_FILE.is_file():
            try:
                line_count = len(PACKAGE_NAME_FILE.read_text(encoding="utf-8-sig").splitlines())
                self.add_check_result(rows, "OK", "BOM料号名汇总", f"可读取 {line_count} 行：{PACKAGE_NAME_FILE}")
            except Exception as error:
                self.add_check_result(rows, "FAIL", "BOM料号名汇总", f"读取失败：{PACKAGE_NAME_FILE}；{error}")
        else:
            self.add_check_result(rows, "FAIL", "BOM料号名汇总", f"不存在或不可访问：{PACKAGE_NAME_FILE}")

        try:
            import pyodbc

            drivers = list(pyodbc.drivers())
            if ERP_DRIVER in drivers:
                self.add_check_result(rows, "OK", "ODBC Driver", f"已安装：{ERP_DRIVER}")
            else:
                self.add_check_result(rows, "FAIL", "ODBC Driver", f"未找到 {ERP_DRIVER}；当前驱动：{', '.join(drivers) or '-'}")

            missing = [
                name
                for name, value in (
                    ("ERP_SERVER", ERP_SERVER),
                    ("ERP_DATABASE", ERP_DATABASE),
                    ("ERP_USER", ERP_USER),
                    ("ERP_PASSWORD", ERP_PASSWORD),
                )
                if not value
            ]
            if missing:
                self.add_check_result(rows, "FAIL", "ERP环境变量", "缺少：" + "、".join(missing))
            else:
                self.add_check_result(rows, "OK", "ERP环境变量", f"SERVER={ERP_SERVER}；DATABASE={ERP_DATABASE}；USER={ERP_USER}")
                try:
                    conn = pyodbc.connect(
                        f"DRIVER={{{ERP_DRIVER}}};"
                        f"SERVER={ERP_SERVER};"
                        f"DATABASE={ERP_DATABASE};"
                        f"UID={ERP_USER};"
                        f"PWD={ERP_PASSWORD};"
                        "Encrypt=no;"
                        "TrustServerCertificate=yes;",
                        timeout=3,
                    )
                    conn.close()
                    self.add_check_result(rows, "OK", "ERP连接", "连接成功")
                except Exception as error:
                    self.add_check_result(rows, "FAIL", "ERP连接", str(error))
        except Exception as error:
            self.add_check_result(rows, "FAIL", "pyodbc", f"不可用：{error}")

        try:
            import docs.batch_set_bartender_database as batch_database
            import erp_runtime_csv
            import label_service

            if erp_runtime_csv.LABEL_CSV_FIELDS == label_service.LABEL_CSV_FIELDS == batch_database.CSV_FIELDS:
                self.add_check_result(rows, "OK", "CSV字段表", f"字段数 {len(label_service.LABEL_CSV_FIELDS)}，三处一致")
            else:
                self.add_check_result(rows, "FAIL", "CSV字段表", "erp_runtime_csv / label_service / batch_set_bartender_database 字段不一致")
        except Exception as error:
            self.add_check_result(rows, "FAIL", "CSV字段表", str(error))

        report_path = LOG_DIR / f"environment_check_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        lines = ["环境自检报告", f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ""]
        for row in rows:
            lines.append(f"[{row['status']}] {row['item']}：{row['detail']}")
        report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

        return {
            "rows": rows,
            "report": str(report_path),
            "ok": sum(1 for row in rows if row["status"] == "OK"),
            "warn": sum(1 for row in rows if row["status"] == "WARN"),
            "fail": sum(1 for row in rows if row["status"] == "FAIL"),
        }

    def poll_worker_queue(self) -> None:
        try:
            while True:
                kind, payload = self.worker_queue.get_nowait()
                if kind == "preview_result":
                    self.handle_preview_result(payload)
                elif kind == "print_result":
                    self.handle_print_result(payload)
                elif kind == "template_db_progress":
                    index, total, template_path = payload
                    self.status_var.set(f"正在设置模板数据库连接：{index}/{total}")
                    self.log(f"[{index}/{total}] {template_path}")
                elif kind == "template_db_result":
                    self.handle_template_database_setup_result(payload)
                elif kind == "template_mapping_result":
                    self.handle_template_mapping_update_result(payload)
                elif kind == "environment_check_result":
                    self.handle_environment_check_result(payload)
                elif kind == "error":
                    self.handle_error(payload)
        except queue.Empty:
            pass
        self.after(100, self.poll_worker_queue)

    def handle_preview_result(self, result) -> None:
        self.result = result
        self.current_label_type = "outer"
        self.current_image_index = 0
        self.render_info()
        self.render_preview()
        self.update_action_state()
        self.set_busy(False)
        self.status_var.set("预览生成成功")
        self.log(
            f"外标预览图片：{len(result['outer']['preview_images'])} 页，"
            f"CSV标签：{result['outer'].get('total_label_count', result['outer'].get('label_count', '-'))} 张"
        )
        if result.get("inner"):
            self.log(
                f"内标预览图片：{len(result['inner']['preview_images'])} 页，"
                f"CSV标签：{result['inner'].get('total_label_count', result['inner'].get('label_count', '-'))} 张"
            )
        else:
            self.log("当前工单无内标")

    def handle_print_result(self, result) -> None:
        self.set_busy(False)
        self.update_action_state()
        self.status_var.set("打印命令已执行")
        self.log(f"打印任务数：{len(result.get('printed', []))}")

    def handle_template_database_setup_result(self, result) -> None:
        self.set_busy(False)
        self.update_action_state()
        self.status_var.set("模板数据库设置完成")
        self.log(
            f"模板数据库设置完成：成功 {result.get('success', 0)}，"
            f"失败 {result.get('fail', 0)}，报告 {result.get('report', '')}"
        )
        if result.get("backup_root"):
            self.log(f"模板备份目录：{result.get('backup_root')}")
        messagebox.showinfo(
            APP_TITLE,
            "模板数据库设置完成。\n\n"
            f"成功：{result.get('success', 0)}\n"
            f"失败：{result.get('fail', 0)}\n"
            f"报告：{result.get('report', '')}\n"
            f"备份：{result.get('backup_root', '-') or '-'}",
        )

    def handle_template_mapping_update_result(self, result) -> None:
        self.set_busy(False)
        self.update_action_state()
        self.status_var.set("模板映射更新完成")
        self.log(
            f"模板映射更新完成：扫描 {result.get('scanned', 0)}，"
            f"已有 {result.get('existing', 0)}，新增 {result.get('added', 0)}，"
            f"输出 {result.get('output', '')}"
        )
        messagebox.showinfo(
            APP_TITLE,
            "模板映射更新完成。\n\n"
            f"扫描模板数：{result.get('scanned', 0)}\n"
            f"已有映射数：{result.get('existing', 0)}\n"
            f"本次新增：{result.get('added', 0)}\n"
            f"输出文件：{result.get('output', '')}",
        )

    def handle_environment_check_result(self, result) -> None:
        self.set_busy(False)
        self.update_action_state()
        fail = result.get("fail", 0)
        warn = result.get("warn", 0)
        ok = result.get("ok", 0)
        self.status_var.set(f"环境自检完成：OK {ok}，WARN {warn}，FAIL {fail}")
        self.log(f"环境自检完成：OK {ok}，WARN {warn}，FAIL {fail}，报告 {result.get('report', '')}")

        dialog = tk.Toplevel(self)
        dialog.title("环境自检报告")
        dialog.transient(self)
        dialog.geometry("880x560")
        dialog.minsize(760, 420)

        frame = ttk.Frame(dialog, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        dialog.rowconfigure(0, weight=1)
        dialog.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

        summary = ttk.Label(
            frame,
            text=f"OK {ok}    WARN {warn}    FAIL {fail}\n报告文件：{result.get('report', '')}",
            anchor="w",
            justify="left",
        )
        summary.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        text = tk.Text(frame, wrap="word")
        text.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        text.configure(yscrollcommand=scrollbar.set)
        text.tag_configure("OK", foreground="#167a2f")
        text.tag_configure("WARN", foreground="#9a6700")
        text.tag_configure("FAIL", foreground="#b00020")

        for row in result.get("rows", []):
            status = row.get("status", "")
            line = f"[{status}] {row.get('item', '')}：{row.get('detail', '')}\n"
            text.insert("end", line, status if status in ("OK", "WARN", "FAIL") else "")
        text.configure(state="disabled")

        ttk.Button(frame, text="关闭", command=dialog.destroy).grid(row=2, column=0, columnspan=2, sticky="e", pady=(10, 0))

    def handle_error(self, error) -> None:
        self.set_busy(False)
        self.update_action_state()
        self.status_var.set("操作失败")
        self.log(f"失败：{error}")
        messagebox.showerror(APP_TITLE, str(error))

    def render_info(self) -> None:
        result = self.result
        outer = result.get("outer") or {}
        inner = result.get("inner") or {}
        remainder_notices = result.get("remainder_notices") or []
        lines = [
            f"MO号：{result.get('mo_no', '')}",
            f"客户编号：{result.get('customer_code', '')}",
            f"客户料号：{result.get('customer_part_no', '')}",
            f"产品编号：{result.get('product_code', '')}",
            f"生产日期：{result.get('mfg_date', '')}",
            f"有效期：{result.get('exp_date', '')}",
            f"保质期：{result.get('shelf_life', '-') or '-'}",
            f"保质期来源：{result.get('shelf_life_source', '-') or '-'}",
            f"打印机：{result.get('printer_name', '') or PRINTER_DEFAULT_TEXT}",
            f"UI指定外标包装：{result.get('outer_package_override', '') or '-'}",
            f"UI指定内标包装：{result.get('inner_package_override', '') or '-'}",
            f"是否有内标：{'是' if result.get('has_inner_label') else '否'}",
            "",
            f"外标数量：{outer.get('qty', '-')}",
            f"外标整标张数：{outer.get('label_count', '-')}",
            f"外标总张数：{outer.get('total_label_count', outer.get('label_count', '-'))}",
            f"外标余数：{outer.get('remainder_qty', '-') or '-'}",
            f"外标零数标签QTY：{outer.get('zero_label_qty', '-') or '-'}",
            f"外标包材编号：{outer.get('package_prd_no', '-') or '-'}",
            f"外标包材名称：{outer.get('package_name', '-') or '-'}",
            "",
            f"内标数量：{inner.get('qty', '-') if inner else '-'}",
            f"内标整标张数：{inner.get('label_count', '-') if inner else '-'}",
            f"内标总张数：{inner.get('total_label_count', inner.get('label_count', '-')) if inner else '-'}",
            f"内标余数：{inner.get('remainder_qty', '-') if inner else '-'}",
            f"内标零数标签QTY：{inner.get('zero_label_qty', '-') if inner else '-'}",
            f"内标包材编号：{inner.get('package_prd_no', '-') if inner else '-'}",
            f"内标包材名称：{inner.get('package_name', '-') if inner else '-'}",
            "",
            f"外标模板：{outer.get('template', '-')}",
            f"内标模板：{inner.get('template', '-') if inner else '-'}",
        ]
        self.info_text.configure(state="normal")
        self.info_text.delete("1.0", "end")
        saved_sections = []
        if remainder_notices:
            warning_text = "零数提醒：\n" + "\n".join(f"- {notice}" for notice in remainder_notices) + "\n\n"
            self.info_text.insert("end", warning_text, "warning")
            saved_sections.append(warning_text.rstrip())
        notices = result.get("notices") or []
        if notices:
            notice_text = "提示：\n" + "\n".join(f"- {notice}" for notice in notices) + "\n\n"
            self.info_text.insert("end", notice_text, "warning")
            saved_sections.append(notice_text.rstrip())
        info_text = "\n".join(lines)
        self.info_text.insert("end", info_text)
        self.info_text.configure(state="disabled")
        saved_sections.append(info_text)
        self.save_work_order_info(result, "\n\n".join(saved_sections))

    def save_work_order_info(self, result, content: str) -> None:
        mo_no = safe_log_segment(str(result.get("mo_no", "")))
        path = LOG_DIR / f"work_order_{mo_no}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        path.write_text(content + "\n", encoding="utf-8")
        self.log(f"工单信息已保存：{path}")

    def current_images(self):
        if not self.result:
            return []
        label = self.result.get(self.current_label_type)
        if not label:
            return []
        return label.get("preview_images") or []

    def render_preview(self) -> None:
        images = self.current_images()
        if not images:
            self.current_photo = None
            self.preview_label.configure(image="", text="当前标签无预览")
            self.counter_var.set("0 / 0 页")
            self.update_nav_state()
            return

        self.current_image_index = max(0, min(self.current_image_index, len(images) - 1))
        path = url_to_file_path(images[self.current_image_index])
        photo = tk.PhotoImage(file=str(path))
        photo = self.fit_photo(photo, max_width=760, max_height=560)
        self.current_photo = photo
        self.preview_label.configure(image=photo, text="")
        self.counter_var.set(f"{self.current_image_index + 1} / {len(images)} 页")
        self.update_nav_state()

    def fit_photo(self, photo: tk.PhotoImage, max_width: int, max_height: int) -> tk.PhotoImage:
        x_factor = max(1, (photo.width() + max_width - 1) // max_width)
        y_factor = max(1, (photo.height() + max_height - 1) // max_height)
        factor = max(x_factor, y_factor)
        if factor > 1:
            return photo.subsample(factor, factor)
        return photo

    def switch_label_type(self, label_type: str) -> None:
        self.current_label_type = label_type
        self.current_image_index = 0
        self.render_preview()

    def prev_image(self) -> None:
        self.current_image_index -= 1
        self.render_preview()

    def next_image(self) -> None:
        self.current_image_index += 1
        self.render_preview()

    def update_nav_state(self) -> None:
        images = self.current_images()
        self.prev_button.configure(state="normal" if self.current_image_index > 0 else "disabled")
        self.next_button.configure(state="normal" if self.current_image_index < len(images) - 1 else "disabled")

    def update_action_state(self) -> None:
        has_result = bool(self.result)
        inner = self.result.get("inner") if self.result else None
        has_inner = bool(inner and inner.get("can_print"))
        self.print_outer_button.configure(state="normal" if has_result else "disabled")
        self.print_inner_button.configure(state="normal" if has_inner else "disabled")
        self.print_all_button.configure(state="normal" if has_result else "disabled")


def main() -> int:
    app = LabelGeneratorApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
