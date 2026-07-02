from __future__ import annotations

import json
import sys
from pathlib import Path


def get_project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


PROJECT_ROOT = get_project_root()
APP_SETTINGS_FILE = PROJECT_ROOT / "app_settings.json"
RUNTIME_DIR = PROJECT_ROOT / "runtime_data"
FINAL_LABELS_DIR = PROJECT_ROOT / "final_labels"
LOG_DIR = PROJECT_ROOT / "logs"
DEFAULT_TEMPLATE_ROOT = PROJECT_ROOT / "Templates"
DEFAULT_PACKAGE_NAME_FILE = PROJECT_ROOT / "bom料号名汇总.txt"
DEFAULT_TEMPLATE_MAPPING_FILE = PROJECT_ROOT / "template_mapping.xlsx"


def load_app_settings() -> dict:
    if not APP_SETTINGS_FILE.exists():
        return {}
    try:
        with APP_SETTINGS_FILE.open("r", encoding="utf-8") as file:
            data = json.load(file)
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def resolve_setting_path(value: str, default: Path) -> Path:
    text = str(value or "").strip()
    if not text:
        return default
    path = Path(text)
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return path.resolve()


def serialize_setting_path(path: Path) -> str:
    resolved = Path(path).resolve()
    try:
        return str(resolved.relative_to(PROJECT_ROOT))
    except ValueError:
        return str(resolved)


def _configured_path(key: str, default: Path) -> Path:
    return resolve_setting_path(load_app_settings().get(key, ""), default)


TEMPLATE_ROOT = _configured_path("template_root", DEFAULT_TEMPLATE_ROOT)
PACKAGE_NAME_FILE = _configured_path("package_name_file", DEFAULT_PACKAGE_NAME_FILE)
TEMPLATE_MAPPING_FILE = _configured_path("template_mapping_file", DEFAULT_TEMPLATE_MAPPING_FILE)


def save_path_settings(template_root: Path, package_name_file: Path, template_mapping_file: Path) -> None:
    settings = load_app_settings()
    settings["template_root"] = serialize_setting_path(template_root)
    settings["package_name_file"] = serialize_setting_path(package_name_file)
    settings["template_mapping_file"] = serialize_setting_path(template_mapping_file)
    with APP_SETTINGS_FILE.open("w", encoding="utf-8") as file:
        json.dump(settings, file, ensure_ascii=False, indent=2)

    global TEMPLATE_ROOT, PACKAGE_NAME_FILE, TEMPLATE_MAPPING_FILE
    TEMPLATE_ROOT = Path(template_root).resolve()
    PACKAGE_NAME_FILE = Path(package_name_file).resolve()
    TEMPLATE_MAPPING_FILE = Path(template_mapping_file).resolve()
