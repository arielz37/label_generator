import os
from pathlib import Path


def find_bartender_exe() -> str:
    configured = os.getenv("BARTENDER_EXE", "").strip()
    if configured:
        return configured

    candidates = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Seagull" / "BarTender Suite" / "bartend.exe",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Seagull" / "BarTender Suite" / "bartend.exe",
    ]
    for base in (
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Seagull",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Seagull",
    ):
        if base.exists():
            candidates.extend(base.glob("BarTender*/bartend.exe"))

    for candidate in candidates:
        if candidate.exists():
            return str(candidate)

    return "bartend.exe"


ERP_DRIVER = os.getenv("ERP_DRIVER", "ODBC Driver 17 for SQL Server")
ERP_SERVER = os.getenv("ERP_SERVER", "")
ERP_DATABASE = os.getenv("ERP_DATABASE", "DB_BD01")
ERP_USER = os.getenv("ERP_USER", "")
ERP_PASSWORD = os.getenv("ERP_PASSWORD", "")
BARTENDER_EXE = find_bartender_exe()
BARTENDER_PRINTER = os.getenv("BARTENDER_PRINTER", "").strip()


def require_erp_config():
    missing = [
        name
        for name, value in (
            ("ERP_SERVER", ERP_SERVER),
            ("ERP_USER", ERP_USER),
            ("ERP_PASSWORD", ERP_PASSWORD),
        )
        if not value
    ]
    if missing:
        raise RuntimeError("Missing ERP config environment variables: " + ", ".join(missing))
