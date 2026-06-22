import os


ERP_DRIVER = os.getenv("ERP_DRIVER", "ODBC Driver 17 for SQL Server")
ERP_SERVER = os.getenv("ERP_SERVER", "")
ERP_DATABASE = os.getenv("ERP_DATABASE", "DB_BD01")
ERP_USER = os.getenv("ERP_USER", "")
ERP_PASSWORD = os.getenv("ERP_PASSWORD", "")


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
