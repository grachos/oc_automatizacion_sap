"""Loads settings from .env."""
import os
import sys
from dataclasses import dataclass, field
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
load_dotenv(APP_DIR / ".env")  # .env lives next to the .exe (or next to the scripts)


def _env(key: str, default: str = "") -> str:
    return os.getenv(key, default).strip()


@dataclass(frozen=True)
class SapSettings:
    logon_path: str = field(default_factory=lambda: _env("SAP_LOGON_PATH", r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"))
    connection: str = field(default_factory=lambda: _env("SAP_CONNECTION"))
    client: str = field(default_factory=lambda: _env("SAP_CLIENT"))
    user: str = field(default_factory=lambda: _env("SAP_USER"))
    password: str = field(default_factory=lambda: _env("SAP_PASSWORD"))
    language: str = field(default_factory=lambda: _env("SAP_LANGUAGE", "ES"))


@dataclass(frozen=True)
class PoDefaults:
    purch_org: str = field(default_factory=lambda: _env("PO_PURCH_ORG", "DA00"))
    purch_group: str = field(default_factory=lambda: _env("PO_PURCH_GROUP", "025"))
    company_code: str = field(default_factory=lambda: _env("PO_COMPANY_CODE", "DA25"))
    payment_days: str = field(default_factory=lambda: _env("PO_PAYMENT_DAYS", "7"))  # "Pago en" (Cond.pago stays blank)
    acct_assignment: str = field(default_factory=lambda: _env("PO_ACCT_ASSIGNMENT", "1"))
    item_category: str = field(default_factory=lambda: _env("PO_ITEM_CATEGORY", "F"))
    short_text: str = field(default_factory=lambda: _env("PO_SHORT_TEXT", "TRANSPORTE"))
    quantity: str = field(default_factory=lambda: _env("PO_QUANTITY", "1"))
    unit: str = field(default_factory=lambda: _env("PO_UNIT", "UP"))
    material_group: str = field(default_factory=lambda: _env("PO_MATERIAL_GROUP", "TRANSPORT"))
    plant: str = field(default_factory=lambda: _env("PO_PLANT", "2500"))
    tax_code: str = field(default_factory=lambda: _env("PO_TAX_CODE", "VC"))
    service_number: str = field(default_factory=lambda: _env("PO_SERVICE_NUMBER", "4220"))
    service_qty: str = field(default_factory=lambda: _env("PO_SERVICE_QTY", "1"))
    gl_account: str = field(default_factory=lambda: _env("PO_GL_ACCOUNT", "2815052000"))


def save_env(values: dict) -> None:
    """Writes settings to the .env next to the app and reloads them (no restart needed)."""
    from dotenv import set_key
    env_file = APP_DIR / ".env"
    env_file.touch(exist_ok=True)
    for k, v in values.items():
        set_key(str(env_file), k, v.strip(), quote_mode="always")
    load_dotenv(env_file, override=True)


def validate() -> list[str]:
    """Returns a list of missing required settings."""
    s, p = SapSettings(), PoDefaults()
    missing = [k for k, v in {
        "SAP_CONNECTION": s.connection, "SAP_CLIENT": s.client,
        "SAP_USER": s.user, "SAP_PASSWORD": s.password, "PO_PLANT": p.plant,
    }.items() if not v]
    if not Path(s.logon_path).exists():
        missing.append(f"SAP_LOGON_PATH (no existe: {s.logon_path})")
    return missing
