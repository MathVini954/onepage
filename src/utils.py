# src/utils.py
from __future__ import annotations

from datetime import datetime
import re
import pandas as pd


def norm(v) -> str:
    """Normaliza texto para comparação (uppercase, sem espaços extras)."""
    if v is None:
        return ""
    return str(v).strip().upper()


def is_blank(v) -> bool:
    """Vazio para célula Excel."""
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def to_float(v):
    """Converte para float (ou None)."""
    if v is None:
        return None
    try:
        if isinstance(v, str):
            s = v.strip()
            if s == "":
                return None
            # troca milhar/decimal BR se vier como string
            s = s.replace("R$", "").strip()
            s = s.replace(".", "").replace(",", ".")
            return float(s)
        return float(v)
    except Exception:
        return None


def to_month(v):
    """
    Converte mês vindo do Excel:
    - datetime / date / Timestamp
    - string tipo 'jan/2026', 'Jan.26', '01/01/2026'
    Retorna Timestamp (ou None).
    """
    if v is None:
        return None

    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()

    if isinstance(v, datetime):
        return v

    # Excel pode mandar date sem hora
    try:
        dt = pd.to_datetime(v, errors="coerce")
        if pd.notna(dt):
            return dt.to_pydatetime()
    except Exception:
        pass

    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None

        # tenta converter direto
        dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.notna(dt):
            return dt.to_pydatetime()

        # tenta "Jan.26" / "Jan/26" etc
        m = re.match(r"^([A-Za-zÀ-ÿ]{3,})\.?[/\- ](\d{2,4})$", s)
        if m:
            mon = m.group(1).lower()
            year = int(m.group(2))
            year = 2000 + year if year < 100 else year
            meses = {
                "jan": 1, "fev": 2, "feb": 2, "mar": 3, "abr": 4, "apr": 4,
                "mai": 5, "may": 5, "jun": 6, "jul": 7, "ago": 8,
                "set": 9, "sep": 9, "out": 10, "oct": 10, "nov": 11, "dez": 12, "dec": 12
            }
            mon3 = mon[:3]
            if mon3 in meses:
                return datetime(year, meses[mon3], 1)

    return None


def fmt_brl(v) -> str:
    if v is None:
        return "—"
    try:
        n = float(v)
    except Exception:
        return "—"
    s = f"{n:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"
