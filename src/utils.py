# src/utils.py
from __future__ import annotations
from datetime import datetime, date
import math
import pandas as pd

PT_MONTHS = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
}

def norm(s) -> str:
    if s is None:
        return ""
    return str(s).strip().upper()

def is_blank(x) -> bool:
    if x is None:
        return True
    if isinstance(x, str) and x.strip() == "":
        return True
    return False

def to_month(v) -> date | None:
    """Converte 'MÊS' para date (1º dia do mês) aceitando date/datetime/strings como 'jan/2026'."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return date(v.year, v.month, 1)
    if isinstance(v, date):
        return date(v.year, v.month, 1)

    s = str(v).strip()
    if not s:
        return None

    # tenta pandas direto
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return date(int(dt.year), int(dt.month), 1)
    except Exception:
        pass

    # tenta 'jan/2026', 'Jan.26', 'jan.26'
    s2 = s.replace(".", "/").replace("-", "/")
    parts = [p for p in s2.split("/") if p]
    if len(parts) == 2:
        m_raw, y_raw = parts[0].strip(), parts[1].strip()
        m_key = norm(m_raw)[:3]
        if m_key in PT_MONTHS:
            m = PT_MONTHS[m_key]
            # ano 2 dígitos
            if len(y_raw) == 2 and y_raw.isdigit():
                y = 2000 + int(y_raw)
            else:
                y = int(y_raw)
            return date(y, m, 1)

    return None

def to_float(v) -> float | None:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return None
        return float(v)
    s = str(v).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

def fmt_brl(v) -> str:
    if v is None:
        return "—"
    try:
        return f"R$ {float(v):,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"

def fmt_pct(v, scale_0_1=True) -> str:
    if v is None:
        return "—"
    try:
        x = float(v)
        if scale_0_1:
            x *= 100
        return f"{x:.1f}%".replace(".", ",")
    except Exception:
        return "—"
