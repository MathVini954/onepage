from __future__ import annotations


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
