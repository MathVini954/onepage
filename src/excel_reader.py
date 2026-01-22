from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet
import unicodedata
import re


# ============================================================
# Normalização / parsing
# ============================================================
_PT_MONTH = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
}

def _strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def _norm(x: Any) -> str:
    if x is None:
        return ""
    s = str(x).strip().replace("\u00a0", " ")
    s = " ".join(s.split())
    s = _strip_accents(s).upper()
    return s

def _is_blank(x: Any) -> bool:
    if x is None:
        return True
    if isinstance(x, str) and not x.strip():
        return True
    return False

def _to_float(x: Any) -> float | None:
    if x is None:
        return None
    try:
        return float(x)
    except Exception:
        try:
            s = str(x).strip()
            if not s:
                return None
            s = s.replace(".", "").replace(",", ".")
            return float(s)
        except Exception:
            return None

def _to_month(x: Any) -> pd.Timestamp | None:
    """
    Converte:
      - datetime/date
      - strings: 'Jan.26', 'JAN/26', 'jan/2026', '01/01/2026'
    Sempre retorna Timestamp (1º dia do mês) ou None. Nunca lança erro.
    """
    try:
        if x is None:
            return None

        # datetime/date do Excel
        if hasattr(x, "year") and hasattr(x, "month") and hasattr(x, "day"):
            try:
                dt = pd.Timestamp(x)
                if pd.isna(dt):
                    return None
                return pd.Timestamp(year=dt.year, month=dt.month, day=1)
            except Exception:
                return None

        # string
        if isinstance(x, str):
            s = x.strip()
            if not s:
                return None
            up = _norm(s)

            # "JAN.26" / "JAN/26" / "JAN-2026"
            m = re.match(r"^([A-Z]{3})[./\- ]*(\d{2,4})$", up)
            if m:
                mon = _PT_MONTH.get(m.group(1))
                yy = int(m.group(2))
                if mon:
                    year = yy if yy >= 100 else (2000 + yy)
                    return pd.Timestamp(year=year, month=mon, day=1)

            # parse padrão
            dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
            if pd.notna(dt):
                dt = pd.Timestamp(dt)
                return pd.Timestamp(year=dt.year, month=dt.month, day=1)
            return None

        # fallback (números etc.)
        dt = pd.to_datetime(x, errors="coerce")
        if pd.notna(dt):
            dt = pd.Timestamp(dt)
            return pd.Timestamp(year=dt.year, month=dt.month, day=1)

        return None
    except Exception:
        return None


def _find_header_row(ws: Worksheet, must_contain_any: list[str], max_scan: int = 500) -> int | None:
    """
    Procura uma linha onde pelo menos os termos em must_contain_any apareçam na linha.
    """
    max_r = min(ws.max_row or 1, max_scan)
    for r in range(1, max_r + 1):
        row_txt = " | ".join(_norm(ws.cell(r, c).value) for c in range(1, 15))
        ok = True
        for token in must_contain_any:
            if token not in row_txt:
                ok = False
                break
        if ok:
            return r
    return None

def _map_cols(ws: Worksheet, header_row: int, max_col: int = 20) -> dict[str, int]:
    """
    Mapeia cabeçalhos (normalizados) -> coluna (1-index).
    """
    m: dict[str, int] = {}
    for c in range(1, max_col + 1):
        h = _norm(ws.cell(header_row, c).value)
        if h:
            m[h] = c
    return m

def _find_col(cols: dict[str, int], *keywords: str) -> int | None:
    """
    Retorna a primeira coluna cujo header contém TODAS as keywords.
    """
    ks = [_norm(k) for k in keywords]
    for h, c in cols.items():
        if all(k in h for k in ks):
            return c
    return None


# ============================================================
# API usada no app.py
# ============================================================
def load_wb(path: str | Path):
    path = Path(path)
    keep = path.suffix.lower() == ".xlsm"
    return openpyxl.load_workbook(path, data_only=True, keep_vba=keep)

def sheetnames(wb) -> list[str]:
    out: list[str] = []
    for name in wb.sheetnames:
        n = _norm(name)
        if n in {"LEIA-ME", "LEIA ME", "README"}:
            continue
        if n.startswith("_"):
            continue
        out.append(name)
    return out


# ============================================================
# Leitura dos blocos
# ============================================================
def read_resumo_financeiro(ws: Worksheet) -> dict[str, float | None]:
    wanted = {
        "ORÇAMENTO INICIAL (R$)",
        "ORÇAMENTO REAJUSTADO (R$)",
        "DESEMBOLSO ACUMULADO (R$)",
        "A PAGAR (R$)",
        "SALDO A INCORRER (R$)",
        "CUSTO FINAL (R$)",
        "VARIAÇÃO (R$)",
    }
    out: dict[str, float | None] = {}
    for r in range(1, min(ws.max_row or 1, 100) + 1):
        label = ws.cell(r, 1).value
        if not isinstance(label, str):
            continue
        lab = str(label).strip()
        if lab in wanted:
            out[lab] = _to_float(ws.cell(r, 2).value)
    return out


def read_indice(ws: Worksheet) -> pd.DataFrame:
    header_row = _find_header_row(ws, ["MES", "INDICE PROJETADO"], max_scan=250)
    if header_row is None:
        return pd.DataFrame(columns=["MÊS", "ÍNDICE PROJETADO"])

    cols = _map_cols(ws, header_row, max_col=10)
    c_mes = _find_col(cols, "MES") or 1
    c_idx = _find_col(cols, "INDICE", "PROJETADO") or 2

    rows = []
    blank_run = 0
    for r in range(header_row + 1, (ws.max_row or header_row + 1) + 1):
        mes = ws.cell(r, c_mes).value
        idx = ws.cell(r, c_idx).value

        if _is_blank(mes) and _is_blank(idx):
            blank_run += 1
            if blank_run >= 6:
                break
            continue
        blank_run = 0

        m = _to_month(mes)
        v = _to_float(idx)
        if m is None or v is None:
            continue
        rows.append((m, v))

    df = pd.DataFrame(rows, columns=["MÊS", "ÍNDICE PROJETADO"])
    if not df.empty:
        df = df.sort_values("MÊS")
    return df


def read_financeiro(ws: Worksheet) -> pd.DataFrame:
    header_row = _find_header_row(ws, ["DESEMBOLSO DO MES", "MEDIDO NO MES"], max_scan=300)
    if header_row is None:
        return pd.DataFrame(columns=["MÊS", "DESEMBOLSO DO MÊS (R$)", "MEDIDO NO MÊS (R$)"])

    cols = _map_cols(ws, header_row, max_col=15)
    c_mes = _find_col(cols, "MES") or 4
    c_des = _find_col(cols, "DESEMBOLSO", "MES") or 5
    c_med = _find_col(cols, "MEDIDO", "MES") or 6

    rows = []
    blank_run = 0
    for r in range(header_row + 1, (ws.max_row or header_row + 1) + 1):
        mes = ws.cell(r, c_mes).value
        des = ws.cell(r, c_des).value
        med = ws.cell(r, c_med).value

        if _is_blank(mes) and _is_blank(des) and _is_blank(med):
            blank_run += 1
            if blank_run >= 6:
                break
            continue
        blank_run = 0

        m = _to_month(mes)
        if m is None:
            continue
        rows.append((m, _to_float(des), _to_float(med)))

    df = pd.DataFrame(rows, columns=["MÊS", "DESEMBOLSO DO MÊS (R$)", "MEDIDO NO MÊS (R$)"])
    if not df.empty:
        df = df.sort_values("MÊS")
    return df


def read_prazo(ws: Worksheet) -> pd.DataFrame:
    """
    Lê PRAZO independente de posição de coluna:
    precisa ter 'MES', 'PLANEJADO ACUM', 'PLANEJADO MES', 'REALIZADO', e 'PREVISTO MENSAL(%)' (ou variações).
    """
    header_row = _find_header_row(ws, ["MES", "PLANEJADO", "REALIZADO"], max_scan=600)
    if header_row is None:
        return pd.DataFrame(
            columns=["MÊS", "PLANEJADO ACUM. (%)", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)", "PREVISTO MENSAL (%)"]
        )

    cols = _map_cols(ws, header_row, max_col=20)

    c_mes  = _find_col(cols, "MES") or 1
    c_pa   = _find_col(cols, "PLANEJADO", "ACUM")  # planejado acumulado
    c_pm   = _find_col(cols, "PLANEJADO", "MES")   # planejado mês
    c_rm   = _find_col(cols, "REALIZADO")          # realizado mês
    c_prev = _find_col(cols, "PREVISTO", "MENSAL") or _find_col(cols, "PREVISTO")  # previsto mensal(%) var

    rows = []
    blank_run = 0
    for r in range(header_row + 1, (ws.max_row or header_row + 1) + 1):
        mes = ws.cell(r, c_mes).value if c_mes else None
        pa  = ws.cell(r, c_pa).value if c_pa else None
        pm  = ws.cell(r, c_pm).value if c_pm else None
        rm  = ws.cell(r, c_rm).value if c_rm else None
        pv  = ws.cell(r, c_prev).value if c_prev else None

        if _is_blank(mes) and _is_blank(pa) and _is_blank(pm) and _is_blank(rm) and _is_blank(pv):
            blank_run += 1
            if blank_run >= 8:
                break
            continue
        blank_run = 0

        m = _to_month(mes)
        if m is None:
            continue

        rows.append((m, _to_float(pa), _to_float(pm), _to_float(rm), _to_float(pv)))

    df = pd.DataFrame(
        rows,
        columns=["MÊS", "PLANEJADO ACUM. (%)", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)", "PREVISTO MENSAL (%)"],
    )
    if not df.empty:
        df = df.sort_values("MÊS")
    return df


def read_acrescimos_economias(ws: Worksheet) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Lê as duas listas (Acréscimos e Economias) incluindo JUSTIFICATIVAS.
    Assume layout padrão:
      Acréscimos: A..F (F = JUSTIFICATIVAS)
      Economias : G..L (L = JUSTIFICATIVAS)
    """
    header_row = _find_header_row(ws, ["ECONOMIAS", "ACRESCIMOS"], max_scan=900)
    if header_row is None:
        cols = ["DESCRIÇÃO", "ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO", "JUSTIFICATIVAS"]
        return pd.DataFrame(columns=cols), pd.DataFrame(columns=cols)

    def read_side(start_col: int) -> pd.DataFrame:
        rows = []
        blank_run = 0
        for r in range(header_row + 2, (ws.max_row or header_row + 2) + 1):
            desc = ws.cell(r, start_col + 0).value
            o_ini = ws.cell(r, start_col + 1).value
            o_rea = ws.cell(r, start_col + 2).value
            c_fin = ws.cell(r, start_col + 3).value
            var_  = ws.cell(r, start_col + 4).value
            just  = ws.cell(r, start_col + 5).value

            if _is_blank(desc) and _is_blank(o_ini) and _is_blank(o_rea) and _is_blank(c_fin) and _is_blank(var_) and _is_blank(just):
                blank_run += 1
                if blank_run >= 10:
                    break
                continue
            blank_run = 0

            if _is_blank(desc):
                continue

            rows.append(
                (
                    str(desc).strip(),
                    _to_float(o_ini),
                    _to_float(o_rea),
                    _to_float(c_fin),
                    _to_float(var_),
                    "" if just is None else str(just).strip(),
                )
            )

        return pd.DataFrame(
            rows,
            columns=["DESCRIÇÃO", "ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO", "JUSTIFICATIVAS"],
        )

    df_acres = read_side(1)   # A..F
    df_econ  = read_side(7)   # G..L
    return df_acres, df_econ
