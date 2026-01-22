from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet


# ============================================================
# Helpers (strings / blanks / parsing)
# ============================================================
_PT_MONTH = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
}

def _norm(s: Any) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = s.replace("\u00a0", " ")
    return " ".join(s.split()).upper()

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
            # "1.234.567,89" -> 1234567.89
            s = s.replace(".", "").replace(",", ".")
            return float(s)
        except Exception:
            return None

def _to_month(x: Any) -> pd.Timestamp | None:
    """Converte datas do Excel / strings tipo 'Jan.26' / 'jan/2026' / '01/01/2026' em Timestamp."""
    if x is None or (isinstance(x, str) and not x.strip()):
        return None

    # Datas já como datetime/date
    if hasattr(x, "year") and hasattr(x, "month") and hasattr(x, "day"):
        try:
            return pd.Timestamp(x).normalize()
        except Exception:
            pass

    # Strings
    if isinstance(x, str):
        s = x.strip()
        up = _norm(s)

        # "JAN.26", "JAN/26", "JAN-2026"
        import re
        m = re.match(r"^([A-ZÇ]{3})[./\- ]*(\d{2,4})$", up)
        if m:
            mon = _PT_MONTH.get(m.group(1), None)
            yy = int(m.group(2))
            if mon:
                year = yy if yy >= 100 else (2000 + yy)
                return pd.Timestamp(year=year, month=mon, day=1)

        # fallback: tenta parse padrão
        try:
            dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
            if pd.notna(dt):
                return pd.Timestamp(dt).normalize()
        except Exception:
            pass

    # Números (serial excel) - raro via openpyxl, mas garante
    try:
        dt = pd.to_datetime(x, errors="coerce")
        if pd.notna(dt):
            return pd.Timestamp(dt).normalize()
    except Exception:
        pass

    return None


def _find_row_with(ws: Worksheet, must_have: list[tuple[int, str]], max_scan: int = 400) -> int | None:
    """
    Procura uma linha onde cada (col, texto) está presente (match por contains).
    col é 1-indexado.
    """
    max_r = min(ws.max_row or 1, max_scan)
    for r in range(1, max_r + 1):
        ok = True
        for col, token in must_have:
            v = ws.cell(r, col).value
            if token not in _norm(v):
                ok = False
                break
        if ok:
            return r
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
    """
    Lê o bloco A2:B8 (labels na col A, valores na col B).
    Retorna dict com as chaves exatamente como estão no Excel (ex: 'ORÇAMENTO INICIAL (R$)').
    """
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
    # varre um pedaço do topo (robusto se mexer em linhas)
    for r in range(1, min(ws.max_row or 1, 80) + 1):
        label = ws.cell(r, 1).value
        if not isinstance(label, str):
            continue
        lab = str(label).strip()
        if lab in wanted:
            out[lab] = _to_float(ws.cell(r, 2).value)
    return out


def read_indice(ws: Worksheet) -> pd.DataFrame:
    """
    Tabela: col A = MÊS, col B = ÍNDICE PROJETADO (header típico na linha 11)
    """
    header_row = _find_row_with(ws, [(1, "MÊS"), (2, "ÍNDICE PROJETADO")], max_scan=200)
    if header_row is None:
        return pd.DataFrame(columns=["MÊS", "ÍNDICE PROJETADO"])

    rows = []
    blank_run = 0
    for r in range(header_row + 1, (ws.max_row or header_row + 1) + 1):
        mes = ws.cell(r, 1).value
        idx = ws.cell(r, 2).value

        if _is_blank(mes) and _is_blank(idx):
            blank_run += 1
            if blank_run >= 5:
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
    """
    Tabela: col D = MÊS, col E = DESEMBOLSO DO MÊS (R$), col F = MEDIDO NO MÊS (R$)
    """
    header_row = _find_row_with(ws, [(5, "DESEMBOLSO DO MÊS"), (6, "MEDIDO NO MÊS")], max_scan=250)
    if header_row is None:
        return pd.DataFrame(columns=["MÊS", "DESEMBOLSO DO MÊS (R$)", "MEDIDO NO MÊS (R$)"])

    rows = []
    blank_run = 0
    for r in range(header_row + 1, (ws.max_row or header_row + 1) + 1):
        mes = ws.cell(r, 4).value
        des = ws.cell(r, 5).value
        med = ws.cell(r, 6).value

        if _is_blank(mes) and _is_blank(des) and _is_blank(med):
            blank_run += 1
            if blank_run >= 5:
                break
            continue
        blank_run = 0

        m = _to_month(mes)
        vdes = _to_float(des)
        vmed = _to_float(med)
        if m is None:
            continue

        rows.append((m, vdes, vmed))

    df = pd.DataFrame(rows, columns=["MÊS", "DESEMBOLSO DO MÊS (R$)", "MEDIDO NO MÊS (R$)"])
    if not df.empty:
        df = df.sort_values("MÊS")
    return df


def read_prazo(ws: Worksheet) -> pd.DataFrame:
    """
    Tabela (linha ~27):
      A: MÊS
      B: PLANEJADO ACUM. (%)
      C: PLANEJADO MÊS (%)
      D: REALIZADO Mês (%)
      E: PREVISTO MENSAL(%)  -> padroniza para 'PREVISTO MENSAL (%)'
    """
    header_row = _find_row_with(
        ws,
        [(1, "MÊS"), (2, "PLANEJADO ACUM"), (3, "PLANEJADO MÊS"), (4, "REALIZADO")],
        max_scan=400,
    )
    if header_row is None:
        return pd.DataFrame(
            columns=["MÊS", "PLANEJADO ACUM. (%)", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)", "PREVISTO MENSAL (%)"]
        )

    # descobre qual texto está no header da coluna E e normaliza para o nome que o app usa
    hE = _norm(ws.cell(header_row, 5).value)
    previsto_name = "PREVISTO MENSAL (%)" if "PREVISTO" in hE else "PREVISTO MENSAL (%)"

    rows = []
    blank_run = 0
    for r in range(header_row + 1, (ws.max_row or header_row + 1) + 1):
        mes = ws.cell(r, 1).value
        pla_a = ws.cell(r, 2).value
        pla_m = ws.cell(r, 3).value
        rea_m = ws.cell(r, 4).value
        prev_m = ws.cell(r, 5).value

        if _is_blank(mes) and _is_blank(pla_a) and _is_blank(pla_m) and _is_blank(rea_m) and _is_blank(prev_m):
            blank_run += 1
            if blank_run >= 6:
                break
            continue
        blank_run = 0

        m = _to_month(mes)
        if m is None:
            continue

        rows.append(
            (
                m,
                _to_float(pla_a),
                _to_float(pla_m),
                _to_float(rea_m),
                _to_float(prev_m),
            )
        )

    df = pd.DataFrame(
        rows,
        columns=["MÊS", "PLANEJADO ACUM. (%)", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)", previsto_name],
    )

    # ✅ padroniza: o app procura exatamente "PREVISTO MENSAL (%)"
    if previsto_name != "PREVISTO MENSAL (%)" and previsto_name in df.columns:
        df = df.rename(columns={previsto_name: "PREVISTO MENSAL (%)"})

    if not df.empty:
        df = df.sort_values("MÊS")
    return df


def read_acrescimos_economias(ws: Worksheet) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Bloco (linha ~47):
      Acréscimos: col A..F  (inclui JUSTIFICATIVAS na F)
      Economias:  col G..L  (inclui JUSTIFICATIVAS na L)

    Retorna: (df_acrescimos, df_economias) com colunas:
      DESCRIÇÃO, ORÇAMENTO INICIAL, ORÇAMENTO REAJUSTADO, CUSTO FINAL, VARIAÇÃO, JUSTIFICATIVAS
    """
    header_row = _find_row_with(ws, [(1, "DESCRIÇÃO"), (6, "JUSTIFICATIVAS"), (7, "DESCRIÇÃO"), (12, "JUSTIFICATIVAS")], max_scan=800)
    if header_row is None:
        cols = ["DESCRIÇÃO", "ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO", "JUSTIFICATIVAS"]
        return pd.DataFrame(columns=cols), pd.DataFrame(columns=cols)

    def read_side(start_col: int) -> pd.DataFrame:
        rows = []
        blank_run = 0
        for r in range(header_row + 1, (ws.max_row or header_row + 1) + 1):
            desc = ws.cell(r, start_col + 0).value
            o_ini = ws.cell(r, start_col + 1).value
            o_rea = ws.cell(r, start_col + 2).value
            c_fin = ws.cell(r, start_col + 3).value
            var_  = ws.cell(r, start_col + 4).value
            just  = ws.cell(r, start_col + 5).value

            if _is_blank(desc) and _is_blank(o_ini) and _is_blank(o_rea) and _is_blank(c_fin) and _is_blank(var_) and _is_blank(just):
                blank_run += 1
                if blank_run >= 8:
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

        df = pd.DataFrame(
            rows,
            columns=["DESCRIÇÃO", "ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO", "JUSTIFICATIVAS"],
        )
        return df

    df_acres = read_side(1)   # A..F
    df_econ  = read_side(7)   # G..L
    return df_acres, df_econ
