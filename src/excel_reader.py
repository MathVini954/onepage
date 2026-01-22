# src/excel_reader.py
from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .utils import norm, is_blank, to_month, to_float


# ----------------------------
# Load / Sheets
# ----------------------------
def load_wb(path: str | Path) -> Workbook:
    path = Path(path)
    keep_vba = path.suffix.lower() == ".xlsm"
    return load_workbook(path, data_only=True, keep_vba=keep_vba)


def sheetnames(wb: Workbook) -> list[str]:
    ignore = {"LEIA-ME", "README", "READ ME"}
    return [s for s in wb.sheetnames if norm(s) not in ignore]


# ----------------------------
# Scan helpers
# ----------------------------
def _find_row_contains(
    ws: Worksheet,
    needle: str,
    col: int = 1,
    max_row: int = 200,
) -> int | None:
    n = norm(needle)
    for r in range(1, min(ws.max_row, max_row) + 1):
        if n in norm(ws.cell(r, col).value):
            return r
    return None


def _find_header_row(
    ws: Worksheet,
    headers: list[str],
    max_row: int = 300,
    max_col: int = 60,
) -> tuple[int | None, list[str] | None]:
    want = [norm(h) for h in headers]
    r_max = min(ws.max_row, max_row)
    c_max = min(ws.max_column, max_col)

    for r in range(1, r_max + 1):
        row_vals = [norm(ws.cell(r, c).value) for c in range(1, c_max + 1)]
        ok = True
        for h in want:
            if h and h not in row_vals:
                ok = False
                break
        if ok:
            return r, row_vals
    return None, None


def _col_idx(row_vals: list[str], header: str) -> int | None:
    h = norm(header)
    for i, v in enumerate(row_vals, start=1):
        if v == h:
            return i
    return None


# ----------------------------
# Block readers
# ----------------------------
def read_resumo_financeiro(ws: Worksheet) -> dict[str, float | None]:
    """
    Bloco:
      RESUMO FINANCEIRO (INSIRA OS VALORES)
      coluna A = item
      coluna B = valor
    """
    title_row = _find_row_contains(ws, "RESUMO FINANCEIRO", col=1, max_row=80)
    if title_row is None:
        return {}

    mapping = {
        "ORÇAMENTO INICIAL": "ORÇAMENTO INICIAL (R$)",
        "ORCAMENTO INICIAL": "ORÇAMENTO INICIAL (R$)",
        "ORÇAMENTO REAJUSTADO": "ORÇAMENTO REAJUSTADO (R$)",
        "ORCAMENTO REAJUSTADO": "ORÇAMENTO REAJUSTADO (R$)",
        "DESEMBOLSO ACUMULADO": "DESEMBOLSO ACUMULADO (R$)",
        "A PAGAR": "A PAGAR (R$)",
        "SALDO A INCORRER": "SALDO A INCORRER (R$)",
        "CUSTO FINAL": "CUSTO FINAL (R$)",
        "VARIAÇÃO": "VARIAÇÃO (R$)",
        "VARIACAO": "VARIAÇÃO (R$)",
    }

    out: dict[str, float | None] = {}
    r = title_row + 1
    while r <= ws.max_row:
        k = ws.cell(r, 1).value
        v = ws.cell(r, 2).value
        if is_blank(k):
            break
        key_norm = norm(k)
        key = mapping.get(key_norm, str(k).strip())
        out[key] = to_float(v)
        r += 1

    return out


def read_indice(ws: Worksheet) -> pd.DataFrame:
    hr, row_vals = _find_header_row(ws, ["MÊS", "ÍNDICE PROJETADO"], max_row=250, max_col=30)
    if hr is None or row_vals is None:
        return pd.DataFrame(columns=["MÊS", "ÍNDICE PROJETADO"])

    c_mes = _col_idx(row_vals, "MÊS")
    c_idx = _col_idx(row_vals, "ÍNDICE PROJETADO")
    if c_mes is None or c_idx is None:
        return pd.DataFrame(columns=["MÊS", "ÍNDICE PROJETADO"])

    data = []
    r = hr + 1
    while r <= ws.max_row:
        mes = ws.cell(r, c_mes).value
        if is_blank(mes):
            break
        data.append(
            {
                "MÊS": to_month(mes),
                "ÍNDICE PROJETADO": to_float(ws.cell(r, c_idx).value),
            }
        )
        r += 1

    df = pd.DataFrame(data)
    df = df.dropna(subset=["MÊS"])
    return df


def read_financeiro(ws: Worksheet) -> pd.DataFrame:
    hr, row_vals = _find_header_row(
        ws,
        ["MÊS", "DESEMBOLSO DO MÊS (R$)", "MEDIDO NO MÊS (R$)"],
        max_row=300,
        max_col=40,
    )
    if hr is None or row_vals is None:
        return pd.DataFrame(columns=["MÊS", "DESEMBOLSO DO MÊS (R$)", "MEDIDO NO MÊS (R$)"])

    c_mes = _col_idx(row_vals, "MÊS")
    c_des = _col_idx(row_vals, "DESEMBOLSO DO MÊS (R$)")
    c_med = _col_idx(row_vals, "MEDIDO NO MÊS (R$)")
    if c_mes is None or c_des is None or c_med is None:
        return pd.DataFrame(columns=["MÊS", "DESEMBOLSO DO MÊS (R$)", "MEDIDO NO MÊS (R$)"])

    data = []
    r = hr + 1
    while r <= ws.max_row:
        mes = ws.cell(r, c_mes).value
        if is_blank(mes):
            break
        data.append(
            {
                "MÊS": to_month(mes),
                "DESEMBOLSO DO MÊS (R$)": to_float(ws.cell(r, c_des).value),
                "MEDIDO NO MÊS (R$)": to_float(ws.cell(r, c_med).value),
            }
        )
        r += 1

    df = pd.DataFrame(data)
    df = df.dropna(subset=["MÊS"])
    return df


def read_prazo(ws: Worksheet) -> pd.DataFrame:
    hr, row_vals = _find_header_row(
        ws,
        ["MÊS", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)"],
        max_row=350,
        max_col=30,
    )
    if hr is None or row_vals is None:
        return pd.DataFrame(columns=["MÊS", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)"])

    c_mes = _col_idx(row_vals, "MÊS")
    c_p = _col_idx(row_vals, "PLANEJADO MÊS (%)")
    c_r = _col_idx(row_vals, "REALIZADO Mês (%)")
    if c_mes is None or c_p is None or c_r is None:
        return pd.DataFrame(columns=["MÊS", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)"])

    data = []
    r = hr + 1
    while r <= ws.max_row:
        mes = ws.cell(r, c_mes).value
        if is_blank(mes):
            break
        data.append(
            {
                "MÊS": to_month(mes),
                "PLANEJADO MÊS (%)": to_float(ws.cell(r, c_p).value),
                "REALIZADO Mês (%)": to_float(ws.cell(r, c_r).value),
            }
        )
        r += 1

    df = pd.DataFrame(data)
    df = df.dropna(subset=["MÊS"])
    return df


def read_acrescimos_economias(ws: Worksheet) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Espera blocos lado a lado:
      A..E = ACRÉSCIMOS
      G..K = ECONOMIAS
    """
    base_row = None
    for r in range(1, min(ws.max_row, 500) + 1):
        a = norm(ws.cell(r, 1).value)
        g = norm(ws.cell(r, 7).value)
        if ("ACRÉSCIM" in a or "ACRESCIM" in a) and ("ECONOM" in g):
            base_row = r
            break

    # fallback: acha linha do header "DESCRIÇÃO" em A e em G
    if base_row is None:
        for r in range(1, min(ws.max_row, 500) + 1):
            if norm(ws.cell(r, 1).value) == "DESCRIÇÃO" and norm(ws.cell(r, 7).value) == "DESCRIÇÃO":
                base_row = r - 1  # uma acima seria o título
                break

    if base_row is None:
        empty = pd.DataFrame(columns=["DESCRIÇÃO", "ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO"])
        return empty, empty

    header_row = base_row + 1

    cols_left = {"DESCRIÇÃO": 1, "ORÇAMENTO INICIAL": 2, "ORÇAMENTO REAJUSTADO": 3, "CUSTO FINAL": 4, "VARIAÇÃO": 5}
    cols_right = {"DESCRIÇÃO": 7, "ORÇAMENTO INICIAL": 8, "ORÇAMENTO REAJUSTADO": 9, "CUSTO FINAL": 10, "VARIAÇÃO": 11}

    def read_side(colmap: dict[str, int]) -> pd.DataFrame:
        data = []
        r = header_row + 1
        while r <= ws.max_row:
            desc = ws.cell(r, colmap["DESCRIÇÃO"]).value
            var = ws.cell(r, colmap["VARIAÇÃO"]).value

            if is_blank(desc):
                # para quando tabela acabar
                if is_blank(var) or to_float(var) == 0:
                    break
                break

            data.append({k: (ws.cell(r, c).value if k == "DESCRIÇÃO" else to_float(ws.cell(r, c).value)) for k, c in colmap.items()})
            r += 1

        df = pd.DataFrame(data)
        return df

    return read_side(cols_left), read_side(cols_right)
