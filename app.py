from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


# ----------------------------
# Load / Sheets
# ----------------------------
def load_wb(path: str | Path) -> Workbook:
    path = Path(path)
    keep_vba = path.suffix.lower() == ".xlsm"
    # data_only=True: lê valores "calculados" (cache do Excel). Salve o arquivo no Excel antes de subir.
    return load_workbook(path, data_only=True, keep_vba=keep_vba)


def sheetnames(wb: Workbook) -> list[str]:
    ignore = {"LEIA-ME", "README", "READ ME"}
    return [s for s in wb.sheetnames if s.strip().upper() not in ignore]


# ----------------------------
# Utils internos (scan)
# ----------------------------
def _norm(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip().upper()


def _find_row_with_headers(
    ws: Worksheet,
    headers: list[str],
    max_row: int = 250,
    max_col: int = 60,
) -> tuple[int | None, list[str] | None]:
    want = [_norm(h) for h in headers]
    r_max = min(ws.max_row, max_row)
    c_max = min(ws.max_column, max_col)

    for r in range(1, r_max + 1):
        row_vals = [_norm(ws.cell(r, c).value) for c in range(1, c_max + 1)]
        ok = True
        for h in want:
            if h and h not in row_vals:
                ok = False
                break
        if ok:
            return r, row_vals
    return None, None


def _col_idx(row_vals: list[str], header: str) -> int | None:
    h = _norm(header)
    for i, v in enumerate(row_vals, start=1):
        if v == h:
            return i
    return None


def _to_df_numeric(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


# ----------------------------
# Leitura dos blocos (dinâmico)
# ----------------------------
def read_resumo_financeiro(ws: Worksheet) -> dict[str, float | None]:
    """
    Procura pelo bloco:
      RESUMO FINANCEIRO (INSIRA OS VALORES)
    E lê as linhas abaixo com:
      COL A = nome do campo
      COL B = valor (R$)
    """
    # A1 costuma ter o título
    # mas a gente é robusto: varre procurando "RESUMO FINANCEIRO"
    title_row = None
    for r in range(1, min(ws.max_row, 50) + 1):
        if "RESUMO FINANCEIRO" in _norm(ws.cell(r, 1).value):
            title_row = r
            break

    if title_row is None:
        return {}

    out: dict[str, float | None] = {}
    r = title_row + 1
    while r <= ws.max_row:
        key = ws.cell(r, 1).value
        val = ws.cell(r, 2).value
        if key in (None, ""):
            break
        out[_norm(key)] = float(val) if isinstance(val, (int, float)) else (None if val is None else float(val))
        r += 1

    # normaliza para as chaves que o app usa
    # (mantém no formato "ORÇAMENTO INICIAL (R$)" etc)
    fixed: dict[str, float | None] = {}
    for k, v in out.items():
        fixed[k] = v
    return fixed


def read_indice(ws: Worksheet) -> pd.DataFrame:
    """
    Tabela:
      MÊS | ÍNDICE PROJETADO
    """
    hr, row_vals = _find_row_with_headers(ws, ["MÊS", "ÍNDICE PROJETADO"], max_row=200, max_col=20)
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
        if mes in (None, ""):
            break
        idx = ws.cell(r, c_idx).value
        data.append({"MÊS": mes, "ÍNDICE PROJETADO": idx})
        r += 1

    df = pd.DataFrame(data)
    if not df.empty:
        df["MÊS"] = pd.to_datetime(df["MÊS"], errors="coerce")
        df["ÍNDICE PROJETADO"] = pd.to_numeric(df["ÍNDICE PROJETADO"], errors="coerce")
    return df


def read_financeiro(ws: Worksheet) -> pd.DataFrame:
    """
    Tabela:
      MÊS | DESEMBOLSO DO MÊS (R$) | MEDIDO NO MÊS (R$)
    """
    hr, row_vals = _find_row_with_headers(
        ws,
        ["MÊS", "DESEMBOLSO DO MÊS (R$)", "MEDIDO NO MÊS (R$)"],
        max_row=220,
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
        if mes in (None, ""):
            break
        data.append(
            {
                "MÊS": mes,
                "DESEMBOLSO DO MÊS (R$)": ws.cell(r, c_des).value,
                "MEDIDO NO MÊS (R$)": ws.cell(r, c_med).value,
            }
        )
        r += 1

    df = pd.DataFrame(data)
    if not df.empty:
        df["MÊS"] = pd.to_datetime(df["MÊS"], errors="coerce")
        df = _to_df_numeric(df, ["DESEMBOLSO DO MÊS (R$)", "MEDIDO NO MÊS (R$)"])
    return df


def read_prazo(ws: Worksheet) -> pd.DataFrame:
    """
    Tabela (Prazo):
      MÊS | PLANEJADO ACUM. (%) | PLANEJADO MÊS (%) | REALIZADO Mês (%) | (às vezes REALIZADO ACUM. (%) em outra coluna)
    """
    hr, row_vals = _find_row_with_headers(
        ws,
        ["MÊS", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)"],
        max_row=300,
        max_col=25,
    )
    if hr is None or row_vals is None:
        return pd.DataFrame(columns=["MÊS", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)"])

    c_mes = _col_idx(row_vals, "MÊS")
    c_p_mes = _col_idx(row_vals, "PLANEJADO MÊS (%)")
    c_r_mes = _col_idx(row_vals, "REALIZADO Mês (%)")
    if c_mes is None or c_p_mes is None or c_r_mes is None:
        return pd.DataFrame(columns=["MÊS", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)"])

    data = []
    r = hr + 1
    while r <= ws.max_row:
        mes = ws.cell(r, c_mes).value
        if mes in (None, ""):
            break
        data.append(
            {
                "MÊS": mes,
                "PLANEJADO MÊS (%)": ws.cell(r, c_p_mes).value,
                "REALIZADO Mês (%)": ws.cell(r, c_r_mes).value,
            }
        )
        r += 1

    df = pd.DataFrame(data)
    if not df.empty:
        df["MÊS"] = pd.to_datetime(df["MÊS"], errors="coerce")
        df = _to_df_numeric(df, ["PLANEJADO MÊS (%)", "REALIZADO Mês (%)"])
    return df


def read_acrescimos_economias(ws: Worksheet) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Bloco lado a lado:
      A..E = ACRÉSCIMOS
      G..K = ECONOMIAS
    Headers esperados em ambos:
      DESCRIÇÃO | ORÇAMENTO INICIAL | ORÇAMENTO REAJUSTADO | CUSTO FINAL | VARIAÇÃO
    """
    # acha linha que tem "ACRÉSCIMOS" em A e "ECONOMIAS" em G (normalmente)
    base_row = None
    for r in range(1, min(ws.max_row, 400) + 1):
        if _norm(ws.cell(r, 1).value) == "ACRÉSCIMOS" and _norm(ws.cell(r, 7).value) == "ECONOMIAS":
            base_row = r
            break

    if base_row is None:
        # fallback: tenta achar headers dos 2 lados
        return (
            pd.DataFrame(columns=["DESCRIÇÃO", "ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO"]),
            pd.DataFrame(columns=["DESCRIÇÃO", "ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO"]),
        )

    header_row = base_row + 1

    # colunas fixas
    cols_left = {"DESCRIÇÃO": 1, "ORÇAMENTO INICIAL": 2, "ORÇAMENTO REAJUSTADO": 3, "CUSTO FINAL": 4, "VARIAÇÃO": 5}
    cols_right = {"DESCRIÇÃO": 7, "ORÇAMENTO INICIAL": 8, "ORÇAMENTO REAJUSTADO": 9, "CUSTO FINAL": 10, "VARIAÇÃO": 11}

    def read_side(colmap: dict[str, int]) -> pd.DataFrame:
        data = []
        r = header_row + 1
        while r <= ws.max_row:
            desc = ws.cell(r, colmap["DESCRIÇÃO"]).value
            var = ws.cell(r, colmap["VARIAÇÃO"]).value
            # regra de parada: descrição vazia (ou linha totalmente vazia)
            if desc in (None, ""):
                # se for a “linha zerada” do template, também para
                if var in (None, "", 0):
                    break
                # se tiver var mas sem desc, ainda para (tabela suja)
                break

            row = {k: ws.cell(r, c).value for k, c in colmap.items()}
            data.append(row)
            r += 1

        df = pd.DataFrame(data)
        if not df.empty:
            df = _to_df_numeric(df, ["ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO"])
        return df

    df_acres = read_side(cols_left)
    df_econ = read_side(cols_right)

    return df_acres, df_econ
