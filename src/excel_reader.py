# src/excel_reader.py
from __future__ import annotations
from io import BytesIO
from typing import Iterable
import pandas as pd
from openpyxl import load_workbook

from .utils import norm, is_blank, to_month, to_float

def load_wb(file_or_bytes) -> "openpyxl.Workbook":
    """
    file_or_bytes: caminho (str) ou bytes/BytesIO do uploader.
    """
    if isinstance(file_or_bytes, (bytes, bytearray)):
        bio = BytesIO(file_or_bytes)
        return load_workbook(bio, data_only=True)
    if hasattr(file_or_bytes, "read"):
        # UploadedFile do Streamlit
        bio = BytesIO(file_or_bytes.read())
        return load_workbook(bio, data_only=True)
    return load_workbook(file_or_bytes, data_only=True)

def sheetnames(wb) -> list[str]:
    # ignora abas “técnicas” se existirem
    return [s for s in wb.sheetnames if norm(s) not in ("LEIA-ME", "LEIAME", "README")]

def _sheet_matrix(ws, max_rows=600, max_cols=40):
    mat = []
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=max_rows, values_only=True), start=1):
        mat.append(list(row[:max_cols]))
    return mat

def _find_header_row(mat, required_headers: Iterable[str]):
    req = [norm(h) for h in required_headers]
    for i, row in enumerate(mat):
        row_norm = [norm(c) for c in row]
        if all(h in row_norm for h in req):
            mapping = {h: row_norm.index(h) for h in req}
            return i, mapping
    return None, None

def read_table_by_headers(ws, required_headers: list[str], stop_on_first_blank_month=True) -> pd.DataFrame:
    """
    Acha a linha do cabeçalho pelo texto e lê até acabar (MÊS vazio ou linha vazia).
    """
    mat = _sheet_matrix(ws)
    h_i, mapping = _find_header_row(mat, required_headers)
    if h_i is None:
        return pd.DataFrame()

    cols = required_headers
    idxs = [mapping[norm(c)] for c in cols]

    data = []
    for r in range(h_i + 1, len(mat)):
        row = mat[r]

        vals = []
        for j in idxs:
            vals.append(row[j] if j < len(row) else None)

        # regra de parada: MÊS vazio (quando existe col MÊS)
        if norm(cols[0]) in ("MÊS", "MES"):
            m = to_month(vals[0])
            if m is None and stop_on_first_blank_month:
                break

        # se linha toda vazia nas colunas do bloco, para
        if all(is_blank(v) for v in vals):
            break

        data.append(vals)

    df = pd.DataFrame(data, columns=cols)
    return df

def read_resumo_financeiro(ws) -> dict:
    """
    Lê os KPIs pelo nome (col A = label, col B = valor).
    """
    keys = [
        "ORÇAMENTO INICIAL (R$)",
        "ORÇAMENTO REAJUSTADO (R$)",
        "DESEMBOLSO ACUMULADO (R$)",
        "A PAGAR (R$)",
        "SALDO A INCORRER (R$)",
        "CUSTO FINAL (R$)",
        "VARIAÇÃO (R$)",
    ]
    wanted = {norm(k): k for k in keys}
    out = {k: None for k in keys}

    for r in range(1, 80):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        if norm(a) in wanted:
            out[wanted[norm(a)]] = to_float(b)

    return out

def read_indice(ws) -> pd.DataFrame:
    df = read_table_by_headers(ws, ["MÊS", "ÍNDICE PROJETADO"])
    if df.empty:
        return df
    df["MÊS"] = df["MÊS"].apply(to_month)
    df["ÍNDICE PROJETADO"] = df["ÍNDICE PROJETADO"].apply(to_float)
    df = df.dropna(subset=["MÊS"])
    return df

def read_financeiro(ws) -> pd.DataFrame:
    df = read_table_by_headers(ws, ["MÊS", "DESEMBOLSO DO MÊS (R$)", "MEDIDO NO MÊS (R$)"])
    if df.empty:
        return df
    df["MÊS"] = df["MÊS"].apply(to_month)
    df["DESEMBOLSO DO MÊS (R$)"] = df["DESEMBOLSO DO MÊS (R$)"].apply(to_float)
    df["MEDIDO NO MÊS (R$)"] = df["MEDIDO NO MÊS (R$)"].apply(to_float)
    df = df.dropna(subset=["MÊS"])
    df = df.sort_values("MÊS")
    df["DESEMBOLSO ACUM. (R$)"] = df["DESEMBOLSO DO MÊS (R$)"].fillna(0).cumsum()
    df["MEDIDO ACUM. (R$)"] = df["MEDIDO NO MÊS (R$)"].fillna(0).cumsum()
    return df

def read_prazo(ws) -> pd.DataFrame:
    """
    Aceita seu layout novo:
    MÊS | PLANEJADO ACUM. (%) | PLANEJADO MÊS (%) | COMPROMETIDO Mês (%) | REALIZADO Mês (%)
    (Se COMPROMETIDO não existir, ok.)
    """
    # tenta com comprometido
    headers_5 = ["MÊS", "PLANEJADO ACUM. (%)", "PLANEJADO MÊS (%)", "COMPROMETIDO Mês (%)", "REALIZADO Mês (%)"]
    df = read_table_by_headers(ws, headers_5)
    if df.empty:
        # fallback sem comprometido
        headers_4 = ["MÊS", "PLANEJADO ACUM. (%)", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)"]
        df = read_table_by_headers(ws, headers_4)

    if df.empty:
        return df

    # normaliza
    df["MÊS"] = df["MÊS"].apply(to_month)
    df["PLANEJADO ACUM. (%)"] = df["PLANEJADO ACUM. (%)"].apply(to_float)
    df["PLANEJADO MÊS (%)"] = df["PLANEJADO MÊS (%)"].apply(to_float)

    if "COMPROMETIDO Mês (%)" in df.columns:
        df["COMPROMETIDO Mês (%)"] = df["COMPROMETIDO Mês (%)"].apply(to_float)

    # Realizado mês
    if "REALIZADO Mês (%)" in df.columns:
        df["REALIZADO Mês (%)"] = df["REALIZADO Mês (%)"].apply(to_float)

    df = df.dropna(subset=["MÊS"]).sort_values("MÊS")

    # Se você alimenta % como 0–1 (ex 0.10), mantém assim.
    # Curva S do Real: acumula a coluna Realizado Mês (%) (se existir)
    if "REALIZADO Mês (%)" in df.columns:
        df["REAL ACUM. (%)"] = df["REALIZADO Mês (%)"].fillna(0).cumsum()

    if "COMPROMETIDO Mês (%)" in df.columns:
        df["COMPROMETIDO ACUM. (%)"] = df["COMPROMETIDO Mês (%)"].fillna(0).cumsum()

    return df

def _find_row_with_text(ws, text: str, max_rows=800, max_cols=20):
    t = norm(text)
    for r in range(1, max_rows + 1):
        for c in range(1, max_cols + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and norm(v) == t:
                return r, c
    return None, None

def read_side_table(ws, title_text: str, headers: list[str], start_col: int, max_rows=120) -> pd.DataFrame:
    """
    Lê uma tabela “colada” (tipo ACRÉSCIMOS/ECONOMIAS) onde:
    - existe uma célula com o título (ACRÉSCIMOS ou ECONOMIAS)
    - a linha seguinte tem o cabeçalho
    - os dados começam na próxima linha
    e para quando DESCRIÇÃO estiver vazia.
    """
    r_title, c_title = _find_row_with_text(ws, title_text)
    if r_title is None:
        return pd.DataFrame()

    r_header = r_title + 1
    # valida header
    for i, h in enumerate(headers):
        v = ws.cell(r_header, start_col + i).value
        if norm(v) != norm(h):
            # se não bater, tenta achar header na mesma linha (variações de colagem)
            pass

    data = []
    r = r_header + 1
    for _ in range(max_rows):
        desc = ws.cell(r, start_col).value
        if is_blank(desc):
            break
        row_vals = [ws.cell(r, start_col + i).value for i in range(len(headers))]
        data.append(row_vals)
        r += 1

    df = pd.DataFrame(data, columns=headers)
    # converte numéricos
    for col in headers[1:]:
        df[col] = df[col].apply(to_float)
    return df

def read_acrescimos_economias(ws) -> tuple[pd.DataFrame, pd.DataFrame]:
    headers = ["DESCRIÇÃO", "ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO"]
    # No seu template: ACRÉSCIMOS começa em A; ECONOMIAS começa em G
    acres = read_side_table(ws, "ACRÉSCIMOS", headers, start_col=1)
    econ  = read_side_table(ws, "ECONOMIAS", headers, start_col=7)
    return acres, econ
