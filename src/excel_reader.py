from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd
import openpyxl

from .utils import norm, is_blank, to_month, to_float


# ------------------------------------------------------------
# Workbook helpers
# ------------------------------------------------------------
def load_wb(path: Path):
    # keep_vba=True keeps XLSM macros intact
    return openpyxl.load_workbook(path, keep_vba=True, data_only=True)


def sheetnames(wb) -> List[str]:
    """Return obra tabs only (exclude auxiliary sheets)."""
    skip = {"ORÇAMENTO_RESUMO", "ORCAMENTO_RESUMO"}
    out = []
    for n in wb.sheetnames:
        if norm(n) in {norm(x) for x in skip}:
            continue
        out.append(n)
    return out


# ------------------------------------------------------------
# Internal search utilities
# ------------------------------------------------------------
def _iter_used_cells(ws, max_row: int = 800, max_col: int = 40):
    for r in range(1, min(ws.max_row, max_row) + 1):
        for c in range(1, min(ws.max_column, max_col) + 1):
            yield r, c, ws.cell(row=r, column=c).value


def _cell_matches_required(cell_nv: str, req_nv: str) -> bool:
    """
    Header matching rules:
      - For MES/OBRA: MUST be exact (avoid matching 'MÊS A MÊS')
      - For others: allow exact or 'contains' (to accept '(R$)' etc)
    """
    if not cell_nv:
        return False
    if req_nv in ("MES", "OBRA"):
        return cell_nv == req_nv
    return cell_nv == req_nv or req_nv in cell_nv


def _find_header_row(ws, required_headers: List[str], max_scan: int = 600) -> Tuple[int, Dict[str, int]]:
    """
    Find a row where all required headers exist in the SAME row (robust).
    Returns (header_row, mapping req_norm->column_index).
    """
    req = [norm(h) for h in required_headers]

    for r in range(1, min(ws.max_row, max_scan) + 1):
        cells = []
        for c in range(1, min(ws.max_column, 80) + 1):
            nv = norm(ws.cell(row=r, column=c).value)
            if nv:
                cells.append((c, nv))

        if not cells:
            continue

        mapping: Dict[str, int] = {}
        used_cols: set[int] = set()
        ok = True

        for h in req:
            candidates = [c for c, nv in cells if _cell_matches_required(nv, h)]
            if not candidates:
                ok = False
                break
            # prefer a column not used yet (avoid mapping 2 headers to same cell)
            pick = next((c for c in candidates if c not in used_cols), candidates[0])
            mapping[h] = pick
            used_cols.add(pick)

        if ok:
            return r, mapping

    raise ValueError(f"Header não encontrado: {required_headers}")


def _read_table(ws, header_row: int, col_map: Dict[str, int], max_rows: int = 600) -> pd.DataFrame:
    """Read rows after header_row until a stop condition (5 consecutive blank rows)."""
    inv = sorted(((col, h) for h, col in col_map.items()), key=lambda x: x[0])
    cols = [h for _, h in inv]

    rows = []
    blanks = 0
    for r in range(header_row + 1, min(ws.max_row, header_row + max_rows) + 1):
        row = {}
        row_has_any = False
        for h in cols:
            c = col_map[h]
            v = ws.cell(row=r, column=c).value
            if not is_blank(v):
                row_has_any = True
            row[h] = v

        if not row_has_any:
            blanks += 1
            if blanks >= 5:
                break
            continue

        blanks = 0
        rows.append(row)

    return pd.DataFrame(rows)


def _find_block_title(ws, title_contains: str, max_scan: int = 900) -> Tuple[int, int] | None:
    target = norm(title_contains)
    for r, c, v in _iter_used_cells(ws, max_row=max_scan, max_col=40):
        if target in norm(v):
            return r, c
    return None


# ------------------------------------------------------------
# Readers
# ------------------------------------------------------------
def read_resumo_financeiro(ws) -> Dict[str, float]:
    keys = [
        "ORÇAMENTO INICIAL",
        "ORÇAMENTO REAJUSTADO",
        "DESEMBOLSO ACUMULADO",
        "A PAGAR",
        "SALDO A INCORRER",
        "CUSTO FINAL",
        "VARIAÇÃO",
    ]
    out: Dict[str, float] = {}
    for r in range(1, min(ws.max_row, 250) + 1):
        label = ws.cell(row=r, column=1).value
        nl = norm(label)
        if not nl:
            continue
        for k in keys:
            if norm(k) in nl:
                out[f"{k} (R$)"] = to_float(ws.cell(row=r, column=2).value)
    return out


def read_indice(ws) -> pd.DataFrame:
    header_row, cols = _find_header_row(ws, ["MÊS", "ÍNDICE PROJETADO"], max_scan=800)
    df = _read_table(ws, header_row, cols, max_rows=800)

    # rename
    rename = {}
    for k in df.columns:
        nk = norm(k)
        if nk == "MES":
            rename[k] = "MÊS"
        elif "INDICE PROJETADO" in nk:
            rename[k] = "ÍNDICE PROJETADO"
    df = df.rename(columns=rename)

    if "MÊS" not in df.columns or "ÍNDICE PROJETADO" not in df.columns:
        return pd.DataFrame()

    df["MÊS"] = df["MÊS"].apply(to_month)
    df["ÍNDICE PROJETADO"] = df["ÍNDICE PROJETADO"].apply(to_float)
    df = df.dropna(subset=["MÊS", "ÍNDICE PROJETADO"]).sort_values("MÊS")
    return df


def read_financeiro(ws) -> pd.DataFrame:
    header_row, cols = _find_header_row(ws, ["MÊS", "DESEMBOLSO", "MEDIDO"], max_scan=1200)
    df = _read_table(ws, header_row, cols, max_rows=1200)

    rename = {}
    for k in df.columns:
        nk = norm(k)
        if nk == "MES":
            rename[k] = "MÊS"
        elif "DESEMBOLSO" in nk:
            rename[k] = "DESEMBOLSO DO MÊS (R$)"
        elif "MEDIDO" in nk:
            rename[k] = "MEDIDO NO MÊS (R$)"
    df = df.rename(columns=rename)

    if "MÊS" not in df.columns:
        return pd.DataFrame()

    df["MÊS"] = df["MÊS"].apply(to_month)
    if "DESEMBOLSO DO MÊS (R$)" in df.columns:
        df["DESEMBOLSO DO MÊS (R$)"] = df["DESEMBOLSO DO MÊS (R$)"].apply(to_float)
    if "MEDIDO NO MÊS (R$)" in df.columns:
        df["MEDIDO NO MÊS (R$)"] = df["MEDIDO NO MÊS (R$)"].apply(to_float)

    df = df.dropna(subset=["MÊS"]).sort_values("MÊS")
    return df


def read_prazo(ws) -> pd.DataFrame:
    header_row, _ = _find_header_row(ws, ["MÊS", "PLANEJADO MÊS", "REALIZADO"], max_scan=1600)

    interesting: Dict[str, int] = {}
    for c in range(1, min(ws.max_column, 80) + 1):
        v = ws.cell(row=header_row, column=c).value
        nv = norm(v)
        if not nv:
            continue
        if nv == "MES":
            interesting["MÊS"] = c
        if "PLANEJADO" in nv and "ACUM" in nv:
            interesting["PLANEJADO ACUM. (%)"] = c
        if "PLANEJADO" in nv and "MES" in nv:
            interesting["PLANEJADO MÊS (%)"] = c
        if "PREVISTO" in nv and "MENSAL" in nv:
            interesting["PREVISTO MENSAL (%)"] = c
        if "REALIZADO" in nv and "MES" in nv:
            interesting["REALIZADO Mês (%)"] = c

    col_map = {norm(k): v for k, v in interesting.items()}
    df = _read_table(ws, header_row, col_map, max_rows=1600)
    rename = {norm(k): k for k in interesting.keys()}
    df = df.rename(columns=rename)

    if "MÊS" not in df.columns:
        return pd.DataFrame()

    df["MÊS"] = df["MÊS"].apply(to_month)
    df = df.dropna(subset=["MÊS"]).sort_values("MÊS")

    for c in [c for c in df.columns if c != "MÊS"]:
        df[c] = df[c].apply(to_float)
    return df


def read_acrescimos_economias(ws) -> Tuple[pd.DataFrame, pd.DataFrame]:
    def read_one(title: str) -> pd.DataFrame:
        pos = _find_block_title(ws, title)
        if not pos:
            return pd.DataFrame()
        tr, _ = pos

        # find header row (DESCRIÇÃO) below title
        header_row = None
        for r in range(tr, tr + 20):
            for c in range(1, min(ws.max_column, 80) + 1):
                if "DESCRICAO" in norm(ws.cell(row=r, column=c).value):
                    header_row = r
                    break
            if header_row:
                break
        if not header_row:
            return pd.DataFrame()

        col_map: Dict[str, int] = {}
        for c in range(1, min(ws.max_column, 80) + 1):
            hv = ws.cell(row=header_row, column=c).value
            nh = norm(hv)
            if not nh:
                continue
            if "DESCRICAO" in nh:
                col_map["DESCRIÇÃO"] = c
            elif "ORCAMENTO" in nh and "INICIAL" in nh:
                col_map["ORÇAMENTO INICIAL"] = c
            elif "ORCAMENTO" in nh and ("REAJUST" in nh or "REAJUSTADO" in nh):
                col_map["ORÇAMENTO REAJUSTADO"] = c
            elif "CUSTO" in nh and "FINAL" in nh:
                col_map["CUSTO FINAL"] = c
            elif "VARIAC" in nh:
                col_map["VARIAÇÃO"] = c
            elif "JUSTIFICAT" in nh:
                col_map["JUSTIFICATIVAS"] = c

        if "DESCRIÇÃO" not in col_map or "VARIAÇÃO" not in col_map:
            return pd.DataFrame()

        rows = []
        blanks = 0
        for r in range(header_row + 1, min(ws.max_row, header_row + 2000) + 1):
            desc = ws.cell(row=r, column=col_map["DESCRIÇÃO"]).value
            if is_blank(desc):
                blanks += 1
                if blanks >= 6:
                    break
                continue
            blanks = 0
            row = {k: ws.cell(row=r, column=c).value for k, c in col_map.items()}
            rows.append(row)

        df = pd.DataFrame(rows)
        for c in ["ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO"]:
            if c in df.columns:
                df[c] = df[c].apply(to_float)
        if "JUSTIFICATIVAS" in df.columns:
            df["JUSTIFICATIVAS"] = df["JUSTIFICATIVAS"].fillna("").astype(str)
        if "DESCRIÇÃO" in df.columns:
            df["DESCRIÇÃO"] = df["DESCRIÇÃO"].astype(str)
        return df

    df_acres = read_one("ACRÉSCIMOS")
    df_econ = read_one("ECONOMIAS")
    return df_acres, df_econ


def read_orcamento_resumo(wb) -> pd.DataFrame:
    name = None
    for n in wb.sheetnames:
        if norm(n) in ("ORCAMENTO_RESUMO", "ORÇAMENTO_RESUMO"):
            name = n
            break
    if not name:
        return pd.DataFrame()

    ws = wb[name]
    header_row = None
    header_col = None
    for r in range(1, min(ws.max_row, 80) + 1):
        for c in range(1, min(ws.max_column, 120) + 1):
            if norm(ws.cell(row=r, column=c).value) == "OBRA":
                header_row = r
                header_col = c
                break
        if header_row:
            break
    if not header_row:
        return pd.DataFrame()

    headers = []
    for c in range(header_col, min(ws.max_column, 200) + 1):
        v = ws.cell(row=header_row, column=c).value
        if is_blank(v):
            continue
        headers.append((c, str(v).strip()))

    if not headers:
        return pd.DataFrame()

    rows = []
    blanks = 0
    for r in range(header_row + 1, min(ws.max_row, header_row + 8000) + 1):
        obra = ws.cell(row=r, column=headers[0][0]).value
        if is_blank(obra):
            blanks += 1
            if blanks >= 8:
                break
            continue
        blanks = 0
        row = {h: ws.cell(row=r, column=c).value for c, h in headers}
        rows.append(row)

    df = pd.DataFrame(rows)
    obra_col = headers[0][1]
    for col in df.columns:
        if col == obra_col:
            continue
        df[col] = df[col].apply(to_float)
    return df
