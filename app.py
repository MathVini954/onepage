from __future__ import annotations

from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .utils import norm, is_blank, to_month, to_float


def load_wb(path: str | Path) -> Workbook:
    path = Path(path)
    keep_vba = path.suffix.lower() == ".xlsm"
    return load_workbook(path, data_only=True, keep_vba=keep_vba)


def sheetnames(wb: Workbook) -> list[str]:
    ignore = {"LEIA-ME", "README", "READ ME"}
    return [s for s in wb.sheetnames if norm(s) not in ignore]


def _find_row_contains(ws: Worksheet, needle: str, col: int = 1, max_row: int = 400) -> int | None:
    n = norm(needle)
    for r in range(1, min(ws.max_row, max_row) + 1):
        if n in norm(ws.cell(r, col).value):
            return r
    return None


def _find_header_row(ws: Worksheet, headers: list[str], max_row: int = 900, max_col: int = 80) -> tuple[int | None, list[str] | None]:
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


def _col_like(row_vals: list[str], must: list[str]) -> int | None:
    must_u = [norm(x) for x in must]
    for i, v in enumerate(row_vals, start=1):
        if all(m in v for m in must_u):
            return i
    return None


def read_resumo_financeiro(ws: Worksheet) -> dict[str, float | None]:
    title_row = _find_row_contains(ws, "RESUMO FINANCEIRO", col=1, max_row=200)
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
        key = mapping.get(norm(k), str(k).strip())
        out[key] = to_float(v)
        r += 1

    return out


def read_indice(ws: Worksheet) -> pd.DataFrame:
    hr, row_vals = _find_header_row(ws, ["MÊS", "ÍNDICE PROJETADO"], max_row=600, max_col=40)
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
        data.append({"MÊS": to_month(mes), "ÍNDICE PROJETADO": to_float(ws.cell(r, c_idx).value)})
        r += 1

    return pd.DataFrame(data).dropna(subset=["MÊS"])


def read_financeiro(ws: Worksheet) -> pd.DataFrame:
    hr, row_vals = _find_header_row(ws, ["MÊS", "DESEMBOLSO DO MÊS (R$)", "MEDIDO NO MÊS (R$)"], max_row=800, max_col=80)
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

    return pd.DataFrame(data).dropna(subset=["MÊS"])


def read_prazo(ws: Worksheet) -> pd.DataFrame:
    # Procurar a tabela do prazo pelo título
    title_row = _find_row_contains(ws, "PRAZO", col=1, max_row=900)
    if title_row is None:
        return pd.DataFrame(columns=["MÊS", "PLANEJADO ACUM. (%)", "PLANEJADO MÊS (%)", "PREVISTO MENSAL (%)", "REALIZADO Mês (%)"])

    header_row = title_row + 1
    row_vals = [norm(ws.cell(header_row, c).value) for c in range(1, min(ws.max_column, 80) + 1)]

    c_mes = _col_idx(row_vals, "MÊS")
    c_plan_ac = _col_like(row_vals, ["PLANEJADO", "ACUM"])
    c_plan_m = _col_like(row_vals, ["PLANEJADO", "MÊS"]) or _col_like(row_vals, ["PLANEJADO", "MES"])
    c_real_m = _col_like(row_vals, ["REALIZADO", "MÊS"]) or _col_like(row_vals, ["REALIZADO", "MES"])
    c_prev_m = (
        _col_like(row_vals, ["PREVISTO", "MENSAL"])
        or _col_like(row_vals, ["PREVISTO", "MÊS"])
        or _col_like(row_vals, ["PREVISTO", "MES"])
        or _col_like(row_vals, ["COMPROMET", "MÊS"])   # caso você use "Comprometido mês" como previsto
        or _col_like(row_vals, ["COMPROMET", "MES"])
    )

    # Se não achou previsto pelo header, NÃO inventa (evita pegar acumulado sem header)
    # -> só lê se existir header.

    if c_mes is None or c_plan_m is None or c_real_m is None:
        return pd.DataFrame(columns=["MÊS", "PLANEJADO ACUM. (%)", "PLANEJADO MÊS (%)", "PREVISTO MENSAL (%)", "REALIZADO Mês (%)"])

    data = []
    r = header_row + 1
    while r <= ws.max_row:
        mes = ws.cell(r, c_mes).value
        if is_blank(mes):
            break

        row = {
            "MÊS": to_month(mes),
            "PLANEJADO MÊS (%)": to_float(ws.cell(r, c_plan_m).value),
            "REALIZADO Mês (%)": to_float(ws.cell(r, c_real_m).value),
        }

        if c_plan_ac is not None:
            row["PLANEJADO ACUM. (%)"] = to_float(ws.cell(r, c_plan_ac).value)

        if c_prev_m is not None:
            row["PREVISTO MENSAL (%)"] = to_float(ws.cell(r, c_prev_m).value)

        data.append(row)
        r += 1

    return pd.DataFrame(data).dropna(subset=["MÊS"])


def read_acrescimos_economias(ws: Worksheet) -> tuple[pd.DataFrame, pd.DataFrame]:
    base_row = None
    for r in range(1, min(ws.max_row, 2000) + 1):
        a = norm(ws.cell(r, 1).value)
        g = norm(ws.cell(r, 7).value)
        if ("ACRÉSCIM" in a or "ACRESCIM" in a) and ("ECONOM" in g):
            base_row = r
            break

    empty = pd.DataFrame(columns=["DESCRIÇÃO", "ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO", "JUSTIFICATIVAS"])
    if base_row is None:
        return empty, empty

    header_row = base_row + 1

    def read_side(start_col: int) -> pd.DataFrame:
        header_map = {}
        for c in range(start_col, start_col + 12):
            h = norm(ws.cell(header_row, c).value)
            if h in ("DESCRIÇÃO", "DESCRICAO"):
                header_map["DESCRIÇÃO"] = c
            elif h in ("ORÇAMENTO INICIAL", "ORCAMENTO INICIAL"):
                header_map["ORÇAMENTO INICIAL"] = c
            elif h in ("ORÇAMENTO REAJUSTADO", "ORCAMENTO REAJUSTADO"):
                header_map["ORÇAMENTO REAJUSTADO"] = c
            elif h == "CUSTO FINAL":
                header_map["CUSTO FINAL"] = c
            elif h in ("VARIAÇÃO", "VARIACAO"):
                header_map["VARIAÇÃO"] = c
            elif h in ("JUSTIFICATIVAS", "JUSTIFICATIVA"):
                header_map["JUSTIFICATIVAS"] = c

        if "DESCRIÇÃO" not in header_map or "VARIAÇÃO" not in header_map:
            return empty.copy()

        data = []
        r = header_row + 1
        while r <= ws.max_row:
            desc = ws.cell(r, header_map["DESCRIÇÃO"]).value
            if is_blank(desc):
                break

            data.append(
                {
                    "DESCRIÇÃO": str(desc).strip(),
                    "ORÇAMENTO INICIAL": to_float(ws.cell(r, header_map.get("ORÇAMENTO INICIAL", -1)).value) if "ORÇAMENTO INICIAL" in header_map else None,
                    "ORÇAMENTO REAJUSTADO": to_float(ws.cell(r, header_map.get("ORÇAMENTO REAJUSTADO", -1)).value) if "ORÇAMENTO REAJUSTADO" in header_map else None,
                    "CUSTO FINAL": to_float(ws.cell(r, header_map.get("CUSTO FINAL", -1)).value) if "CUSTO FINAL" in header_map else None,
                    "VARIAÇÃO": to_float(ws.cell(r, header_map["VARIAÇÃO"]).value),
                    "JUSTIFICATIVAS": (
                        str(ws.cell(r, header_map["JUSTIFICATIVAS"]).value).strip()
                        if "JUSTIFICATIVAS" in header_map and not is_blank(ws.cell(r, header_map["JUSTIFICATIVAS"]).value)
                        else ""
                    ),
                }
            )
            r += 1

        df = pd.DataFrame(data)
        return df if not df.empty else empty.copy()

    return read_side(1), read_side(7)
