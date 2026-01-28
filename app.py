from __future__ import annotations

import html
import unicodedata
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode

from src.excel_reader import (
    load_wb,
    sheetnames,
    read_resumo_financeiro,
    read_indice,
    read_financeiro,
    read_prazo,
    read_acrescimos_economias,
)
from src.logos import find_logo_path
from src.utils import fmt_brl


# ============================================================
# Config
# ============================================================
st.set_page_config(page_title="Controle Prazo e Custo", layout="wide")
LOGOS_DIR = "assets/logos"

GOOD = "#22c55e"
BAD = "#ef4444"
BLUE = "#3b82f6"


# ============================================================
# Excel único (sem upload)
# ============================================================
def find_default_excel() -> Path | None:
    for name in ["Excel.xlsm", "Excel.xlsx", "excel.xlsm", "excel.xlsx"]:
        p = Path(name)
        if p.exists():
            return p
    return None


excel_path = find_default_excel()
if excel_path is None:
    st.error("Não achei **Excel.xlsm** (ou Excel.xlsx) na raiz do projeto.")
    st.stop()

wb = load_wb(excel_path)
obras = sheetnames(wb)
if not obras:
    st.error("Nenhuma aba de obra encontrada no Excel.")
    st.stop()


# ============================================================
# Sidebar
# ============================================================
st.sidebar.title("Controle de Prazo e Custo")
obra = st.sidebar.selectbox("Obra (aba)", obras, index=0)

top_opt = st.sidebar.selectbox("Mostrar Top", ["5", "10", "Todas"], index=0)
top_n = None if top_opt == "Todas" else int(top_opt)

st.sidebar.markdown("---")
dark_mode = st.sidebar.toggle("Modo escuro", value=True)
st.sidebar.caption(f"Tema: {'Escuro' if dark_mode else 'Claro'}")

debug = st.sidebar.toggle("Debug", value=False)

ws = wb[obra]


# ============================================================
# Tema
# ============================================================
def hex_to_rgb(h: str) -> tuple[int, int, int]:
    h = h.strip().lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def rgba(hex_color: str, a: float) -> str:
    r, g, b = hex_to_rgb(hex_color)
    return f"rgba({r},{g},{b},{a})"


if dark_mode:
    PALETTE = {
        "bg": "#0b1220",
        "sidebar_bg": "#0b1220",
        "text": "#e5e7eb",
        "muted": "#9aa4b2",
        "card": "rgba(255,255,255,0.04)",
        "border": "rgba(255,255,255,0.10)",
        "track": "rgba(255,255,255,0.10)",
        "good": GOOD,
        "bad": BAD,
        "good_bg": rgba(GOOD, 0.10),
        "good_border": rgba(GOOD, 0.28),
        "bad_bg": rgba(BAD, 0.10),
        "bad_border": rgba(BAD, 0.28),
        "bar_des": rgba(BLUE, 0.85),
        "bar_med": rgba(GOOD, 0.85),
        "plotly_template": "plotly_dark",
        "good_grad": [[0, rgba(GOOD, 0.20)], [1, rgba(GOOD, 1.0)]],
        "bad_grad": [[0, rgba(BAD, 0.20)], [1, rgba(BAD, 1.0)]],
        "planned_bar": rgba(BLUE, 0.35),
        "real_bar": rgba(BLUE, 0.95),
    }
else:
    PALETTE = {
        "bg": "#f7f8fc",
        "sidebar_bg": "#ffffff",
        "text": "#0f172a",
        "muted": "#475569",
        "card": "rgba(255,255,255,0.92)",
        "border": "rgba(15,23,42,0.10)",
        "track": "rgba(15,23,42,0.10)",
        "good": GOOD,
        "bad": BAD,
        "good_bg": rgba(GOOD, 0.12),
        "good_border": rgba(GOOD, 0.28),
        "bad_bg": rgba(BAD, 0.10),
        "bad_border": rgba(BAD, 0.25),
        "bar_des": rgba(BLUE, 0.80),
        "bar_med": rgba(GOOD, 0.80),
        "plotly_template": "plotly_white",
        "good_grad": [[0, rgba(GOOD, 0.18)], [1, rgba(GOOD, 0.95)]],
        "bad_grad": [[0, rgba(BAD, 0.18)], [1, rgba(BAD, 0.95)]],
        "planned_bar": rgba(BLUE, 0.22),
        "real_bar": rgba(BLUE, 0.85),
    }

PLOTLY_TEMPLATE = PALETTE["plotly_template"]

# CSS (blindado)
st.markdown(
    f"""
<style>
  html, body, [data-testid="stAppViewContainer"], .stApp {{
    background: {PALETTE["bg"]} !important;
  }}

  header[data-testid="stHeader"] {{
    background: {PALETTE["bg"]} !important;
    border-bottom: 1px solid {PALETTE["border"]} !important;
  }}

  section[data-testid="stSidebar"] {{
    display: block !important;
    visibility: visible !important;
    opacity: 1 !important;
  }}
  section[data-testid="stSidebar"] > div {{
    background: {PALETTE["sidebar_bg"]} !important;
    border-right: 1px solid {PALETTE["border"]} !important;
  }}

  [data-testid="collapsedControl"] {{
    display: block !important;
    visibility: visible !important;
    opacity: 1 !important;
  }}

  .block-container {{
    padding-top: 1.25rem;
    padding-bottom: 2rem;
  }}

  @media (max-width: 900px){
    .block-container {{ padding-left: 0.8rem; padding-right: 0.8rem; }}
    h1 {{ font-size: 1.6rem !important; }}
    h2 {{ font-size: 1.25rem !important; }}
    h3 {{ font-size: 1.1rem !important; }}
    [data-testid="stSidebar"] {{ width: 85vw !important; }}
  }}
</style>
""",
    unsafe_allow_html=True,
)


def apply_plotly_theme(fig: go.Figure) -> go.Figure:
    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10, r=10, t=10, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    return fig


# ============================================================
# Helpers
# ============================================================
def fmt_brl_no_dec(v: float) -> str:
    s = f"{float(v):,.0f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def brl_compact(v: float | None) -> str:
    if v is None:
        return "—"
    n = float(v)
    a = abs(n)
    if a >= 1_000_000_000:
        return f"R$ {n/1_000_000_000:.2f} bi".replace(".", ",")
    if a >= 1_000_000:
        return f"R$ {n/1_000_000:.2f} mi".replace(".", ",")
    if a >= 1_000:
        return f"R$ {n/1_000:.2f} mil".replace(".", ",")
    return fmt_brl(n)


def to_ratio(x) -> float | None:
    """Aceita 0-1 ou 0-100 e converte para 0-1."""
    if x is None:
        return None
    try:
        v = float(x)
    except Exception:
        return None
    return v if v <= 1.5 else (v / 100.0)


def clamp01(v: float | None) -> float:
    if v is None:
        return 0.0
    return max(0.0, min(1.0, float(v)))


def pct(v_ratio: float | None) -> str:
    if v_ratio is None:
        return "—"
    return f"{v_ratio*100:.1f}%".replace(".", ",")


def kpi_card_money(label: str, value: float | None):
    st.markdown(
        f"""
<div style="border:1px solid {PALETTE["border"]}; border-radius:14px; padding:12px 14px; background:{PALETTE["card"]}; height:92px;">
  <div style="font-size:12px; color:{PALETTE["muted"]}; margin-bottom:6px;">{html.escape(label)}</div>
  <div style="font-size:24px; font-weight:900; line-height:1.05; color:{PALETTE["text"]};">{html.escape(brl_compact(value))}</div>
  <div style="font-size:11px; color:{PALETTE["muted"]}; margin-top:6px;">{html.escape(fmt_brl(value))}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def kpi_card_money_highlight(label: str, value: float | None, value_color: str, subtitle: str = ""):
    st.markdown(
        f"""
<div style="border:1px solid {PALETTE["border"]}; border-radius:14px; padding:12px 14px; background:{PALETTE["card"]}; height:92px;">
  <div style="font-size:12px; color:{PALETTE["muted"]}; margin-bottom:6px;">{html.escape(label)}</div>
  <div style="font-size:24px; font-weight:900; line-height:1.05; color:{value_color};">{html.escape(brl_compact(value))}</div>
  <div style="font-size:11px; color:{PALETTE["muted"]}; margin-top:6px;">{html.escape(subtitle) if subtitle else html.escape(fmt_brl(value))}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def kpi_card_pct(label: str, value_ratio: float | None, sub: str = ""):
    st.markdown(
        f"""
<div style="border:1px solid {PALETTE["border"]}; border-radius:14px; padding:12px 14px; background:{PALETTE["card"]}; height:92px;">
  <div style="font-size:12px; color:{PALETTE["muted"]}; margin-bottom:6px;">{html.escape(label)}</div>
  <div style="font-size:24px; font-weight:900; line-height:1.05; color:{PALETTE["text"]};">{html.escape(pct(value_ratio))}</div>
  <div style="font-size:11px; color:{PALETTE["muted"]}; margin-top:6px;">{html.escape(sub)}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def kpi_card_index(label: str, idx: float | None, month_label: str):
    if idx is None:
        val = "—"
        color = PALETTE["muted"]
    else:
        val = f"{idx:.3f}".replace(".", ",")
        if idx > 1.0:
            color = PALETTE["bad"]
        elif idx < 1.0:
            color = PALETTE["good"]
        else:
            color = PALETTE["text"]

    st.markdown(
        f"""
<div style="border:1px solid {PALETTE["border"]}; border-radius:14px; padding:12px 14px; background:{PALETTE["card"]}; height:92px;">
  <div style="font-size:12px; color:{PALETTE["muted"]}; margin-bottom:6px;">{html.escape(label)}</div>
  <div style="font-size:24px; font-weight:900; line-height:1.05; color:{color};">{html.escape(val)}</div>
  <div style="font-size:11px; color:{PALETTE["muted"]}; margin-top:6px;">{html.escape(month_label)}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def progress_card(real_ratio: float | None, planned_ratio: float | None, ref_month_label: str):
    real_ratio = clamp01(real_ratio)
    planned_ratio = clamp01(planned_ratio)

    real_pct = real_ratio * 100
    planned_pct = planned_ratio * 100

    st.markdown(
        f"""
<div style="border:1px solid {PALETTE["border"]}; background:{PALETTE["card"]}; border-radius:16px; padding:14px 16px;">
  <div style="display:flex; justify-content:space-between; align-items:center;">
    <div style="font-size:12px; color:{PALETTE["text"]}; font-weight:900;">Obra vs. Planejado (acum.)</div>
    <div style="font-size:12px; color:{PALETTE["muted"]};">{html.escape(ref_month_label)}</div>
  </div>

  <div style="margin-top:12px; display:flex; justify-content:space-between; align-items:flex-end;">
    <div>
      <div style="font-size:12px; color:{PALETTE["muted"]};">Real</div>
      <div style="font-size:28px; font-weight:900; line-height:1; color:{PALETTE["text"]};">{real_pct:.0f}%</div>
    </div>
    <div style="text-align:right;">
      <div style="font-size:12px; color:{PALETTE["muted"]};">Planejado</div>
      <div style="font-size:16px; font-weight:900; color:{PALETTE["text"]};">{planned_pct:.0f}%</div>
    </div>
  </div>

  <div style="margin-top:12px;">
    <div style="height:10px; background:{PALETTE["track"]}; border-radius:999px; position:relative;">
      <div style="width:{planned_pct:.2f}%; height:10px; background:{PALETTE["planned_bar"]}; border-radius:999px;"></div>
      <div style="width:{real_pct:.2f}%; height:10px; background:{PALETTE["real_bar"]}; border-radius:999px; position:absolute; top:0; left:0;"></div>
    </div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


def build_rows(items: list[tuple[str, float]], color: str, prefix: str = "") -> str:
    rows = ""
    for desc, val in items:
        desc = html.escape(str(desc))
        val_show = fmt_brl_no_dec(abs(val))
        rows += f"""
<div style="display:flex; justify-content:space-between; align-items:center; padding:10px 0; border-top:1px solid {PALETTE["border"]};">
  <div style="font-size:13px; font-weight:600; color:{PALETTE["text"]};">{desc}</div>
  <div style="font-size:13px; font-weight:800; color:{color};">{prefix}{val_show}</div>
</div>
"""
    return rows


def card_resumo(title: str, icon: str, rows_html: str, border: str, bg: str) -> str:
    return f"""
<div style="border:1px solid {border}; background:{bg}; border-radius:16px; padding:14px 16px;">
  <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
    <div style="font-size:12px; color:{PALETTE["text"]}; font-weight:900; letter-spacing:0.3px;">{html.escape(title)}</div>
    <div style="font-size:12px;">{icon}</div>
  </div>
  {rows_html if rows_html else f'<div style="color:{PALETTE["muted"]}; font-size:12px;">Sem dados</div>'}
</div>
"""


def styled_dataframe(df: pd.DataFrame):
    if df is None or df.empty:
        st.info("Sem dados.")
        return
    tbl = df.copy()
    money_cols = ["ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO", "VARIAÇÃO (R$)"]
    for c in money_cols:
        if c in tbl.columns:
            tbl[c] = pd.to_numeric(tbl[c], errors="coerce")
    fmt_map = {c: fmt_brl for c in money_cols if c in tbl.columns}
    st.dataframe(tbl.style.format(fmt_map), use_container_width=True, hide_index=True)


def sum_abs_column(df: pd.DataFrame, col: str) -> float:
    if df is None or df.empty or col not in df.columns:
        return 0.0
    s = pd.to_numeric(df[col], errors="coerce").dropna()
    return float(s.abs().sum()) if not s.empty else 0.0


# ============================================================
# Leitura ORÇAMENTO_RESUMO (BI) - dentro do app.py
# ============================================================
def _norm(s: str) -> str:
    s = "" if s is None else str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return " ".join(s.upper().split())


def read_orcamento_resumo_from_wb(workbook) -> pd.DataFrame | None:
    # tenta achar aba
    target = None
    for name in workbook.sheetnames:
        if _norm(name) in ["ORCAMENTO_RESUMO", "ORÇAMENTO_RESUMO", "ORCAMENTO RESUMO", "ORÇAMENTO RESUMO"]:
            target = name
            break
    if target is None:
        return None

    ws_r = workbook[target]

    # achar header row com "OBRA"
    header_row = None
    max_scan = 250
    for r in range(1, max_scan + 1):
        row_vals = [ws_r.cell(row=r, column=c).value for c in range(1, 80)]
        if any(_norm(v) == "OBRA" for v in row_vals if v is not None):
            header_row = r
            break
    if header_row is None:
        return None

    # headers até último não vazio
    headers = []
    last_c = 1
    for c in range(1, 200):
        v = ws_r.cell(row=header_row, column=c).value
        if v is None and c > 10:
            # heurística: depois de um bom trecho, parar
            pass
        if v is not None:
            last_c = c
        headers.append("" if v is None else str(v).strip())

    headers = headers[:last_c]
    # limpar headers vazios
    headers = [h if h else f"COL_{i+1}" for i, h in enumerate(headers)]

    data = []
    for r in range(header_row + 1, header_row + 5000):
        first = ws_r.cell(row=r, column=1).value
        if first is None:
            # para quando encontrar várias linhas vazias
            # (mas tolera buracos pequenos)
            empty_row = True
            for c in range(1, last_c + 1):
                if ws_r.cell(row=r, column=c).value is not None:
                    empty_row = False
                    break
            if empty_row:
                break

        row = [ws_r.cell(row=r, column=c).value for c in range(1, last_c + 1)]
        data.append(row)

    df = pd.DataFrame(data, columns=headers)
    # drop linhas sem OBRA
    if "OBRA" in df.columns:
        df["OBRA"] = df["OBRA"].astype(str).str.strip()
        df = df[df["OBRA"].astype(str).str.strip().ne("")].copy()
    return df


df_orc_resumo = read_orcamento_resumo_from_wb(wb)


# ============================================================
# Header + logo
# ============================================================
colL, colR = st.columns([1, 5])
with colL:
    logo_path = find_logo_path(obra, LOGOS_DIR)
    if logo_path:
        st.image(logo_path, use_container_width=True)
with colR:
    st.title(f"Controle de Prazo e Custo — {obra}")

st.divider()

if debug:
    st.write("Arquivo:", excel_path.name)
    st.write("Abas:", obras)
    st.write("Resumo Obras:", "OK" if df_orc_resumo is not None else "Não achou ORÇAMENTO_RESUMO")


# ============================================================
# Ler dados (obra selecionada)
# ============================================================
resumo = read_resumo_financeiro(ws)
df_idx = read_indice(ws)
df_fin = read_financeiro(ws)
df_prazo = read_prazo(ws)
df_acres, df_econ = read_acrescimos_economias(ws)

total_economias = sum_abs_column(df_econ, "VARIAÇÃO")
total_acrescimos = sum_abs_column(df_acres, "VARIAÇÃO")
desvio_liquido = total_acrescimos - total_economias


# ============================================================
# Índice do mês (último)
# ============================================================
idx_last = None
idx_month_label = "—"
if not df_idx.empty and "ÍNDICE PROJETADO" in df_idx.columns:
    df_idx2 = df_idx.dropna(subset=["MÊS"]).sort_values("MÊS")
    df_idx2["ÍNDICE PROJETADO"] = pd.to_numeric(df_idx2["ÍNDICE PROJETADO"], errors="coerce")
    df_idx2 = df_idx2.dropna(subset=["ÍNDICE PROJETADO"])
    if not df_idx2.empty:
        idx_last = float(df_idx2["ÍNDICE PROJETADO"].iloc[-1])
        m = df_idx2["MÊS"].iloc[-1]
        idx_month_label = pd.to_datetime(m).strftime("%b/%Y").lower()


# ============================================================
# Prazo (corte no último mês preenchido)
# ============================================================
temp = pd.DataFrame()
ref_month_label = "—"

k_real_acum = None
k_plan_acum = None
k_prev_acum = None
k_real_m = None
k_plan_m = None
k_prev_m = None
k_ader_acc = None

planned_m = []
previsto_m = []
real_m = []

planned_acum = []
previsto_acum = []
real_acum = []

if not df_prazo.empty and "MÊS" in df_prazo.columns:
    temp = df_prazo.copy().dropna(subset=["MÊS"]).sort_values("MÊS").reset_index(drop=True)

    temp["PLANEJADO_M"] = (
        temp["PLANEJADO MÊS (%)"].apply(to_ratio) if "PLANEJADO MÊS (%)" in temp.columns else pd.NA
    )
    temp["PREVISTO_M"] = (
        temp["PREVISTO MENSAL (%)"].apply(to_ratio) if "PREVISTO MENSAL (%)" in temp.columns else pd.NA
    )
    temp["REAL_M"] = (
        temp["REALIZADO Mês (%)"].apply(to_ratio) if "REALIZADO Mês (%)" in temp.columns else pd.NA
    )

    if "PLANEJADO ACUM. (%)" in temp.columns:
        temp["PLANEJADO_ACUM"] = temp["PLANEJADO ACUM. (%)"].apply(to_ratio)
    else:
        temp["PLANEJADO_ACUM"] = pd.to_numeric(temp["PLANEJADO_M"], errors="coerce").cumsum()

    temp["PREVISTO_ACUM"] = pd.to_numeric(temp["PREVISTO_M"], errors="coerce").cumsum()
    temp["REAL_ACUM"] = pd.to_numeric(temp["REAL_M"], errors="coerce").cumsum()

    last_idxs = []
    for col in ["PLANEJADO_M", "PREVISTO_M", "REAL_M", "PLANEJADO_ACUM", "PREVISTO_ACUM", "REAL_ACUM"]:
        idx = temp[col].last_valid_index()
        if idx is not None:
            last_idxs.append(idx)
    if last_idxs:
        temp = temp.iloc[: max(last_idxs) + 1].copy()

    def series_stop_at_last(s: pd.Series) -> list[float | None]:
        s = pd.to_numeric(s, errors="coerce")
        last = s.last_valid_index()
        if last is None:
            return [None] * len(s)
        out = s.copy()
        for i in range(last + 1, len(out)):
            out.iloc[i] = pd.NA
        return [None if pd.isna(v) else float(v) for v in out.tolist()]

    planned_m = [None if v is None else v * 100 for v in series_stop_at_last(temp["PLANEJADO_M"])]
    previsto_m = [None if v is None else v * 100 for v in series_stop_at_last(temp["PREVISTO_M"])]
    real_m = [None if v is None else v * 100 for v in series_stop_at_last(temp["REAL_M"])]

    planned_acum = [None if v is None else v * 100 for v in series_stop_at_last(temp["PLANEJADO_ACUM"])]
    previsto_acum = [None if v is None else v * 100 for v in series_stop_at_last(temp["PREVISTO_ACUM"])]
    real_acum = [None if v is None else v * 100 for v in series_stop_at_last(temp["REAL_ACUM"])]

    last_real = pd.to_numeric(temp["REAL_M"], errors="coerce").last_valid_index()
    if last_real is not None:
        m = pd.to_datetime(temp.loc[last_real, "MÊS"])
        ref_month_label = m.strftime("%b/%Y").lower()

        k_real_m = temp.loc[last_real, "REAL_M"]
        k_plan_m = temp.loc[last_real, "PLANEJADO_M"]
        k_prev_m = temp.loc[last_real, "PREVISTO_M"]

        k_real_acum = temp.loc[last_real, "REAL_ACUM"]
        k_plan_acum = temp.loc[last_real, "PLANEJADO_ACUM"]
        k_prev_acum = temp.loc[last_real, "PREVISTO_ACUM"]

        if pd.notna(k_plan_acum) and float(k_plan_acum) != 0:
            k_ader_acc = (float(k_real_acum or 0) / float(k_plan_acum)) * 100


# ============================================================
# Tabs
# ============================================================
tab_dash, tab_just, tab_resumo = st.tabs(["Dashboard", "Justificativas", "Resumo Obras"])


# ============================================================
# TAB Dashboard
# ============================================================
with tab_dash:
    row1 = st.columns(4)
    with row1[0]:
        kpi_card_index("Índice do mês", idx_last, idx_month_label)
    with row1[1]:
        kpi_card_money("Orç. Inicial", resumo.get("ORÇAMENTO INICIAL (R$)"))
    with row1[2]:
        kpi_card_money("Orç. Reajust.", resumo.get("ORÇAMENTO REAJUSTADO (R$)"))
    with row1[3]:
        kpi_card_money("Desembolso Acum.", resumo.get("DESEMBOLSO ACUMULADO (R$)"))

    st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)

    row2 = st.columns(4)
    with row2[0]:
        kpi_card_money("A Pagar", resumo.get("A PAGAR (R$)"))
    with row2[1]:
        kpi_card_money("Saldo a Incorrer", resumo.get("SALDO A INCORRER (R$)"))
    with row2[2]:
        kpi_card_money("Custo Final", resumo.get("CUSTO FINAL (R$)"))
    with row2[3]:
        kpi_card_money("Variação", resumo.get("VARIAÇÃO (R$)"))

    st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)

    row3 = st.columns(3)
    with row3[0]:
        kpi_card_money_highlight("Total Economias (mês)", total_economias, PALETTE["good"])
    with row3[1]:
        kpi_card_money_highlight("Total Acréscimos (mês)", total_acrescimos, PALETTE["bad"])
    with row3[2]:
        color_desvio = PALETTE["bad"] if desvio_liquido > 0 else PALETTE["good"]
        kpi_card_money_highlight("Desvio Líquido (Acrésc. − Econ.)", desvio_liquido, color_desvio)

    st.divider()

    left, right = st.columns([2.2, 1])

    with left:
        g1, g2 = st.columns(2)

        with g1:
            st.subheader("Índice Projetado (baseline 1,000)")
            if df_idx.empty:
                st.info("Sem dados do índice.")
            else:
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=df_idx["MÊS"], y=df_idx["ÍNDICE PROJETADO"], mode="lines+markers", name="Índice"))
                fig.add_hline(y=1.0, line_dash="dash", line_width=1)
                fig.update_layout(height=320)
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

        with g2:
            st.subheader("Desembolso x Medido (mês a mês)")
            if df_fin.empty:
                st.info("Sem dados financeiros.")
            else:
                fig = go.Figure()
                fig.add_trace(go.Bar(x=df_fin["MÊS"], y=df_fin["DESEMBOLSO DO MÊS (R$)"], name="Desembolso", marker_color=PALETTE["bar_des"]))
                fig.add_trace(go.Bar(x=df_fin["MÊS"], y=df_fin["MEDIDO NO MÊS (R$)"], name="Medido", marker_color=PALETTE["bar_med"]))
                fig.update_layout(barmode="group", height=320)
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

        st.subheader("Prazo — Curva S (Acumulado) + Curva Mensal")
        if temp.empty:
            st.info("Sem dados de prazo.")
        else:
            st.markdown("### KPIs de Prazo")
            r1 = st.columns(3)
            with r1[0]:
                kpi_card_pct("Realizado acumulado", k_real_acum, f"ref: {ref_month_label}")
            with r1[1]:
                kpi_card_pct("Planejado acumulado", k_plan_acum, f"ref: {ref_month_label}")
            with r1[2]:
                ader_ratio = (k_ader_acc / 100) if k_ader_acc is not None else None
                kpi_card_pct("Aderência acumulada", ader_ratio, "(Real acum ÷ Plan acum)")

            st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

            r2 = st.columns(3)
            with r2[0]:
                kpi_card_pct("Realizado mensal", k_real_m, f"ref: {ref_month_label}")
            with r2[1]:
                kpi_card_pct("Previsto mensal", k_prev_m, f"ref: {ref_month_label}")
            with r2[2]:
                kpi_card_pct("Planejado mensal", k_plan_m, f"ref: {ref_month_label}")

            x = temp["MÊS"].tolist()
            t1, t2 = st.tabs(["Curva S (Acumulado)", "Curva Mensal (Individual)"])

            with t1:
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=x, y=planned_acum, mode="lines+markers", name="Planejado acum. (%)"))
                fig.add_trace(go.Scatter(x=x, y=previsto_acum, mode="lines+markers", name="Previsto acum. (%)"))
                fig.add_trace(go.Scatter(x=x, y=real_acum, mode="lines+markers", name="Realizado acum. (%)"))
                fig.update_layout(height=320, yaxis_title="%")
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

            with t2:
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=x, y=planned_m, mode="lines+markers", name="Planejado mês (%)"))
                fig.add_trace(go.Scatter(x=x, y=previsto_m, mode="lines+markers", name="Previsto mês (%)"))
                fig.add_trace(go.Scatter(x=x, y=real_m, mode="lines+markers", name="Realizado mês (%)"))
                fig.update_layout(height=320, yaxis_title="% (mensal)")
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

    with right:
        econ_items: list[tuple[str, float]] = []
        acres_items: list[tuple[str, float]] = []

        if not df_econ.empty and "VARIAÇÃO" in df_econ.columns:
            econ_sorted = df_econ.copy()
            econ_sorted["__v"] = pd.to_numeric(econ_sorted["VARIAÇÃO"], errors="coerce")
            econ_sorted = econ_sorted.dropna(subset=["__v"])
            econ_sorted["__abs"] = econ_sorted["__v"].abs()
            econ_sorted = econ_sorted.sort_values("__abs", ascending=False).head(3)
            for _, r in econ_sorted.iterrows():
                econ_items.append((str(r.get("DESCRIÇÃO", "")), float(r.get("__v", 0) or 0)))

        if not df_acres.empty and "VARIAÇÃO" in df_acres.columns:
            acres_sorted = df_acres.copy()
            acres_sorted["__v"] = pd.to_numeric(acres_sorted["VARIAÇÃO"], errors="coerce")
            acres_sorted = acres_sorted.dropna(subset=["__v"])
            acres_sorted["__abs"] = acres_sorted["__v"].abs()
            acres_sorted = acres_sorted.sort_values("__abs", ascending=False).head(3)
            for _, r in acres_sorted.iterrows():
                acres_items.append((str(r.get("DESCRIÇÃO", "")), float(r.get("__v", 0) or 0)))

        econ_rows = build_rows(econ_items, color=PALETTE["good"], prefix="")
        acres_rows = build_rows(acres_items, color=PALETTE["bad"], prefix="- ")

        st.markdown(card_resumo("PRINCIPAIS ECONOMIAS", "✅", econ_rows, PALETTE["good_border"], PALETTE["good_bg"]), unsafe_allow_html=True)
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        st.markdown(card_resumo("DESVIOS DO MÊS", "⚠️", acres_rows, PALETTE["bad_border"], PALETTE["bad_bg"]), unsafe_allow_html=True)

        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        progress_card(k_real_acum, k_plan_acum, ref_month_label)

    st.divider()

    st.subheader("Detalhamento — Tabelas completas (com barras em degradê)")
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### ACRÉSCIMOS / DESVIOS")
        if df_acres.empty:
            st.info("Sem dados.")
        else:
            show = df_acres.copy()
            show["VARIAÇÃO"] = pd.to_numeric(show["VARIAÇÃO"], errors="coerce")
            show = show.dropna(subset=["VARIAÇÃO"])
            show["__abs"] = show["VARIAÇÃO"].abs()
            show = show.sort_values("__abs", ascending=False)
            show_top = show.head(top_n) if top_n is not None else show

            top_bar = show.head(10).iloc[::-1]
            vals = top_bar["VARIAÇÃO"].abs()
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=vals,
                y=top_bar["DESCRIÇÃO"],
                orientation="h",
                marker=dict(color=vals, colorscale=PALETTE["bad_grad"], showscale=False),
                name="R$",
            ))
            fig.update_layout(height=340, xaxis_title="R$")
            st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

            with st.expander("Ver tabela (Acréscimos)"):
                styled_dataframe(show_top.drop(columns=["__abs"], errors="ignore"))

    with c2:
        st.markdown("### ECONOMIAS")
        if df_econ.empty:
            st.info("Sem dados.")
        else:
            show = df_econ.copy()
            show["VARIAÇÃO"] = pd.to_numeric(show["VARIAÇÃO"], errors="coerce")
            show = show.dropna(subset=["VARIAÇÃO"])
            show["__abs"] = show["VARIAÇÃO"].abs()
            show = show.sort_values("__abs", ascending=False)
            show_top = show.head(top_n) if top_n is not None else show

            top_bar = show.head(10).iloc[::-1]
            vals = top_bar["VARIAÇÃO"].abs()
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=vals,
                y=top_bar["DESCRIÇÃO"],
                orientation="h",
                marker=dict(color=vals, colorscale=PALETTE["good_grad"], showscale=False),
                name="R$",
            ))
            fig.update_layout(height=340, xaxis_title="R$")
            st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

            with st.expander("Ver tabela (Economias)"):
                styled_dataframe(show_top.drop(columns=["__abs"], errors="ignore"))


# ============================================================
# TAB Justificativas
# ============================================================
with tab_just:
    st.subheader("Justificativas — Top 5 Economias e Top 5 Desvios")

    def list_just(df: pd.DataFrame, title: str, color: str, topk: int = 5):
        st.markdown(
            f"""
<div style="border:1px solid {PALETTE["border"]}; background:{PALETTE["card"]}; border-radius:16px; padding:14px 16px;">
  <div style="font-size:12px; color:{PALETTE["muted"]}; font-weight:900; margin-bottom:10px;">{html.escape(title)}</div>
""",
            unsafe_allow_html=True,
        )

        if df is None or df.empty:
            st.markdown(f"<div style='color:{PALETTE['muted']}; font-size:12px;'>Sem dados</div></div>", unsafe_allow_html=True)
            return

        tempj = df.copy()
        tempj["VARIAÇÃO"] = pd.to_numeric(tempj.get("VARIAÇÃO", 0), errors="coerce").fillna(0)
        tempj["__abs"] = tempj["VARIAÇÃO"].abs()
        tempj = tempj.sort_values("__abs", ascending=False).head(topk)

        for _, r in tempj.iterrows():
            desc = str(r.get("DESCRIÇÃO", "")).strip()
            var = float(r.get("VARIAÇÃO", 0) or 0)
            just = str(r.get("JUSTIFICATIVAS", "") or "").strip() or "—"

            st.markdown(
                f"""
<div style="padding:10px 0; border-top:1px solid {PALETTE["border"]};">
  <div style="display:flex; justify-content:space-between; align-items:center;">
    <div style="font-size:13px; font-weight:800; color:{PALETTE["text"]};">{html.escape(desc)}</div>
    <div style="font-size:13px; font-weight:900; color:{color};">{html.escape(fmt_brl_no_dec(abs(var)))}</div>
  </div>
  <div style="margin-top:6px; font-size:12px; color:{PALETTE["muted"]}; line-height:1.35;">
    {html.escape(just)}
  </div>
</div>
""",
                unsafe_allow_html=True,
            )

        st.markdown("</div>", unsafe_allow_html=True)

    a, b = st.columns(2)
    with a:
        list_just(df_econ, "TOP 5 — ECONOMIAS (com justificativa)", PALETTE["good"])
    with b:
        list_just(df_acres, "TOP 5 — DESVIOS / ACRÉSCIMOS (com justificativa)", PALETTE["bad"])


# ============================================================
# TAB Resumo Obras (BI)
# ============================================================
with tab_resumo:
    st.subheader("Resumo das Obras — ORÇAMENTO_RESUMO")

    if df_orc_resumo is None or df_orc_resumo.empty:
        st.info("Não encontrei a aba **ORÇAMENTO_RESUMO** no Excel (ou ela está vazia).")
        st.stop()

    df_show = df_orc_resumo.copy()
    if "OBRA" not in df_show.columns:
        st.error("A aba ORÇAMENTO_RESUMO precisa ter a coluna **OBRA**.")
        st.stop()

    df_show["OBRA"] = df_show["OBRA"].astype(str).str.strip()

    # detectar colunas mês
    PT_MON = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}

    def parse_mes_header(label: str):
        s = str(label).strip().lower()
        s = s.replace(".", "").replace("-", "/")
        if "/" not in s:
            return None
        a, b = s.split("/", 1)
        a = a.strip()[:3]
        b = b.strip()
        if a not in PT_MON:
            return None
        try:
            y = int(b)
            if y < 100:
                y += 2000
        except Exception:
            return None
        return pd.Timestamp(year=y, month=PT_MON[a], day=1)

    month_meta = []
    for c in df_show.columns:
        if _norm(c) == "OBRA":
            continue
        dt = parse_mes_header(c)
        if dt is not None:
            month_meta.append((c, dt))

    month_meta.sort(key=lambda x: x[1])
    month_cols_all = [c for c, _ in month_meta]

    if not month_cols_all:
        st.error("Não encontrei colunas de mês no formato tipo **dez/2025** dentro da ORÇAMENTO_RESUMO.")
        st.stop()

    # converter meses para numérico
    for c in month_cols_all:
        df_show[c] = pd.to_numeric(df_show[c], errors="coerce")

    # último mês com dado
    last_month_col = None
    for c in reversed(month_cols_all):
        if df_show[c].notna().any():
            last_month_col = c
            break
    if last_month_col is None:
        last_month_col = month_cols_all[-1]

    # ----------------------------
    # Controles BI
    # ----------------------------
    ctl1, ctl2, ctl3, ctl4 = st.columns([1.2, 2.0, 2.8, 2.0])

    with ctl1:
        expandir = st.toggle("Expandir meses", value=False)

    with ctl2:
        # período
        if expandir:
            default_start = month_cols_all[max(0, len(month_cols_all) - 6)]
            default_end = month_cols_all[-1]
            periodo = st.select_slider("Período", options=month_cols_all, value=(default_start, default_end))
            start_label, end_label = periodo
        else:
            start_label, end_label = last_month_col, last_month_col

    # filtro obra
    obras_resumo = sorted([o for o in df_show["OBRA"].dropna().unique().tolist() if str(o).strip()])
    with ctl3:
        obras_filtro = st.multiselect("Filtrar obras", obras_resumo, default=obras_resumo)

    with ctl4:
        busca = st.text_input("Buscar obra", value="")

    # aplicar filtros
    df_f = df_show[df_show["OBRA"].isin(obras_filtro)].copy()
    if busca.strip():
        df_f = df_f[df_f["OBRA"].str.contains(busca.strip(), case=False, na=False)].copy()

    # meses no range
    start_idx = month_cols_all.index(start_label)
    end_idx = month_cols_all.index(end_label)
    if start_idx > end_idx:
        start_idx, end_idx = end_idx, start_idx
    months_in_range = month_cols_all[start_idx:end_idx + 1]

    # view: meses + variação do período (sempre calculada)
    df_view = df_f[["OBRA"] + months_in_range].copy()

    first_m = months_in_range[0]
    last_m = months_in_range[-1]
    df_view["VARIAÇÃO (R$)"] = df_view[last_m] - df_view[first_m]

    # status para agrupamento (expand/collapse)
    def status_from_var(v):
        if pd.isna(v):
            return "Neutro"
        if v > 0:
            return "Desvio (↑)"
        if v < 0:
            return "Economia (↓)"
        return "Neutro"

    df_view["STATUS"] = df_view["VARIAÇÃO (R$)"].apply(status_from_var)

    # TOTAL (carteira)
    total_row = {"OBRA": "TOTAL (CARTEIRA)", "STATUS": "Carteira"}
    for m in months_in_range:
        total_row[m] = float(pd.to_numeric(df_view[m], errors="coerce").fillna(0).sum())
    total_row["VARIAÇÃO (R$)"] = float(total_row[last_m] - total_row[first_m])

    # Δ M/M (carteira)
    delta_row = {"OBRA": "Δ MÊS/MÊS (CARTEIRA)", "STATUS": "Carteira"}
    prev = None
    for m in months_in_range:
        cur = total_row.get(m)
        if prev is None:
            delta_row[m] = None
        else:
            delta_row[m] = (cur - prev) if (cur is not None and prev is not None) else None
        prev = cur
    delta_row["VARIAÇÃO (R$)"] = None

    df_view = pd.concat([df_view, pd.DataFrame([total_row, delta_row])], ignore_index=True)

    # ----------------------------
    # Obra em foco (para cards bonitos)
    # ----------------------------
    if "obra_foco" not in st.session_state:
        # pega primeira obra real
        first_real = None
        for o in df_f["OBRA"].tolist():
            if o in obras:
                first_real = o
                break
        st.session_state["obra_foco"] = first_real or (obras[0] if obras else "")

    foco_candidates = [o for o in df_f["OBRA"].dropna().unique().tolist() if o in obras]
    if not foco_candidates:
        foco_candidates = obras

    # placeholder para obra foco (atualiza quando clica na tabela)
    obra_foco = st.selectbox(
        "Obra em foco (cards abaixo)",
        foco_candidates,
        index=foco_candidates.index(st.session_state["obra_foco"]) if st.session_state["obra_foco"] in foco_candidates else 0,
        key="obra_foco",
    )

    # cards lado a lado (como o print) — para obra foco
    if obra_foco in obras:
        ws_det = wb[obra_foco]
        df_acres_det, df_econ_det = read_acrescimos_economias(ws_det)

        econ_items = []
        if df_econ_det is not None and not df_econ_det.empty and "VARIAÇÃO" in df_econ_det.columns:
            econ_sorted = df_econ_det.copy()
            econ_sorted["__v"] = pd.to_numeric(econ_sorted["VARIAÇÃO"], errors="coerce")
            econ_sorted = econ_sorted.dropna(subset=["__v"])
            econ_sorted["__abs"] = econ_sorted["__v"].abs()
            econ_sorted = econ_sorted.sort_values("__abs", ascending=False).head(3)
            for _, r in econ_sorted.iterrows():
                econ_items.append((str(r.get("DESCRIÇÃO", "")).strip(), float(r.get("__v", 0) or 0)))

        acres_items = []
        if df_acres_det is not None and not df_acres_det.empty and "VARIAÇÃO" in df_acres_det.columns:
            acres_sorted = df_acres_det.copy()
            acres_sorted["__v"] = pd.to_numeric(acres_sorted["VARIAÇÃO"], errors="coerce")
            acres_sorted = acres_sorted.dropna(subset=["__v"])
            acres_sorted["__abs"] = acres_sorted["__v"].abs()
            acres_sorted = acres_sorted.sort_values("__abs", ascending=False).head(3)
            for _, r in acres_sorted.iterrows():
                acres_items.append((str(r.get("DESCRIÇÃO", "")).strip(), float(r.get("__v", 0) or 0)))

        econ_rows = build_rows(econ_items, color=PALETTE["good"], prefix="")
        acres_rows = build_rows(acres_items, color=PALETTE["bad"], prefix="- ")

        cA, cB = st.columns(2)
        with cA:
            st.markdown(card_resumo("PRINCIPAIS ECONOMIAS", "✅", econ_rows, PALETTE["good_border"], PALETTE["good_bg"]), unsafe_allow_html=True)
        with cB:
            st.markdown(card_resumo("DESVIOS DO MÊS", "⚠️", acres_rows, PALETTE["bad_border"], PALETTE["bad_bg"]), unsafe_allow_html=True)

    st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

    # ----------------------------
    # AgGrid (BI)
    # ----------------------------
    money_formatter = JsCode("""
    function(params){
      if (params.value === null || params.value === undefined || params.value === "") return "";
      const v = Number(params.value);
      if (isNaN(v)) return params.value;
      return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(v);
    }
    """)

    cell_style = JsCode("""
    function(params){
      const field = params.colDef.field;
      const row = (params.data && params.data.OBRA) ? String(params.data.OBRA) : "";
      const v = params.value;

      if (field === "OBRA"){
        if (row.startsWith("TOTAL")) return {fontWeight: 900};
        if (row.startsWith("Δ")) return {fontWeight: 900};
        return {fontWeight: 900};
      }

      if (v === null || v === undefined || v === "") return {};

      // Linha Δ: cores vibrantes
      if (row.startsWith("Δ")){
        const n = Number(v);
        if (isNaN(n)) return {};
        if (n > 0) return { backgroundColor: "#ef4444", color: "white", fontWeight: 900 };
        if (n < 0) return { backgroundColor: "#22c55e", color: "white", fontWeight: 900 };
        return { backgroundColor: "#94a3b8", color: "white", fontWeight: 900 };
      }

      // Coluna variação: vibrante
      if (field && field.toUpperCase().includes("VARIA")){
        const n = Number(v);
        if (isNaN(n)) return {};
        if (n > 0) return { backgroundColor: "#ef4444", color: "white", fontWeight: 900 };
        if (n < 0) return { backgroundColor: "#22c55e", color: "white", fontWeight: 900 };
        return { backgroundColor: "#94a3b8", color: "white", fontWeight: 900 };
      }

      // TOTAL: destaque leve
      if (row.startsWith("TOTAL")){
        return { fontWeight: 900, backgroundColor: "rgba(255,255,255,0.04)" };
      }

      return {};
    }
    """)

    # builder
    gb = GridOptionsBuilder.from_dataframe(df_view)
    gb.configure_default_column(
        resizable=True,
        sortable=True,
        filter=True,
        wrapText=True,
        autoHeight=True,
    )

    gb.configure_grid_options(
        enableRangeSelection=True,
        rowSelection="single",
        sideBar=True,  # painel de filtros
        groupDefaultExpanded=0,  # grupos recolhidos por padrão
    )

    # Agrupamento por STATUS (clica e abre o grupo)
    gb.configure_column("STATUS", rowGroup=True, hide=True)

    # OBRA fixa
    gb.configure_column("OBRA", pinned="left", width=210)

    # meses
    for m in months_in_range:
        gb.configure_column(m, type=["numericColumn"], valueFormatter=money_formatter, cellStyle=cell_style, width=140)

    # variação fixa à direita
    gb.configure_column("VARIAÇÃO (R$)", pinned="right", type=["numericColumn"], valueFormatter=money_formatter, cellStyle=cell_style, width=160)

    grid_options = gb.build()

    grid = AgGrid(
        df_view,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        allow_unsafe_jscode=True,
        fit_columns_on_grid_load=False,
        height=280 if not expandir else 360,
    )

    # seleção (blindado)
    selected = grid.get("selected_rows", None)
    if selected is None:
        selected_rows = []
    elif isinstance(selected, (list, tuple)):
        selected_rows = selected
    elif hasattr(selected, "to_dict"):
        try:
            selected_rows = selected.to_dict("records")
        except Exception:
            selected_rows = []
    else:
        selected_rows = []

    if len(selected_rows) > 0:
        obra_sel = selected_rows[0].get("OBRA")
        if obra_sel in obras and obra_sel != st.session_state.get("obra_foco"):
            st.session_state["obra_foco"] = obra_sel
            st.rerun()
