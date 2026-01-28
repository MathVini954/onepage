from __future__ import annotations

import html
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from src.excel_reader import (
    load_wb,
    sheetnames,
    read_resumo_financeiro,
    read_indice,
    read_financeiro,
    read_prazo,
    read_acrescimos_economias,
    read_orcamento_resumo,   # <- NOVO
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
# ✅ Aba extra (não interfere no Dashboard/Justificativas)
df_orc_resumo = read_orcamento_resumo(wb)

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
def clean_month_col(df: pd.DataFrame, col: str = "MÊS") -> pd.DataFrame:
    """
    FIX do eixo: remove microssegundos/horas e força mês puro (1º dia, 00:00:00).
    Evita aparecer 23:59:59.9995 / 00:00:00.0005 e qualquer “epoch weird”.
    """
    if df is None or df.empty or col not in df.columns:
        return df
    out = df.copy()
    out[col] = pd.to_datetime(out[col], errors="coerce")
    out = out.dropna(subset=[col])
    out[col] = out[col].dt.to_period("M").dt.to_timestamp()
    return out


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
    money_cols = ["ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO"]
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


# ============================================================
# Ler dados
# ============================================================
resumo = read_resumo_financeiro(ws)
df_idx = read_indice(ws)
df_fin = read_financeiro(ws)
df_prazo = read_prazo(ws)
df_acres, df_econ = read_acrescimos_economias(ws)

# ✅ FIX do eixo em todos os blocos com mês (remove microsegundos/horas)
df_idx = clean_month_col(df_idx, "MÊS")
df_fin = clean_month_col(df_fin, "MÊS")
df_prazo = clean_month_col(df_prazo, "MÊS")

# Totais
total_economias = sum_abs_column(df_econ, "VARIAÇÃO")
total_acrescimos = sum_abs_column(df_acres, "VARIAÇÃO")
desvio_liquido = total_acrescimos - total_economias  # >0 pior, <0 melhor


# ============================================================
# Índice do mês (último)
# ============================================================
idx_last = None
idx_month_label = "—"
if df_idx is not None and not df_idx.empty and "ÍNDICE PROJETADO" in df_idx.columns:
    df_idx2 = df_idx.dropna(subset=["MÊS"]).sort_values("MÊS")
    df_idx2["ÍNDICE PROJETADO"] = pd.to_numeric(df_idx2["ÍNDICE PROJETADO"], errors="coerce")
    df_idx2 = df_idx2.dropna(subset=["ÍNDICE PROJETADO"])
    if not df_idx2.empty:
        idx_last = float(df_idx2["ÍNDICE PROJETADO"].iloc[-1])
        m = df_idx2["MÊS"].iloc[-1]
        idx_month_label = pd.to_datetime(m).strftime("%b/%Y").lower()


# ============================================================
# Prazo — preparar séries e CORTAR no último mês preenchido
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

if df_prazo is not None and not df_prazo.empty and "MÊS" in df_prazo.columns:
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

    # corta no último mês com qualquer valor válido
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
            if df_idx is None or df_idx.empty:
                st.info("Sem dados do índice.")
            else:
                fig = go.Figure()
                fig.add_trace(
                    go.Scatter(
                        x=df_idx["MÊS"],
                        y=df_idx["ÍNDICE PROJETADO"],
                        mode="lines+markers",
                        name="Índice",
                    )
                )
                fig.add_hline(y=1.0, line_dash="dash", line_width=1)
                fig.update_layout(height=320)
                fig.update_xaxes(dtick="M1", tickformat="%b/%Y")  # ✅ sem hora/micro
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

        with g2:
            st.subheader("Desembolso x Medido (mês a mês)")
            if df_fin is None or df_fin.empty:
                st.info("Sem dados financeiros.")
            else:
                fig = go.Figure()
                fig.add_trace(
                    go.Bar(
                        x=df_fin["MÊS"],
                        y=df_fin["DESEMBOLSO DO MÊS (R$)"],
                        name="Desembolso",
                        marker_color=PALETTE["bar_des"],
                    )
                )
                fig.add_trace(
                    go.Bar(
                        x=df_fin["MÊS"],
                        y=df_fin["MEDIDO NO MÊS (R$)"],
                        name="Medido",
                        marker_color=PALETTE["bar_med"],
                    )
                )
                fig.update_layout(barmode="group", height=320)
                fig.update_xaxes(dtick="M1", tickformat="%b/%Y")  # ✅
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
                fig.update_xaxes(dtick="M1", tickformat="%b/%Y")  # ✅
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

            with t2:
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=x, y=planned_m, mode="lines+markers", name="Planejado mês (%)"))
                fig.add_trace(go.Scatter(x=x, y=previsto_m, mode="lines+markers", name="Previsto mês (%)"))
                fig.add_trace(go.Scatter(x=x, y=real_m, mode="lines+markers", name="Realizado mês (%)"))
                fig.update_layout(height=320, yaxis_title="% (mensal)")
                fig.update_xaxes(dtick="M1", tickformat="%b/%Y")  # ✅
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

    with right:
        econ_items: list[tuple[str, float]] = []
        acres_items: list[tuple[str, float]] = []

        if df_econ is not None and not df_econ.empty and "VARIAÇÃO" in df_econ.columns:
            econ_sorted = df_econ.copy()
            econ_sorted["__v"] = pd.to_numeric(econ_sorted["VARIAÇÃO"], errors="coerce")
            econ_sorted = econ_sorted.dropna(subset=["__v"])
            econ_sorted["__abs"] = econ_sorted["__v"].abs()
            econ_sorted = econ_sorted.sort_values("__abs", ascending=False)
            for _, r in econ_sorted.head(3).iterrows():
                econ_items.append((str(r.get("DESCRIÇÃO", "")), float(r.get("__v", 0) or 0)))

        if df_acres is not None and not df_acres.empty and "VARIAÇÃO" in df_acres.columns:
            acres_sorted = df_acres.copy()
            acres_sorted["__v"] = pd.to_numeric(acres_sorted["VARIAÇÃO"], errors="coerce")
            acres_sorted = acres_sorted.dropna(subset=["__v"])
            acres_sorted["__abs"] = acres_sorted["__v"].abs()
            acres_sorted = acres_sorted.sort_values("__abs", ascending=False)
            for _, r in acres_sorted.head(3).iterrows():
                acres_items.append((str(r.get("DESCRIÇÃO", "")), float(r.get("__v", 0) or 0)))

        econ_rows = build_rows(econ_items, color=PALETTE["good"], prefix="")
        acres_rows = build_rows(acres_items, color=PALETTE["bad"], prefix="- ")

        st.markdown(
            card_resumo("PRINCIPAIS ECONOMIAS", "✅", econ_rows, PALETTE["good_border"], PALETTE["good_bg"]),
            unsafe_allow_html=True,
        )
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        st.markdown(
            card_resumo("DESVIOS DO MÊS", "⚠️", acres_rows, PALETTE["bad_border"], PALETTE["bad_bg"]),
            unsafe_allow_html=True,
        )

        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        progress_card(k_real_acum, k_plan_acum, ref_month_label)

    st.divider()

    st.subheader("Detalhamento — Tabelas completas (com barras em degradê)")

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### ACRÉSCIMOS / DESVIOS")
        if df_acres is None or df_acres.empty:
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
            fig.add_trace(
                go.Bar(
                    x=vals,
                    y=top_bar["DESCRIÇÃO"],
                    orientation="h",
                    marker=dict(color=vals, colorscale=PALETTE["bad_grad"], showscale=False),
                    name="R$",
                )
            )
            fig.update_layout(height=340, xaxis_title="R$")
            st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

            with st.expander("Ver tabela (Acréscimos)"):
                styled_dataframe(show_top.drop(columns=["__abs"], errors="ignore"))

    with c2:
        st.markdown("### ECONOMIAS")
        if df_econ is None or df_econ.empty:
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
            fig.add_trace(
                go.Bar(
                    x=vals,
                    y=top_bar["DESCRIÇÃO"],
                    orientation="h",
                    marker=dict(color=vals, colorscale=PALETTE["good_grad"], showscale=False),
                    name="R$",
                )
            )
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
            st.markdown(
                f"<div style='color:{PALETTE['muted']}; font-size:12px;'>Sem dados</div></div>",
                unsafe_allow_html=True,
            )
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
        with st.expander("Ver tabela completa (Economias)"):
            styled_dataframe(df_econ)

    with b:
        list_just(df_acres, "TOP 5 — DESVIOS / ACRÉSCIMOS (com justificativa)", PALETTE["bad"])
        with st.expander("Ver tabela completa (Desvios)"):
            styled_dataframe(df_acres)


if debug:
    st.write("Arquivo:", excel_path.name)
    st.write("Obras:", obras)
    st.write("df_idx.head():", df_idx.head() if df_idx is not None else None)


# === ABA: RESUMO (ORÇAMENTO_RESUMO) ===
with tab_resumo:
    st.subheader("Resumo das Obras — ORÇAMENTO_RESUMO")

    if df_orc_resumo is None or df_orc_resumo.empty:
        st.info("A aba **ORÇAMENTO_RESUMO** não foi encontrada ou está vazia.")
    else:
        df_show = df_orc_resumo.copy()

        # =========================
        # Helpers / Preparação
        # =========================
        import re
        import unicodedata
        import pandas as pd

        def _norm_colname(x: str) -> str:
            s = "" if x is None else str(x).strip()
            s = unicodedata.normalize("NFKD", s)
            s = "".join(ch for ch in s if not unicodedata.combining(ch))
            return " ".join(s.upper().split())

        # garante OBRA
        if "OBRA" in df_show.columns:
            df_show["OBRA"] = df_show["OBRA"].astype(str).str.strip()
        else:
            st.error("A coluna **OBRA** não foi encontrada na aba ORÇAMENTO_RESUMO.")
            st.stop()

        # detectar coluna variação (primeira que contenha 'VARIA')
        variacao_col = None
        for c in df_show.columns:
            if "VARIA" in _norm_colname(c):
                variacao_col = c
                break

        # detectar colunas de mês (tudo que não é OBRA e não é VARIAÇÃO)
        month_cols = []
        for c in df_show.columns:
            nc = _norm_colname(c)
            if nc == "OBRA":
                continue
            if variacao_col is not None and c == variacao_col:
                continue
            month_cols.append(c)

        # converte meses + variação pra número
        for c in month_cols + ([variacao_col] if variacao_col else []):
            if c is None:
                continue
            df_show[c] = pd.to_numeric(df_show[c], errors="coerce")

        # =========================
        # Ordenação dos meses
        # =========================
        def _month_sort_key(col):
            # tenta entender: 01/2026, 1/2026, JAN/2026, JAN 2026 etc.
            s = _norm_colname(col)

            # 01/2026
            m = re.search(r"\b(\d{1,2})\s*/\s*(\d{4})\b", s)
            if m:
                mm = int(m.group(1))
                yy = int(m.group(2))
                if 1 <= mm <= 12:
                    return pd.Timestamp(yy, mm, 1)

            # JAN/2026
            pt = {
                "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
                "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12
            }
            m2 = re.search(r"\b(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\b", s)
            y2 = re.search(r"\b(20\d{2})\b", s)
            if m2 and y2:
                mm = pt[m2.group(1)]
                yy = int(y2.group(1))
                return pd.Timestamp(yy, mm, 1)

            # fallback: joga pro fim mantendo algo estável
            return pd.Timestamp(2999, 12, 1)

        # mantém só meses com algum valor e ordena
        month_cols_sorted = [c for c in month_cols if df_show[c].notna().any()]
        month_cols_sorted = sorted(month_cols_sorted, key=_month_sort_key) if month_cols_sorted else []

        # =========================
        # FILTRO DE PERÍODO (TOPO)
        # =========================
        st.markdown("#### Período de visualização")

        if not month_cols_sorted:
            st.warning("Não encontrei colunas de mês com valores para montar o período.")
            sel_month_cols = []
        else:
            if len(month_cols_sorted) == 1:
                sel_month_cols = month_cols_sorted[:]
                st.caption(f"Somente 1 mês disponível: **{sel_month_cols[0]}**")
            else:
                # default: últimos 6 meses (se tiver)
                start_idx = max(0, len(month_cols_sorted) - 6)
                default_range = (month_cols_sorted[start_idx], month_cols_sorted[-1])

                periodo = st.select_slider(
                    "Selecione o período (mês inicial → mês final)",
                    options=month_cols_sorted,
                    value=default_range,
                    key="periodo_resumo_orc",
                )

                i0 = month_cols_sorted.index(periodo[0])
                i1 = month_cols_sorted.index(periodo[1])
                if i0 > i1:
                    i0, i1 = i1, i0
                sel_month_cols = month_cols_sorted[i0:i1 + 1]

        last_month_col = sel_month_cols[-1] if sel_month_cols else None

        # =========================
        # CONTROLES VISUAIS
        # =========================
        c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.4, 2.2])
        with c1:
            expandir = st.toggle("Mostrar colunas mês a mês", value=False, key="exp_meses_orc")
        with c2:
            ordenar_impacto = st.toggle("Ordenar por maior impacto", value=True, key="ord_impacto_orc")
        with c3:
            somente_com_mov = st.toggle("Somente com movimento", value=False, key="mov_orc")
        with c4:
            busca_obra = st.text_input("Buscar obra", value="", placeholder="Digite parte do nome...", key="busca_orc")

        if last_month_col:
            st.caption(f"Período até **{last_month_col}** | Variação final: **{variacao_col or 'N/A'}**")

        # =========================
        # MONTA TABELA MAIS VISUAL
        # =========================
        df_view = df_show.copy()

        # Filtro: busca obra
        if busca_obra.strip():
            mask = df_view["OBRA"].astype(str).str.upper().str.contains(busca_obra.strip().upper(), na=False)
            df_view = df_view[mask].copy()

        # Filtro: somente com movimento no período (qualquer mês != 0 / notna)
        if somente_com_mov and sel_month_cols:
            m = df_view[sel_month_cols].copy()
            has_mov = m.notna().any(axis=1) & (m.fillna(0).abs().sum(axis=1) > 0)
            df_view = df_view[has_mov].copy()

        # Sparkline: linha mês a mês do período selecionado
        if sel_month_cols:
            df_view["TENDÊNCIA (período)"] = df_view[sel_month_cols].values.tolist()
        else:
            df_view["TENDÊNCIA (período)"] = [[] for _ in range(len(df_view))]

        # Sinal visual da variação final (mantém a VARIAÇÃO como coluna numérica)
        if variacao_col and variacao_col in df_view.columns:
            def _sinal(v):
                try:
                    v = float(v)
                except Exception:
                    return "—"
                if v > 0:
                    return "🟥"  # desvio (pior)
                if v < 0:
                    return "🟩"  # economia (melhor)
                return "⚪"
            df_view["SINAL"] = df_view[variacao_col].apply(_sinal)

        # Seleção de colunas (visual)
        cols_out = ["OBRA", "TENDÊNCIA (período)"]

        if expandir and sel_month_cols:
            cols_out += sel_month_cols
        else:
            if last_month_col:
                cols_out += [last_month_col]

        if variacao_col:
            if "SINAL" in df_view.columns:
                cols_out += ["SINAL", variacao_col]
            else:
                cols_out += [variacao_col]

        df_view = df_view[cols_out].copy()

        # Ordenação por maior impacto (abs da variação final)
        if ordenar_impacto and variacao_col and variacao_col in df_view.columns:
            df_view["__abs_var"] = pd.to_numeric(df_view[variacao_col], errors="coerce").abs()
            df_view = df_view.sort_values("__abs_var", ascending=False).drop(columns=["__abs_var"])

        # Configuração de colunas (sparkline + moeda)
        col_cfg = {
            "TENDÊNCIA (período)": st.column_config.LineChartColumn(
                "Tendência (período)",
                help="Linha mês a mês do período selecionado (por obra).",
                width="medium",
            )
        }

        # Formato moeda (se preferir, pode trocar por fmt_brl via string no seu fmt)
        money_fmt = "R$ %.2f"

        for c in df_view.columns:
            if _norm_colname(c) == "OBRA" or c in ("TENDÊNCIA (período)", "SINAL"):
                continue
            col_cfg[c] = st.column_config.NumberColumn(c, format=money_fmt)

        if "SINAL" in df_view.columns:
            col_cfg["SINAL"] = st.column_config.TextColumn(" ", width="small")

        st.dataframe(
            df_view,
            use_container_width=True,
            hide_index=True,
            column_config=col_cfg,
        )

        st.markdown("---")

        # =========================
        # DETALHES (mantém como você já tinha)
        # =========================
        st.subheader("Detalhes (Economias e Desvios do mês)")

        obras_opts = df_show["OBRA"].dropna().astype(str).str.strip().tolist()
        obras_opts = [o for o in obras_opts if o]  # remove vazios
        obra_sel = st.selectbox(
            "Escolha a obra para ver os detalhes",
            options=obras_opts,
            index=0 if len(obras_opts) > 0 else None,
            key="obra_sel_orc",
        )

        if obra_sel:
            # abre a planilha da obra
            try:
                ws_det = wb[obra_sel]  # wb deve existir no seu app (openpyxl workbook)
            except Exception:
                st.error(f"Não encontrei a aba da obra **{obra_sel}** dentro do arquivo.")
            else:
                df_acres_det, df_econ_det = read_acrescimos_economias(ws_det)

                top_cards = 3  # mude para 5 se quiser

                econ_items = []
                if df_econ_det is not None and not df_econ_det.empty and "VARIAÇÃO" in df_econ_det.columns:
                    econ_sorted = df_econ_det.copy()
                    econ_sorted["__v"] = pd.to_numeric(econ_sorted["VARIAÇÃO"], errors="coerce")
                    econ_sorted = econ_sorted.dropna(subset=["__v"])
                    econ_sorted["__abs"] = econ_sorted["__v"].abs()
                    econ_sorted = econ_sorted.sort_values("__abs", ascending=False).head(top_cards)
                    for _, r in econ_sorted.iterrows():
                        econ_items.append((str(r.get("DESCRIÇÃO", "")).strip(), float(r.get("__v", 0) or 0)))

                acres_items = []
                if df_acres_det is not None and not df_acres_det.empty and "VARIAÇÃO" in df_acres_det.columns:
                    acres_sorted = df_acres_det.copy()
                    acres_sorted["__v"] = pd.to_numeric(acres_sorted["VARIAÇÃO"], errors="coerce")
                    acres_sorted = acres_sorted.dropna(subset=["__v"])
                    acres_sorted["__abs"] = acres_sorted["__v"].abs()
                    acres_sorted = acres_sorted.sort_values("__abs", ascending=False).head(top_cards)
                    for _, r in acres_sorted.iterrows():
                        acres_items.append((str(r.get("DESCRIÇÃO", "")).strip(), float(r.get("__v", 0) or 0)))

                econ_rows = build_rows(econ_items, color=PALETTE["good"], prefix="")
                acres_rows = build_rows(acres_items, color=PALETTE["bad"], prefix="- ")

                st.markdown(
                    card_resumo("PRINCIPAIS ECONOMIAS", "✅", econ_rows, PALETTE["good_border"], PALETTE["good_bg"]),
                    unsafe_allow_html=True,
                )
                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
                st.markdown(
                    card_resumo("DESVIOS DO MÊS", "⚠️", acres_rows, PALETTE["bad_border"], PALETTE["bad_bg"]),
                    unsafe_allow_html=True,
                )
