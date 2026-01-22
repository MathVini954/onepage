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
# Tema (blindado: NÃO mexe em cor global)
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

# ✅ CSS blindado (não some com a tela)
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


# ✅ NOVO: % com cor (para aderência)
def kpi_card_pct_highlight(label: str, value_ratio: float | None, sub: str = "", value_color: str | None = None):
    col = value_color if value_color else PALETTE["text"]
    st.markdown(
        f"""
<div style="border:1px solid {PALETTE["border"]}; border-radius:14px; padding:12px 14px; background:{PALETTE["card"]}; height:92px;">
  <div style="font-size:12px; color:{PALETTE["muted"]}; margin-bottom:6px;">{html.escape(label)}</div>
  <div style="font-size:24px; font-weight:900; line-height:1.05; color:{col};">{html.escape(pct(value_ratio))}</div>
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
# Header + logo (título já com nome da obra)
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


# ============================================================
# Ler dados
# ============================================================
resumo = read_resumo_financeiro(ws)
df_idx = read_indice(ws)
df_fin = read_financeiro(ws)
df_prazo = read_prazo(ws)
df_acres, df_econ = read_acrescimos_economias(ws)

# Totais (cards que você pediu voltar)
total_economias = sum_abs_column(df_econ, "VARIAÇÃO")
total_acrescimos = sum_abs_column(df_acres, "VARIAÇÃO")
desvio_liquido = total_acrescimos - total_economias  # >0 pior, <0 melhor


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
k_ader_mes = None  # ✅ NOVO

planned_m = []
previsto_m = []
real_m = []

planned_acum = []
previsto_acum = []
real_acum = []

if not df_prazo.empty and "MÊS" in df_prazo.columns:
    temp = df_prazo.copy().dropna(subset=["MÊS"]).sort_values("MÊS").reset_index(drop=True)

    # Series mensais (0-1)
    temp["PLANEJADO_M"] = (
        temp["PLANEJADO MÊS (%)"].apply(to_ratio) if "PLANEJADO MÊS (%)" in temp.columns else pd.NA
    )
    temp["PREVISTO_M"] = (
        temp["PREVISTO MENSAL (%)"].apply(to_ratio) if "PREVISTO MENSAL (%)" in temp.columns else pd.NA
    )
    temp["REAL_M"] = (
        temp["REALIZADO Mês (%)"].apply(to_ratio) if "REALIZADO Mês (%)" in temp.columns else pd.NA
    )

    # Planejado acumulado: preferir do Excel (se existir). Se não, cumsum mantendo NaN (não preenche com 0)
    if "PLANEJADO ACUM. (%)" in temp.columns:
        temp["PLANEJADO_ACUM"] = temp["PLANEJADO ACUM. (%)"].apply(to_ratio)
    else:
        temp["PLANEJADO_ACUM"] = pd.to_numeric(temp["PLANEJADO_M"], errors="coerce").cumsum()

    # Previsto e Real acumulados: cumsum mantendo NaN (isso evita “linha reta até o fim”)
    temp["PREVISTO_ACUM"] = pd.to_numeric(temp["PREVISTO_M"], errors="coerce").cumsum()
    temp["REAL_ACUM"] = pd.to_numeric(temp["REAL_M"], errors="coerce").cumsum()

    # 🔥 CORTAR O DATAFRAME NO ÚLTIMO MÊS QUE TEM QUALQUER DADO (plan/prev/real)
    last_idxs = []
    for col in ["PLANEJADO_M", "PREVISTO_M", "REAL_M", "PLANEJADO_ACUM", "PREVISTO_ACUM", "REAL_ACUM"]:
        if col in temp.columns:
            idx = temp[col].last_valid_index()
            if idx is not None:
                last_idxs.append(idx)

    if last_idxs:
        last_any = max(last_idxs)
        temp = temp.iloc[: last_any + 1].copy()

    # helper: após último valor válido, vira None (Plotly para a linha)
    def series_stop_at_last(s: pd.Series) -> list[float | None]:
        s = pd.to_numeric(s, errors="coerce")
        last = s.last_valid_index()
        if last is None:
            return [None] * len(s)
        out = s.copy()
        for i in range(last + 1, len(out)):
            out.iloc[i] = pd.NA
        return [None if pd.isna(v) else float(v) for v in out.tolist()]

    # % para gráfico
    planned_m = [None if v is None else v * 100 for v in series_stop_at_last(temp["PLANEJADO_M"])]
    previsto_m = [None if v is None else v * 100 for v in series_stop_at_last(temp["PREVISTO_M"])]
    real_m = [None if v is None else v * 100 for v in series_stop_at_last(temp["REAL_M"])]

    planned_acum = [None if v is None else v * 100 for v in series_stop_at_last(temp["PLANEJADO_ACUM"])]
    previsto_acum = [None if v is None else v * 100 for v in series_stop_at_last(temp["PREVISTO_ACUM"])]
    real_acum = [None if v is None else v * 100 for v in series_stop_at_last(temp["REAL_ACUM"])]

    # ref mês: último real
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

        # ✅ Aderência (evita pd.NA em boolean)
        plan_a = pd.to_numeric(pd.Series([k_plan_acum]), errors="coerce").iloc[0]
        real_a = pd.to_numeric(pd.Series([k_real_acum]), errors="coerce").iloc[0]
        if pd.notna(plan_a) and float(plan_a) != 0:
            k_ader_acc = (float(real_a) / float(plan_a)) * 100

        plan_m = pd.to_numeric(pd.Series([k_plan_m]), errors="coerce").iloc[0]
        real_mx = pd.to_numeric(pd.Series([k_real_m]), errors="coerce").iloc[0]
        if pd.notna(plan_m) and float(plan_m) != 0:
            k_ader_mes = (float(real_mx) / float(plan_m)) * 100


# ============================================================
# Tabs
# ============================================================
tab_dash, tab_just = st.tabs(["Dashboard", "Justificativas"])


# ============================================================
# TAB Dashboard
# ============================================================
with tab_dash:
    # ===== KPIs financeiros (8)
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

    st.divider()

    left, right = st.columns([2.2, 1])

    # ====== Left: gráficos + prazo
    with left:


        g1, g2 = st.columns(2)

        with g1:
            st.subheader("Índice Projetado (baseline 1,000)")
            if df_idx.empty:
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
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

        with g2:
            st.subheader("Desembolso x Medido (mês a mês)")
            if df_fin.empty:
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
                ader_color = PALETTE["good"] if (k_ader_acc is not None and k_ader_acc >= 100) else PALETTE["bad"]
                kpi_card_pct_highlight("Aderência acumulada", ader_ratio, "(Real acum ÷ Plan acum)", value_color=ader_color)

            st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

            r2 = st.columns(4)
            with r2[0]:
                kpi_card_pct("Realizado mensal", k_real_m, f"ref: {ref_month_label}")
            with r2[1]:
                kpi_card_pct("Previsto mensal", k_prev_m, f"ref: {ref_month_label}")
            with r2[2]:
                kpi_card_pct("Planejado mensal", k_plan_m, f"ref: {ref_month_label}")
            with r2[3]:
                ader_mes_ratio = (k_ader_mes / 100) if k_ader_mes is not None else None
                ader_mes_color = PALETTE["good"] if (k_ader_mes is not None and k_ader_mes >= 100) else PALETTE["bad"]
                kpi_card_pct_highlight("Aderência do mês", ader_mes_ratio, "(Real mês ÷ Plan mês)", value_color=ader_mes_color)

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

    # ====== Right: cards economias/desvios + progresso
    with right:
        econ_items: list[tuple[str, float]] = []
        acres_items: list[tuple[str, float]] = []

        if not df_econ.empty and "VARIAÇÃO" in df_econ.columns:
            econ_sorted = df_econ.copy()
            econ_sorted["__v"] = pd.to_numeric(econ_sorted["VARIAÇÃO"], errors="coerce")
            econ_sorted = econ_sorted.dropna(subset=["__v"])
            econ_sorted["__abs"] = econ_sorted["__v"].abs()
            econ_sorted = econ_sorted.sort_values("__abs", ascending=False)
            for _, r in econ_sorted.head(3).iterrows():
                econ_items.append((str(r.get("DESCRIÇÃO", "")), float(r.get("__v", 0) or 0)))

        if not df_acres.empty and "VARIAÇÃO" in df_acres.columns:
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

   

    # ====== Detalhamento completo com degradê
    st.subheader("Detalhamento — Tabelas completas (com barras em degradê)")

    # ✅ 3 cards (ANTES dos 2 gráficos do detalhamento)
    row3 = st.columns(3)
    with row3[0]:
        kpi_card_money_highlight("Total Economias (ACM)", total_economias, PALETTE["good"])
    with row3[1]:
        kpi_card_money_highlight("Total Acréscimos (ACM)", total_acrescimos, PALETTE["bad"])
    with row3[2]:
        color_desvio = PALETTE["bad"] if desvio_liquido > 0 else PALETTE["good"]
        kpi_card_money_highlight("Desvio Orçamento (Acrésc. − Econ.)", desvio_liquido, color_desvio)

    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

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
