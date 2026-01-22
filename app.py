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
# CONFIG
# ============================================================
st.set_page_config(page_title="Controle Prazo e Custo", layout="wide")

LOGOS_DIR = "assets/logos"
GOOD = "#22c55e"
BAD = "#ef4444"
BLUE = "#3b82f6"


# ============================================================
# EXCEL ÚNICO (SEM UPLOAD)
# ============================================================
def find_excel_in_root() -> Path | None:
    """Procura o 1º xlsx/xlsm na raiz (ignorando temporários)."""
    root = Path(".")
    candidates = []
    for ext in ("*.xlsm", "*.xlsx"):
        for p in root.glob(ext):
            if p.name.startswith("~$"):
                continue
            if p.is_file():
                candidates.append(p)
    # preferir Excel.xlsm/xlsx se existir
    for prefer in ("Excel.xlsm", "Excel.xlsx", "excel.xlsm", "excel.xlsx"):
        pp = Path(prefer)
        if pp.exists() and pp.is_file():
            return pp
    return sorted(candidates)[0] if candidates else None


excel_path = find_excel_in_root()
if excel_path is None:
    st.error("Não achei nenhum Excel (.xlsx/.xlsm) na raiz do projeto.")
    st.stop()

wb = load_wb(excel_path)
obras = sheetnames(wb)
if not obras:
    st.error("Nenhuma aba de obra encontrada no Excel.")
    st.stop()


# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.title("Controle Prazo e Custo")
obra = st.sidebar.selectbox("Obra (aba)", obras, index=0)

top_opt = st.sidebar.selectbox("Mostrar Top", ["5", "10", "Todas"], index=0)
top_n = None if top_opt == "Todas" else int(top_opt)

st.sidebar.markdown("---")
dark_mode = st.sidebar.toggle("Modo escuro", value=True)
debug = st.sidebar.toggle("Debug", value=False)

ws = wb[obra]


# ============================================================
# TEMA (CSS SEGURO: NÃO ESCONDE NADA)
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
        "card": "rgba(255,255,255,0.05)",
        "border": "rgba(255,255,255,0.12)",
        "track": "rgba(255,255,255,0.12)",
        "plotly_template": "plotly_dark",
        "bar_des": rgba(BLUE, 0.90),
        "bar_med": rgba(GOOD, 0.90),
        "good": GOOD,
        "bad": BAD,
        "good_grad": [[0, rgba(GOOD, 0.15)], [1, rgba(GOOD, 0.95)]],
        "bad_grad": [[0, rgba(BAD, 0.15)], [1, rgba(BAD, 0.95)]],
        "planned_bar": rgba(BLUE, 0.35),
        "real_bar": rgba(BLUE, 0.95),
    }
else:
    PALETTE = {
        "bg": "#f7f8fc",
        "sidebar_bg": "#ffffff",
        "text": "#0f172a",
        "muted": "#475569",
        "card": "rgba(255,255,255,0.95)",
        "border": "rgba(15,23,42,0.10)",
        "track": "rgba(15,23,42,0.10)",
        "plotly_template": "plotly_white",
        "bar_des": rgba(BLUE, 0.85),
        "bar_med": rgba(GOOD, 0.85),
        "good": GOOD,
        "bad": BAD,
        "good_grad": [[0, rgba(GOOD, 0.12)], [1, rgba(GOOD, 0.85)]],
        "bad_grad": [[0, rgba(BAD, 0.12)], [1, rgba(BAD, 0.85)]],
        "planned_bar": rgba(BLUE, 0.22),
        "real_bar": rgba(BLUE, 0.85),
    }

PLOTLY_TEMPLATE = PALETTE["plotly_template"]

st.markdown(
    f"""
<style>
  html, body, [data-testid="stAppViewContainer"], .stApp {{
    background: {PALETTE["bg"]} !important;
    color: {PALETTE["text"]} !important;
  }}
  section[data-testid="stSidebar"] > div {{
    background: {PALETTE["sidebar_bg"]} !important;
    border-right: 1px solid {PALETTE["border"]} !important;
  }}
  header[data-testid="stHeader"] {{
    background: {PALETTE["bg"]} !important;
    border-bottom: 1px solid {PALETTE["border"]} !important;
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
# HELPERS
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


def pct_ratio(r: float | None) -> str:
    if r is None:
        return "—"
    return f"{r*100:.1f}%".replace(".", ",")


def clamp01(v: float | None) -> float:
    if v is None:
        return 0.0
    return max(0.0, min(1.0, float(v)))


def kpi_card_money(label: str, value: float | None):
    st.markdown(
        f"""
<div style="border:1px solid {PALETTE["border"]}; border-radius:14px; padding:12px 14px; background:{PALETTE["card"]}; height:92px;">
  <div style="font-size:12px; color:{PALETTE["muted"]}; margin-bottom:6px;">{html.escape(label)}</div>
  <div style="font-size:24px; font-weight:900; line-height:1.05;">{html.escape(brl_compact(value))}</div>
  <div style="font-size:11px; color:{PALETTE["muted"]}; margin-top:6px;">{html.escape(fmt_brl(value))}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def kpi_card_money_color(label: str, value: float | None, color: str, sub: str = ""):
    st.markdown(
        f"""
<div style="border:1px solid {PALETTE["border"]}; border-radius:14px; padding:12px 14px; background:{PALETTE["card"]}; height:92px;">
  <div style="font-size:12px; color:{PALETTE["muted"]}; margin-bottom:6px;">{html.escape(label)}</div>
  <div style="font-size:24px; font-weight:900; line-height:1.05; color:{color};">{html.escape(brl_compact(value))}</div>
  <div style="font-size:11px; color:{PALETTE["muted"]}; margin-top:6px;">{html.escape(sub) if sub else html.escape(fmt_brl(value))}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def kpi_card_pct(label: str, ratio: float | None, sub: str = "", color: str | None = None):
    col = color if color else PALETTE["text"]
    st.markdown(
        f"""
<div style="border:1px solid {PALETTE["border"]}; border-radius:14px; padding:12px 14px; background:{PALETTE["card"]}; height:92px;">
  <div style="font-size:12px; color:{PALETTE["muted"]}; margin-bottom:6px;">{html.escape(label)}</div>
  <div style="font-size:24px; font-weight:900; line-height:1.05; color:{col};">{html.escape(pct_ratio(ratio))}</div>
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
        color = PALETTE["bad"] if idx > 1.0 else (PALETTE["good"] if idx < 1.0 else PALETTE["text"])

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
    <div style="font-size:12px; font-weight:900;">Deveria estar vs. Estou (acum.)</div>
    <div style="font-size:12px; color:{PALETTE["muted"]};">{html.escape(ref_month_label)}</div>
  </div>

  <div style="margin-top:12px; display:flex; justify-content:space-between; align-items:flex-end;">
    <div>
      <div style="font-size:12px; color:{PALETTE["muted"]};">Estou</div>
      <div style="font-size:28px; font-weight:900; line-height:1;">{real_pct:.0f}%</div>
    </div>
    <div style="text-align:right;">
      <div style="font-size:12px; color:{PALETTE["muted"]};">Deveria</div>
      <div style="font-size:16px; font-weight:900;">{planned_pct:.0f}%</div>
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
        rows += f"""
<div style="display:flex; justify-content:space-between; align-items:center; padding:10px 0; border-top:1px solid {PALETTE["border"]};">
  <div style="font-size:13px; font-weight:600;">{html.escape(str(desc))}</div>
  <div style="font-size:13px; font-weight:800; color:{color};">{prefix}{html.escape(fmt_brl_no_dec(abs(val)))}</div>
</div>
"""
    return rows


def card_resumo(title: str, icon: str, rows_html: str, border: str, bg: str) -> str:
    return f"""
<div style="border:1px solid {border}; background:{bg}; border-radius:16px; padding:14px 16px;">
  <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
    <div style="font-size:12px; font-weight:900; letter-spacing:0.3px;">{html.escape(title)}</div>
    <div style="font-size:12px;">{icon}</div>
  </div>
  {rows_html if rows_html else f'<div style="color:{PALETTE["muted"]}; font-size:12px;">Sem dados</div>'}
</div>
"""


def gradient_list(base_hex: str, n: int, a0: float = 0.25, a1: float = 0.95) -> list[str]:
    if n <= 1:
        return [rgba(base_hex, a1)]
    out = []
    for i in range(n):
        a = a0 + (a1 - a0) * (i / (n - 1))
        out.append(rgba(base_hex, a))
    return out


def safe_series(df: pd.DataFrame, col: str) -> pd.Series:
    return pd.to_numeric(df.get(col), errors="coerce")


# ============================================================
# HEADER
# ============================================================
cL, cR = st.columns([1, 5])
with cL:
    logo = find_logo_path(obra, LOGOS_DIR)
    if logo:
        st.image(logo, use_container_width=True)
with cR:
    st.title(f"Controle de Prazo e Custo — {obra}")

if debug:
    st.caption(f"Excel: {excel_path.name} • Abas: {len(obras)}")


# ============================================================
# LER DADOS
# ============================================================
resumo = read_resumo_financeiro(ws)
df_idx = read_indice(ws)
df_fin = read_financeiro(ws)
df_prazo = read_prazo(ws)
df_acres, df_econ = read_acrescimos_economias(ws)

total_economias = float(pd.to_numeric(df_econ.get("VARIAÇÃO"), errors="coerce").abs().sum()) if not df_econ.empty else 0.0
total_acrescimos = float(pd.to_numeric(df_acres.get("VARIAÇÃO"), errors="coerce").abs().sum()) if not df_acres.empty else 0.0
desvio_liquido = total_acrescimos - total_economias


# ============================================================
# ÍNDICE (último)
# ============================================================
idx_last = None
idx_month_label = "—"
if not df_idx.empty and "ÍNDICE PROJETADO" in df_idx.columns:
    aux = df_idx.dropna(subset=["MÊS"]).copy()
    aux["ÍNDICE PROJETADO"] = pd.to_numeric(aux["ÍNDICE PROJETADO"], errors="coerce")
    aux = aux.dropna(subset=["ÍNDICE PROJETADO"]).sort_values("MÊS")
    if not aux.empty:
        idx_last = float(aux["ÍNDICE PROJETADO"].iloc[-1])
        idx_month_label = pd.to_datetime(aux["MÊS"].iloc[-1]).strftime("%b/%Y").lower()


# ============================================================
# PRAZO (aderência do mês + previsto)  ✅ FIX AttributeError
# ============================================================
ref_month_label = "—"
k_real_acum = k_plan_acum = k_prev_acum = None
k_real_m = k_plan_m = k_prev_m = None
k_ader_acc = k_ader_mes = None

temp = pd.DataFrame()
planned_m = previsto_m = real_m = []
planned_ac = previsto_ac = real_ac = []
ader_mes_pct = []

if not df_prazo.empty and "MÊS" in df_prazo.columns:
    temp = df_prazo.dropna(subset=["MÊS"]).copy()
    temp = temp.sort_values("MÊS").reset_index(drop=True)

    # normaliza para ratio 0-1 (aceita 0-100)
    def to_ratio(v):
        v = pd.to_numeric(v, errors="coerce")
        if pd.isna(v):
            return pd.NA
        return v if v <= 1.5 else (v / 100.0)

    # ✅ garante Series mesmo se a coluna não existir
    def col_series(df: pd.DataFrame, name: str) -> pd.Series:
        if name in df.columns:
            return df[name]
        return pd.Series([pd.NA] * len(df), index=df.index)

    # nomes possíveis do previsto
    prev_col = None
    for cand in ["PREVISTO MENSAL (%)", "PREVISTO MENSAL(%)", "COMPROMETIDO MÊS (%)", "COMPROMETIDO MES (%)"]:
        if cand in temp.columns:
            prev_col = cand
            break

    temp["PLAN_M"] = col_series(temp, "PLANEJADO MÊS (%)").apply(to_ratio)
    temp["REAL_M"] = col_series(temp, "REALIZADO Mês (%)").apply(to_ratio)
    temp["PREV_M"] = (col_series(temp, prev_col).apply(to_ratio) if prev_col else col_series(temp, "__NA__"))

    # acumulados
    if "PLANEJADO ACUM. (%)" in temp.columns:
        temp["PLAN_A"] = col_series(temp, "PLANEJADO ACUM. (%)").apply(to_ratio)
    else:
        temp["PLAN_A"] = pd.to_numeric(temp["PLAN_M"], errors="coerce").cumsum(skipna=True)

    temp["REAL_A"] = pd.to_numeric(temp["REAL_M"], errors="coerce").cumsum(skipna=True)
    temp["PREV_A"] = pd.to_numeric(temp["PREV_M"], errors="coerce").cumsum(skipna=True)

    # ✅ cortar no último mês com valor REAL/PLAN/PREV (evita repetir último mês pra frente)
    def last_meaningful_idx(s: pd.Series) -> int | None:
        s2 = pd.to_numeric(s, errors="coerce")
        ok = s2.notna() & (s2.abs() > 1e-9)
        if ok.any():
            return int(ok[ok].index.max())
        return None

    lasts = []
    for col in ["PLAN_M", "REAL_M", "PREV_M", "PLAN_A", "REAL_A", "PREV_A"]:
        i = last_meaningful_idx(temp[col]) if col in temp.columns else None
        if i is not None:
            lasts.append(i)
    if lasts:
        temp = temp.iloc[: max(lasts) + 1].copy()

    # aderência do mês (ratio)
    denom = pd.to_numeric(temp["PLAN_M"], errors="coerce")
    num = pd.to_numeric(temp["REAL_M"], errors="coerce")
    ader_mes = num / denom
    ader_mes[(denom.isna()) | (denom == 0)] = pd.NA
    temp["ADER_M"] = ader_mes

    # listas para plot
    def to_plot_list(s: pd.Series, mul: float = 1.0) -> list[float | None]:
        s2 = pd.to_numeric(s, errors="coerce")
        return [None if pd.isna(v) else float(v) * mul for v in s2.tolist()]

    planned_m = to_plot_list(temp["PLAN_M"], 100)
    real_m = to_plot_list(temp["REAL_M"], 100)
    previsto_m = to_plot_list(temp["PREV_M"], 100)

    planned_ac = to_plot_list(temp["PLAN_A"], 100)
    real_ac = to_plot_list(temp["REAL_A"], 100)
    previsto_ac = to_plot_list(temp["PREV_A"], 100)

    ader_mes_pct = to_plot_list(temp["ADER_M"], 100)

    # último mês REAL
    last_real = pd.to_numeric(temp["REAL_M"], errors="coerce").last_valid_index()
    if last_real is not None:
        ref_month_label = pd.to_datetime(temp.loc[last_real, "MÊS"]).strftime("%b/%Y").lower()

        k_real_m = float(temp.loc[last_real, "REAL_M"]) if pd.notna(temp.loc[last_real, "REAL_M"]) else None
        k_plan_m = float(temp.loc[last_real, "PLAN_M"]) if pd.notna(temp.loc[last_real, "PLAN_M"]) else None
        k_prev_m = float(temp.loc[last_real, "PREV_M"]) if pd.notna(temp.loc[last_real, "PREV_M"]) else None

        k_real_acum = float(temp.loc[last_real, "REAL_A"]) if pd.notna(temp.loc[last_real, "REAL_A"]) else None
        k_plan_acum = float(temp.loc[last_real, "PLAN_A"]) if pd.notna(temp.loc[last_real, "PLAN_A"]) else None
        k_prev_acum = float(temp.loc[last_real, "PREV_A"]) if pd.notna(temp.loc[last_real, "PREV_A"]) else None

        if k_plan_acum and k_plan_acum != 0:
            k_ader_acc = k_real_acum / k_plan_acum
        if k_plan_m and k_plan_m != 0:
            k_ader_mes = k_real_m / k_plan_m



# ============================================================
# TABS
# ============================================================
tab_dash, tab_just = st.tabs(["Dashboard", "Justificativas"])

with tab_dash:
    # KPIs financeiros
    r1 = st.columns(4)
    with r1[0]:
        kpi_card_index("Índice do mês", idx_last, idx_month_label)
    with r1[1]:
        kpi_card_money("Orç. Inicial", resumo.get("ORÇAMENTO INICIAL (R$)"))
    with r1[2]:
        kpi_card_money("Orç. Reajust.", resumo.get("ORÇAMENTO REAJUSTADO (R$)"))
    with r1[3]:
        kpi_card_money("Desembolso Acum.", resumo.get("DESEMBOLSO ACUMULADO (R$)"))

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    r2 = st.columns(4)
    with r2[0]:
        kpi_card_money("A Pagar", resumo.get("A PAGAR (R$)"))
    with r2[1]:
        kpi_card_money("Saldo a Incorrer", resumo.get("SALDO A INCORRER (R$)"))
    with r2[2]:
        kpi_card_money("Custo Final", resumo.get("CUSTO FINAL (R$)"))
    with r2[3]:
        kpi_card_money("Variação", resumo.get("VARIAÇÃO (R$)"))

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # Totais mês
    r3 = st.columns(3)
    with r3[0]:
        kpi_card_money_color("Total Economias (mês)", total_economias, PALETTE["good"])
    with r3[1]:
        kpi_card_money_color("Total Acréscimos (mês)", total_acrescimos, PALETTE["bad"])
    with r3[2]:
        c = PALETTE["bad"] if desvio_liquido > 0 else PALETTE["good"]
        kpi_card_money_color("Desvio Líquido (Acrésc. − Econ.)", desvio_liquido, c)

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
            st.subheader("Desembolso x Medido (mês a mês) — degradê")
            if df_fin.empty:
                st.info("Sem dados financeiros.")
            else:
                df_f = df_fin.dropna(subset=["MÊS"]).copy()
                df_f = df_f.sort_values("MÊS")

                des = pd.to_numeric(df_f["DESEMBOLSO DO MÊS (R$)"], errors="coerce").fillna(0)
                med = pd.to_numeric(df_f["MEDIDO NO MÊS (R$)"], errors="coerce").fillna(0)

                des_colors = gradient_list(BLUE, len(df_f), 0.25, 0.95)
                med_colors = gradient_list(GOOD, len(df_f), 0.20, 0.90)

                fig = go.Figure()
                fig.add_trace(go.Bar(x=df_f["MÊS"], y=des, name="Desembolso", marker_color=des_colors))
                fig.add_trace(go.Bar(x=df_f["MÊS"], y=med, name="Medido", marker_color=med_colors))
                fig.update_layout(barmode="group", height=320)
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

        st.subheader("Prazo — Curva S + Aderência do mês")

        if temp.empty:
            st.info("Sem dados de prazo.")
        else:
            # avisar previsto vazio
            prev_ok = pd.to_numeric(temp.get("PREV_M"), errors="coerce").notna().any()
            if not prev_ok:
                st.warning("Não encontrei valores na coluna **PREVISTO MENSAL (%)** (ou header equivalente).")

            st.markdown("### Cards de Prazo")
            rr1 = st.columns(4)
            with rr1[0]:
                kpi_card_pct("Realizado acum.", k_real_acum, f"ref: {ref_month_label}")
            with rr1[1]:
                kpi_card_pct("Planejado acum.", k_plan_acum, f"ref: {ref_month_label}")
            with rr1[2]:
                c = PALETTE["good"] if (k_ader_acc is not None and k_ader_acc >= 1) else PALETTE["bad"]
                kpi_card_pct("Aderência acum.", k_ader_acc, "(Real acum ÷ Plan acum)", color=c)
            with rr1[3]:
                c = PALETTE["good"] if (k_ader_mes is not None and k_ader_mes >= 1) else PALETTE["bad"]
                kpi_card_pct("Aderência do mês", k_ader_mes, "(Real mês ÷ Plan mês)", color=c)

            st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

            rr2 = st.columns(3)
            with rr2[0]:
                kpi_card_pct("Realizado mensal", k_real_m, f"ref: {ref_month_label}")
            with rr2[1]:
                kpi_card_pct("Previsto mensal", k_prev_m, f"ref: {ref_month_label}")
            with rr2[2]:
                kpi_card_pct("Planejado mensal", k_plan_m, f"ref: {ref_month_label}")

            x = temp["MÊS"].tolist()

            t1, t2 = st.tabs(["Curva S (Acumulado)", "Curva Mensal + Aderência"])

            with t1:
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=x, y=planned_ac, mode="lines+markers", name="Planejado acum. (%)"))
                if prev_ok:
                    fig.add_trace(go.Scatter(x=x, y=previsto_ac, mode="lines+markers", name="Previsto acum. (%)"))
                fig.add_trace(go.Scatter(x=x, y=real_ac, mode="lines+markers", name="Realizado acum. (%)"))
                fig.update_layout(height=320, yaxis_title="%")
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

            with t2:
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=x, y=planned_m, mode="lines+markers", name="Planejado mês (%)"))
                if prev_ok:
                    fig.add_trace(go.Scatter(x=x, y=previsto_m, mode="lines+markers", name="Previsto mês (%)"))
                fig.add_trace(go.Scatter(x=x, y=real_m, mode="lines+markers", name="Realizado mês (%)"))
                fig.update_layout(height=290, yaxis_title="% (mensal)")
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

                # ✅ aderência mensal em barras (verde >=100, vermelho <100)
                st.markdown("#### Aderência mensal (Real ÷ Planejado) — base 100%")
                bar_colors = []
                for v in ader_mes_pct:
                    if v is None:
                        bar_colors.append(rgba(PALETTE["muted"], 0.25))
                    else:
                        bar_colors.append(PALETTE["good"] if v >= 100 else PALETTE["bad"])

                fig2 = go.Figure()
                fig2.add_trace(go.Bar(x=x, y=ader_mes_pct, marker_color=bar_colors, name="Aderência (%)"))
                fig2.add_hline(y=100, line_dash="dash", line_width=1)
                fig2.update_layout(height=260, yaxis_title="% (base 100)")
                st.plotly_chart(apply_plotly_theme(fig2), use_container_width=True)

    with right:
        # Cards top 3
        econ_items: list[tuple[str, float]] = []
        acres_items: list[tuple[str, float]] = []

        if not df_econ.empty and "VARIAÇÃO" in df_econ.columns:
            eco = df_econ.copy()
            eco["__v"] = pd.to_numeric(eco["VARIAÇÃO"], errors="coerce")
            eco = eco.dropna(subset=["__v"])
            eco["__abs"] = eco["__v"].abs()
            eco = eco.sort_values("__abs", ascending=False)
            for _, r in eco.head(3).iterrows():
                econ_items.append((str(r.get("DESCRIÇÃO", "")), float(r.get("__v", 0) or 0)))

        if not df_acres.empty and "VARIAÇÃO" in df_acres.columns:
            ac = df_acres.copy()
            ac["__v"] = pd.to_numeric(ac["VARIAÇÃO"], errors="coerce")
            ac = ac.dropna(subset=["__v"])
            ac["__abs"] = ac["__v"].abs()
            ac = ac.sort_values("__abs", ascending=False)
            for _, r in ac.head(3).iterrows():
                acres_items.append((str(r.get("DESCRIÇÃO", "")), float(r.get("__v", 0) or 0)))

        econ_rows = build_rows(econ_items, PALETTE["good"], prefix="")
        acres_rows = build_rows(acres_items, PALETTE["bad"], prefix="- ")

        st.markdown(
            card_resumo("PRINCIPAIS ECONOMIAS", "✅", econ_rows, rgba(GOOD, 0.35), rgba(GOOD, 0.10)),
            unsafe_allow_html=True,
        )
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        st.markdown(
            card_resumo("DESVIOS DO MÊS", "⚠️", acres_rows, rgba(BAD, 0.35), rgba(BAD, 0.10)),
            unsafe_allow_html=True,
        )
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        progress_card(k_real_acum, k_plan_acum, ref_month_label)

    st.divider()

    st.subheader("Detalhamento — Tabelas completas")
    c1, c2 = st.columns(2)

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
            with st.expander("Ver tabela (Economias)"):
                styled_dataframe(show_top.drop(columns=["__abs"], errors="ignore"))


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
    <div style="font-size:13px; font-weight:800;">{html.escape(desc)}</div>
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
