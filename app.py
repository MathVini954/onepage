# app.py
from __future__ import annotations

from pathlib import Path
import html

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

st.set_page_config(page_title="Prazo & Custo", layout="wide")

LOGOS_DIR = "assets/logos"


# ----------------------------
# Excel único (sem upload)
# ----------------------------
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
file_label = excel_path.name

obras = sheetnames(wb)
if not obras:
    st.error("Nenhuma aba de obra encontrada no Excel.")
    st.stop()

st.sidebar.title("Controle de Prazo e Custo")
obra = st.sidebar.selectbox("Obra (aba)", obras)
top_opt = st.sidebar.selectbox("Mostrar Top", ["5", "10", "Todas"], index=0)
top_n = None if top_opt == "Todas" else int(top_opt)

ws = wb[obra]


# ----------------------------
# Helpers UI
# ----------------------------
def fmt_brl_no_dec(v: float) -> str:
    s = f"{float(v):,.0f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def to_ratio(x) -> float | None:
    """Aceita 0-1 (ex 0.0496) ou 0-100 (ex 4.96)."""
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


def kpi_card(label: str, value: float | None):
    st.markdown(
        f"""
        <div style="
            border: 1px solid rgba(255,255,255,0.10);
            border-radius: 14px;
            padding: 12px 14px;
            background: rgba(255,255,255,0.03);
            height: 92px;
        ">
            <div style="font-size: 12px; opacity: 0.75; margin-bottom: 6px;">{html.escape(label)}</div>
            <div style="font-size: 24px; font-weight: 900; line-height: 1.05;">{html.escape(brl_compact(value))}</div>
            <div style="font-size: 11px; opacity: 0.65; margin-top: 6px;">{html.escape(fmt_brl(value))}</div>
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
<div style="display:flex; justify-content:space-between; align-items:center; padding:10px 0; border-top:1px solid rgba(255,255,255,0.07);">
  <div style="font-size:13px; font-weight:600;">{desc}</div>
  <div style="font-size:13px; font-weight:800; color:{color};">{prefix}{val_show}</div>
</div>
"""
    return rows


def card_resumo(title: str, icon: str, rows_html: str, border: str, bg: str) -> str:
    return f"""
<div style="border:1px solid {border}; background:{bg}; border-radius:16px; padding:14px 16px;">
  <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
    <div style="font-size:12px; opacity:0.85; font-weight:800; letter-spacing:0.3px;">{html.escape(title)}</div>
    <div style="font-size:12px;">{icon}</div>
  </div>
  {rows_html if rows_html else '<div style="opacity:0.65; font-size:12px;">Sem dados</div>'}
</div>
"""


def progress_card(real_ratio: float | None, planned_ratio: float | None, start_label: str, ref_month_label: str):
    real_ratio = clamp01(real_ratio)
    planned_ratio = clamp01(planned_ratio)

    real_pct = real_ratio * 100
    planned_pct = planned_ratio * 100

    st.markdown(
        f"""
        <div style="
          border:1px solid rgba(255,255,255,0.10);
          background:rgba(255,255,255,0.03);
          border-radius:16px;
          padding:14px 16px;
        ">
          <div style="display:flex; justify-content:space-between; align-items:center;">
            <div style="font-size:12px; opacity:0.85; font-weight:800;">Obra vs. Planejado</div>
            <div style="font-size:12px; opacity:0.65;">{html.escape(start_label)}</div>
          </div>

          <div style="font-size:12px; opacity:0.65; margin-top:6px;">
            referência: <b>{html.escape(ref_month_label)}</b>
          </div>

          <div style="margin-top:12px; display:flex; justify-content:space-between; align-items:flex-end;">
            <div>
              <div style="font-size:12px; opacity:0.75;">Progresso Real (acum.)</div>
              <div style="font-size:28px; font-weight:900; line-height:1;">{real_pct:.0f}%</div>
            </div>
            <div style="text-align:right;">
              <div style="font-size:12px; opacity:0.75;">Previsto (acum.)</div>
              <div style="font-size:16px; font-weight:900;">{planned_pct:.0f}%</div>
            </div>
          </div>

          <div style="margin-top:12px;">
            <div style="height:10px; background:rgba(255,255,255,0.08); border-radius:999px; position:relative;">
              <div style="width:{planned_pct:.2f}%; height:10px; background:rgba(59,130,246,0.35); border-radius:999px;"></div>
              <div style="width:{real_pct:.2f}%; height:10px; background:rgba(59,130,246,0.95); border-radius:999px; position:absolute; top:0; left:0;"></div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def adherence_card(month_label: str, ader_pct: float | None, delta_pp: float | None):
    if ader_pct is None:
        st.markdown(
            """
            <div style="border:1px solid rgba(255,255,255,0.10); background:rgba(255,255,255,0.03); border-radius:16px; padding:14px 16px;">
              <div style="font-size:12px; opacity:0.85; font-weight:800;">Aderência do mês</div>
              <div style="opacity:0.65; font-size:12px; margin-top:6px;">Sem dados suficientes (precisa Planejado Mês e Realizado Mês no mesmo mês).</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    color = "#22c55e" if ader_pct >= 100 else "#ef4444"
    sign = "+" if (delta_pp is not None and delta_pp >= 0) else ""
    delta_txt = "—" if delta_pp is None else f"{sign}{delta_pp:.1f} p.p.".replace(".", ",")

    st.markdown(
        f"""
        <div style="border:1px solid rgba(255,255,255,0.10); background:rgba(255,255,255,0.03); border-radius:16px; padding:14px 16px;">
          <div style="display:flex; justify-content:space-between; align-items:center;">
            <div style="font-size:12px; opacity:0.85; font-weight:800;">Aderência do mês</div>
            <div style="font-size:12px; opacity:0.65;">{html.escape(month_label)}</div>
          </div>
          <div style="margin-top:10px; display:flex; justify-content:space-between; align-items:flex-end;">
            <div style="font-size:28px; font-weight:900; color:{color}; line-height:1;">{ader_pct:.0f}%</div>
            <div style="font-size:12px; opacity:0.75; font-weight:700;">{delta_txt}</div>
          </div>
          <div style="font-size:12px; opacity:0.65; margin-top:6px;">(Realizado ÷ Planejado)</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def resumo_variacoes(df_acres: pd.DataFrame, df_econ: pd.DataFrame) -> dict:
    acres_var = pd.to_numeric(df_acres.get("VARIAÇÃO", pd.Series(dtype=float)), errors="coerce").fillna(0)
    econ_var = pd.to_numeric(df_econ.get("VARIAÇÃO", pd.Series(dtype=float)), errors="coerce").fillna(0)
    return {
        "total_acres": float(acres_var.abs().sum()) if not acres_var.empty else 0.0,
        "total_econ": float(econ_var.abs().sum()) if not econ_var.empty else 0.0,
        "saldo": float(acres_var.sum() + econ_var.sum()) if (not acres_var.empty or not econ_var.empty) else 0.0,
        "qtd_acres": int(len(df_acres)) if df_acres is not None else 0,
        "qtd_econ": int(len(df_econ)) if df_econ is not None else 0,
    }


def styled_dataframe(df: pd.DataFrame):
    if df.empty:
        st.info("Sem dados.")
        return
    tbl = df.copy()
    for col in ["ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO"]:
        if col in tbl.columns:
            tbl[col] = pd.to_numeric(tbl[col], errors="coerce")
    fmt_map = {}
    for col in ["ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO"]:
        if col in tbl.columns:
            fmt_map[col] = fmt_brl
    st.dataframe(tbl.style.format(fmt_map), use_container_width=True, hide_index=True)


# ----------------------------
# Header + logo
# ----------------------------
colL, colR = st.columns([1, 5])
with colL:
    logo_path = find_logo_path(obra, LOGOS_DIR)
    if logo_path:
        st.image(logo_path, use_container_width=True)
with colR:
    st.title(obra)
    st.caption(f"Arquivo: {file_label}")

st.divider()

# ----------------------------
# Ler blocos
# ----------------------------
resumo = read_resumo_financeiro(ws)
df_idx = read_indice(ws)
df_fin = read_financeiro(ws)
df_prazo = read_prazo(ws)
df_acres, df_econ = read_acrescimos_economias(ws)

# ----------------------------
# KPIs (2 linhas: 4 + 3) + espaçamento
# ----------------------------
row1 = st.columns(4)
with row1[0]:
    kpi_card("Orç. Inicial", resumo.get("ORÇAMENTO INICIAL (R$)"))
with row1[1]:
    kpi_card("Orç. Reajust.", resumo.get("ORÇAMENTO REAJUSTADO (R$)"))
with row1[2]:
    kpi_card("Desembolso Acum.", resumo.get("DESEMBOLSO ACUMULADO (R$)"))
with row1[3]:
    kpi_card("A Pagar", resumo.get("A PAGAR (R$)"))

st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)

row2 = st.columns(3)
with row2[0]:
    kpi_card("Saldo a Incorrer", resumo.get("SALDO A INCORRER (R$)"))
with row2[1]:
    kpi_card("Custo Final", resumo.get("CUSTO FINAL (R$)"))
with row2[2]:
    kpi_card("Variação", resumo.get("VARIAÇÃO (R$)"))

st.divider()


# ----------------------------
# Pré-cálculo do Prazo (Acumulado + Mensal + Aderência mensal)
# ----------------------------
temp = pd.DataFrame()
start_label = "—"
ref_month_label = "—"

planned_ratio = None
real_ratio = None

ader_pct_last = None
delta_pp_last = None
month_label_last = "—"

# séries mensais (para gráfico)
planned_m_pct = []
real_m_pct = []
adh_m_pct = []

# séries acumuladas (para curva S)
planned_acum_pct = []
real_acum_pct = []

if (not df_prazo.empty) and ("PLANEJADO MÊS (%)" in df_prazo.columns) and ("REALIZADO Mês (%)" in df_prazo.columns):
    temp = df_prazo[["MÊS", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)"]].copy()
    temp["PLANEJADO_M"] = temp["PLANEJADO MÊS (%)"].apply(to_ratio)
    temp["REAL_M"] = temp["REALIZADO Mês (%)"].apply(to_ratio)
    temp = temp.dropna(subset=["MÊS"]).reset_index(drop=True)

    if not temp.empty:
        start = temp["MÊS"].iloc[0]
        start_label = f"INÍCIO: {start.strftime('%b/%Y').lower()}"

    temp["PLANEJADO_ACUM"] = temp["PLANEJADO_M"].fillna(0).cumsum()
    temp["REAL_ACUM"] = temp["REAL_M"].fillna(0).cumsum()

    last_planned = temp["PLANEJADO_M"].last_valid_index()
    last_real = temp["REAL_M"].last_valid_index()

    # acumulado em %
    planned_acum_pct = (temp["PLANEJADO_ACUM"] * 100).tolist()
    real_acum_pct = (temp["REAL_ACUM"] * 100).tolist()

    # mensal em %
    planned_m_pct = (temp["PLANEJADO_M"] * 100).tolist()
    real_m_pct = (temp["REAL_M"] * 100).tolist()

    # “morrer” no último mês com dado
    if last_planned is not None:
        for i in range(last_planned + 1, len(planned_acum_pct)):
            planned_acum_pct[i] = None
            planned_m_pct[i] = None
    else:
        planned_acum_pct = [None] * len(planned_acum_pct)
        planned_m_pct = [None] * len(planned_m_pct)

    if last_real is not None:
        for i in range(last_real + 1, len(real_acum_pct)):
            real_acum_pct[i] = None
            real_m_pct[i] = None
    else:
        real_acum_pct = [None] * len(real_acum_pct)
        real_m_pct = [None] * len(real_m_pct)

    # aderência mensal (%) = Real mês / Planejado mês * 100
    adh_m_pct = [None] * len(temp)
    for i in range(len(temp)):
        p = temp.loc[i, "PLANEJADO_M"]
        r = temp.loc[i, "REAL_M"]
        if pd.notna(p) and pd.notna(r) and float(p) != 0:
            adh_m_pct[i] = (float(r) / float(p)) * 100

    # referência: último mês REAL preenchido (para progresso e aderência “do mês”)
    if last_real is not None:
        real_ratio = float(temp.loc[last_real, "REAL_ACUM"])
        planned_ratio = float(temp.loc[last_real, "PLANEJADO_ACUM"])

        m = temp.loc[last_real, "MÊS"]
        ref_month_label = m.strftime("%b/%Y").lower()
        month_label_last = ref_month_label

        p_m = temp.loc[last_real, "PLANEJADO_M"]
        r_m = temp.loc[last_real, "REAL_M"]
        if pd.notna(p_m) and pd.notna(r_m) and float(p_m) != 0:
            ader = float(r_m) / float(p_m)
            ader_pct_last = ader * 100
            delta_pp_last = ader_pct_last - 100


# ----------------------------
# Linha: gráficos + painel lateral
# ----------------------------
left, right = st.columns([2.2, 1])

with left:
    g1, g2 = st.columns(2)

    with g1:
        st.subheader("Índice Projetado (baseline 1,000)")
        if df_idx.empty:
            st.info("Sem dados do índice.")
        else:
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=df_idx["MÊS"],
                y=df_idx["ÍNDICE PROJETADO"],
                mode="lines+markers",
                name="Índice Projetado"
            ))
            fig.add_hline(y=1.0, line_dash="dash", line_width=1)
            fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig, use_container_width=True)

    with g2:
        st.subheader("Desembolso x Medido (mês a mês)")
        if df_fin.empty:
            st.info("Sem dados financeiros.")
        else:
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=df_fin["MÊS"],
                y=df_fin["DESEMBOLSO DO MÊS (R$)"],
                name="Desembolso",
                marker_color="rgba(59,130,246,0.85)"  # azul
            ))
            fig.add_trace(go.Bar(
                x=df_fin["MÊS"],
                y=df_fin["MEDIDO NO MÊS (R$)"],
                name="Medido",
                marker_color="rgba(34,197,94,0.85)"   # verde
            ))
            fig.update_layout(barmode="group", height=320, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig, use_container_width=True)

    st.subheader("Prazo — Curva S + Avanço mensal + Aderência")
    if temp.empty:
        st.info("Sem dados de prazo (precisa 'PLANEJADO MÊS (%)' e 'REALIZADO Mês (%)').")
    else:
        tab1, tab2 = st.tabs(["Curva S (Acumulado)", "Mensal (Avanço + Aderência)"])

        with tab1:
            x = temp["MÊS"].tolist()

            fig = go.Figure()
            fig.add_trace(go.Scatter(x=x, y=planned_acum_pct, mode="lines+markers", name="Planejado acum. (%)"))
            fig.add_trace(go.Scatter(x=x, y=real_acum_pct, mode="lines+markers", name="Real acum. (%)"))

            fig.update_layout(
                height=320,
                margin=dict(l=10, r=10, t=10, b=10),
                yaxis_title="%"
            )
            st.plotly_chart(fig, use_container_width=True)

        with tab2:
            x = temp["MÊS"].tolist()

            fig = go.Figure()

            # avanço individual (mês)
            fig.add_trace(go.Scatter(
                x=x, y=planned_m_pct, mode="lines+markers",
                name="Planejado mês (%)"
            ))
            fig.add_trace(go.Scatter(
                x=x, y=real_m_pct, mode="lines+markers",
                name="Real mês (%)"
            ))

            # aderência mês a mês (eixo direito)
            fig.add_trace(go.Scatter(
                x=x, y=adh_m_pct, mode="lines+markers",
                name="Aderência (%)",
                yaxis="y2"
            ))

            # linha base 100% na aderência
            if len(x) >= 2:
                fig.add_shape(
                    type="line",
                    x0=x[0], x1=x[-1],
                    y0=100, y1=100,
                    xref="x", yref="y2",
                    line=dict(dash="dash", width=1)
                )

            fig.update_layout(
                height=320,
                margin=dict(l=10, r=10, t=10, b=10),
                yaxis=dict(title="% (Avanço mensal)"),
                yaxis2=dict(
                    title="Aderência (%)",
                    overlaying="y",
                    side="right",
                    showgrid=False
                )
            )
            st.plotly_chart(fig, use_container_width=True)

with right:
    # ---------- Cards "Principais Economias" e "Desvios do mês" ----------
    econ_items: list[tuple[str, float]] = []
    acres_items: list[tuple[str, float]] = []

    if not df_econ.empty and "VARIAÇÃO" in df_econ.columns:
        econ_sorted = df_econ.copy()
        econ_sorted["__v"] = pd.to_numeric(econ_sorted["VARIAÇÃO"], errors="coerce")
        econ_sorted = econ_sorted.sort_values("__v", ascending=True)  # mais negativo primeiro
        for _, r in econ_sorted.head(3).iterrows():
            econ_items.append((str(r.get("DESCRIÇÃO", "")), float(r.get("__v", 0) or 0)))

    if not df_acres.empty and "VARIAÇÃO" in df_acres.columns:
        acres_sorted = df_acres.copy()
        acres_sorted["__v"] = pd.to_numeric(acres_sorted["VARIAÇÃO"], errors="coerce")
        acres_sorted = acres_sorted.sort_values("__v", ascending=False)
        for _, r in acres_sorted.head(3).iterrows():
            acres_items.append((str(r.get("DESCRIÇÃO", "")), float(r.get("__v", 0) or 0)))

    econ_rows = build_rows(econ_items, color="#22c55e", prefix="")
    acres_rows = build_rows(acres_items, color="#ef4444", prefix="- ")

    st.markdown(card_resumo(
        "PRINCIPAIS ECONOMIAS", "✅",
        econ_rows,
        border="rgba(34,197,94,0.25)",
        bg="rgba(34,197,94,0.08)"
    ), unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    st.markdown(card_resumo(
        "DESVIOS DO MÊS", "⚠️",
        acres_rows,
        border="rgba(239,68,68,0.25)",
        bg="rgba(239,68,68,0.08)"
    ), unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    progress_card(real_ratio, planned_ratio, start_label, ref_month_label)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    adherence_card(month_label_last, ader_pct_last, delta_pp_last)

st.divider()

# ----------------------------
# Resumo do mês (Economias x Desvios)
# ----------------------------
st.subheader("Resumo do mês (Economias x Desvios)")

stats = resumo_variacoes(df_acres, df_econ)
saldo = stats["saldo"]

colA, colB, colC = st.columns([1.2, 1, 1])
with colA:
    label = "Economia líquida" if saldo < 0 else "Acréscimo líquido"
    color = "#22c55e" if saldo < 0 else "#ef4444"
    st.markdown(
        f"""
        <div style="border:1px solid rgba(255,255,255,0.10); background:rgba(255,255,255,0.03); border-radius:16px; padding:14px 16px;">
          <div style="font-size:12px; opacity:0.75; font-weight:800;">{html.escape(label)}</div>
          <div style="font-size:28px; font-weight:900; color:{color};">{fmt_brl(abs(saldo))}</div>
          <div style="font-size:12px; opacity:0.65;">(saldo = desvios + economias)</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with colB:
    st.markdown(
        f"""
        <div style="border:1px solid rgba(255,255,255,0.10); background:rgba(255,255,255,0.03); border-radius:16px; padding:14px 16px;">
          <div style="font-size:12px; opacity:0.75; font-weight:800;">Total Economias</div>
          <div style="font-size:24px; font-weight:900; color:#22c55e;">{fmt_brl(stats["total_econ"])}</div>
          <div style="font-size:12px; opacity:0.65;">Itens: {stats["qtd_econ"]}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with colC:
    st.markdown(
        f"""
        <div style="border:1px solid rgba(255,255,255,0.10); background:rgba(255,255,255,0.03); border-radius:16px; padding:14px 16px;">
          <div style="font-size:12px; opacity:0.75; font-weight:800;">Total Desvios</div>
          <div style="font-size:24px; font-weight:900; color:#ef4444;">{fmt_brl(stats["total_acres"])}</div>
          <div style="font-size:12px; opacity:0.65;">Itens: {stats["qtd_acres"]}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.divider()

# ----------------------------
# Detalhamento — tabelas completas + barras em degradê
# ----------------------------
st.subheader("Detalhamento — Tabelas completas (com barras em degradê)")

t1, t2 = st.columns(2)

with t1:
    st.markdown("### ACRÉSCIMOS / DESVIOS (mês)")
    if df_acres.empty:
        st.info("Sem dados.")
    else:
        show = df_acres.copy()
        show["VARIAÇÃO"] = pd.to_numeric(show["VARIAÇÃO"], errors="coerce")
        show = show.sort_values("VARIAÇÃO", ascending=False)
        show_top = show.head(top_n) if top_n is not None else show

        top_bar = show.head(10).iloc[::-1]
        vals = top_bar["VARIAÇÃO"].abs()

        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=vals,
            y=top_bar["DESCRIÇÃO"],
            orientation="h",
            marker=dict(
                color=vals,
                colorscale=[[0, "rgba(239,68,68,0.25)"], [1, "rgba(239,68,68,1.0)"]],
                showscale=False
            ),
            name="R$"
        ))
        fig.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="R$")
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("Ver tabela completa (Acréscimos)"):
            styled_dataframe(show_top)

with t2:
    st.markdown("### ECONOMIAS (mês)")
    if df_econ.empty:
        st.info("Sem dados.")
    else:
        show = df_econ.copy()
        show["VARIAÇÃO"] = pd.to_numeric(show["VARIAÇÃO"], errors="coerce")
        show = show.sort_values("VARIAÇÃO", ascending=True)
        show_top = show.head(top_n) if top_n is not None else show

        top_bar = show.head(10).iloc[::-1]
        vals = top_bar["VARIAÇÃO"].abs()

        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=vals,
            y=top_bar["DESCRIÇÃO"],
            orientation="h",
            marker=dict(
                color=vals,
                colorscale=[[0, "rgba(34,197,94,0.25)"], [1, "rgba(34,197,94,1.0)"]],
                showscale=False
            ),
            name="R$"
        ))
        fig.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="R$")
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("Ver tabela completa (Economias)"):
            styled_dataframe(show_top)
