# app.py
from __future__ import annotations
from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.graph_objects as go

from src.excel_reader import (
    load_wb, sheetnames, read_resumo_financeiro, read_indice, read_financeiro,
    read_prazo, read_acrescimos_economias
)
from src.logos import find_logo_path
from src.utils import fmt_brl, fmt_pct

st.set_page_config(page_title="Prazo & Custo", layout="wide")

DATA_DIR = Path("data")
LOGOS_DIR = "assets/logos"

def list_data_files():
    if not DATA_DIR.exists():
        return []
    return sorted([p.name for p in DATA_DIR.glob("*.xlsx")])

st.sidebar.title("Controle de Prazo e Custo")

files = list_data_files()
uploaded = st.sidebar.file_uploader("Ou faça upload do Excel (.xlsx)", type=["xlsx"])

if uploaded is not None:
    wb = load_wb(uploaded)
    file_label = uploaded.name
else:
    if not files:
        st.warning("Coloque um Excel em /data (ex: Jan.26.xlsx) ou faça upload pela sidebar.")
        st.stop()
    chosen = st.sidebar.selectbox("Arquivo do mês (pasta /data)", files, index=len(files)-1)
    wb = load_wb(DATA_DIR / chosen)
    file_label = chosen

obras = sheetnames(wb)
obra = st.sidebar.selectbox("Obra (aba)", obras)

top_opt = st.sidebar.selectbox("Mostrar Top", ["5", "10", "Todas"], index=0)
top_n = None if top_opt == "Todas" else int(top_opt)

ws = wb[obra]

# Header com logo
colL, colR = st.columns([1, 4])
with colL:
    logo = find_logo_path(obra, LOGOS_DIR)
    if logo:
        st.image(logo, use_container_width=True)
with colR:
    st.title(f"{obra}")
    st.caption(f"Arquivo: {file_label}")

# Leitura dos blocos
resumo = read_resumo_financeiro(ws)
df_idx = read_indice(ws)
df_fin = read_financeiro(ws)
df_prazo = read_prazo(ws)
df_acres, df_econ = read_acrescimos_economias(ws)

# -------------------------
# KPIs (cards)
# -------------------------
c1, c2, c3, c4 = st.columns(4)

orc_ini = resumo.get("ORÇAMENTO INICIAL (R$)")
orc_reaj = resumo.get("ORÇAMENTO REAJUSTADO (R$)")
des_acum = resumo.get("DESEMBOLSO ACUMULADO (R$)")
custo_final = resumo.get("CUSTO FINAL (R$)")
variacao = resumo.get("VARIAÇÃO (R$)")

with c1:
    st.metric("Orçamento Reajustado", fmt_brl(orc_reaj), delta=None)
with c2:
    st.metric("Desembolso Acumulado", fmt_brl(des_acum), delta=None)
with c3:
    st.metric("Custo Final", fmt_brl(custo_final), delta=None)
with c4:
    st.metric("Variação", fmt_brl(variacao), delta=None)

st.divider()

# -------------------------
# Gráficos: Índice / Financeiro / Prazo
# -------------------------
g1, g2 = st.columns(2)

with g1:
    st.subheader("Índice Projetado (baseline 1,000)")
    if df_idx.empty:
        st.info("Bloco de Índice não encontrado ou sem dados.")
    else:
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df_idx["MÊS"], y=df_idx["ÍNDICE PROJETADO"],
            mode="lines+markers", name="Índice Projetado"
        ))
        fig.add_hline(y=1.0, line_dash="dash", line_width=1)
        fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)

with g2:
    st.subheader("Desembolso x Medido (mês a mês)")
    if df_fin.empty:
        st.info("Bloco Financeiro não encontrado ou sem dados.")
    else:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df_fin["MÊS"], y=df_fin["DESEMBOLSO DO MÊS (R$)"], name="Desembolso"))
        fig.add_trace(go.Bar(x=df_fin["MÊS"], y=df_fin["MEDIDO NO MÊS (R$)"], name="Medido"))
        fig.update_layout(barmode="group", height=360, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)

st.subheader("Prazo — Curva S + Progresso")
if df_prazo.empty:
    st.info("Bloco de Prazo não encontrado ou sem dados.")
else:
    # Curva S (Planejado Acum x Real Acum)
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df_prazo["MÊS"], y=df_prazo["PLANEJADO ACUM. (%)"],
        mode="lines+markers", name="Planejado Acum"
    ))

    if "REAL ACUM. (%)" in df_prazo.columns:
        fig.add_trace(go.Scatter(
            x=df_prazo["MÊS"], y=df_prazo["REAL ACUM. (%)"],
            mode="lines+markers", name="Real Acum"
        ))

    if "COMPROMETIDO ACUM. (%)" in df_prazo.columns:
        fig.add_trace(go.Scatter(
            x=df_prazo["MÊS"], y=df_prazo["COMPROMETIDO ACUM. (%)"],
            mode="lines+markers", name="Comprometido Acum"
        ))

    fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

    # Progresso (último mês preenchido)
    last = df_prazo.dropna(subset=["PLANEJADO ACUM. (%)"]).tail(1)
    planned = float(last["PLANEJADO ACUM. (%)"].iloc[0]) if not last.empty else None

    real = None
    if "REAL ACUM. (%)" in df_prazo.columns:
        real = float(df_prazo["REAL ACUM. (%)"].dropna().tail(1).iloc[0]) if df_prazo["REAL ACUM. (%)"].notna().any() else None

    pcol1, pcol2, pcol3 = st.columns([1, 1, 1])
    with pcol1:
        st.write(f"**Deveria estar (Planejado):** {fmt_pct(planned, scale_0_1=True)}")
        st.progress(0 if planned is None else max(0, min(1, planned)))
    with pcol2:
        st.write(f"**Estou (Real):** {fmt_pct(real, scale_0_1=True)}")
        st.progress(0 if real is None else max(0, min(1, real)))
    with pcol3:
        gap = None if (planned is None or real is None) else (real - planned)
        st.write(f"**Gap:** {fmt_pct(gap, scale_0_1=True)}")

st.divider()

# -------------------------
# ACRÉSCIMOS / ECONOMIAS — Top 5/10/Todas
# -------------------------
t1, t2 = st.columns(2)

def cut_top(df: pd.DataFrame, n: int | None, ascending=False):
    if df.empty:
        return df
    df2 = df.copy()
    if "VARIAÇÃO" in df2.columns:
        df2 = df2.sort_values("VARIAÇÃO", ascending=ascending)
    if n is None:
        return df2
    return df2.head(n)

with t1:
    st.subheader("ACRÉSCIMOS")
    if df_acres.empty:
        st.info("Tabela de Acréscimos não encontrada.")
    else:
        show = cut_top(df_acres, top_n, ascending=False)  # maiores primeiro
        st.dataframe(show, use_container_width=True, hide_index=True)

with t2:
    st.subheader("ECONOMIAS")
    if df_econ.empty:
        st.info("Tabela de Economias não encontrada.")
    else:
        show = cut_top(df_econ, top_n, ascending=True)  # mais negativo primeiro
        st.dataframe(show, use_container_width=True, hide_index=True)
