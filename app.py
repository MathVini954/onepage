# app.py
from __future__ import annotations

from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.graph_objects as go

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
from src.utils import fmt_brl, norm

st.set_page_config(page_title="Prazo & Custo", layout="wide")

LOGOS_DIR = "assets/logos"


# ----------------------------
# Helpers
# ----------------------------
def find_default_excel() -> Path | None:
    """Procura um único arquivo Excel na raiz (preferência: Excel.xlsm)."""
    candidates = [
        Path("Excel.xlsm"),
        Path("Excel.xlsx"),
        Path("excel.xlsm"),
        Path("excel.xlsx"),
    ]
    for p in candidates:
        if p.exists():
            return p
    return None


def to_ratio(x) -> float | None:
    """
    Converte valor de % para razão (0-1), aceitando:
    - 0.35 (já razão)
    - 35 (percentual)
    """
    if x is None:
        return None
    try:
        v = float(x)
    except Exception:
        return None
    return (v / 100.0) if v > 1.5 else v


def to_pct_display(series: pd.Series) -> pd.Series:
    """
    Converte uma série para percentuais 0-100 pra plot.
    Se a série parece estar em 0-1, multiplica por 100.
    """
    s = pd.to_numeric(series, errors="coerce")
    if s.dropna().empty:
        return s
    # Se valores típicos > 1.5, assumimos que já está em 0-100
    if s.dropna().median() > 1.5:
        return s
    return s * 100


def safe_sort_by_variacao(df: pd.DataFrame, ascending: bool) -> pd.DataFrame:
    if df.empty:
        return df
    if "VARIAÇÃO" not in df.columns:
        return df
    return df.sort_values("VARIAÇÃO", ascending=ascending)


def apply_top(df: pd.DataFrame, top_n: int | None) -> pd.DataFrame:
    if df.empty:
        return df
    if top_n is None:
        return df
    return df.head(top_n)


# ----------------------------
# Sidebar: arquivo único
# ----------------------------
st.sidebar.title("Controle de Prazo e Custo")

uploaded = st.sidebar.file_uploader(
    "Upload do Excel (opcional)", type=["xlsx", "xlsm"]
)

default_excel = find_default_excel()

if uploaded is not None:
    wb = load_wb(uploaded)
    file_label = uploaded.name
else:
    if default_excel is None:
        st.error(
            "Não achei Excel.xlsm / Excel.xlsx na raiz do projeto.\n\n"
            "✅ Solução: suba um arquivo chamado **Excel.xlsm** (ou **Excel.xlsx**) na raiz do repositório "
            "ou faça upload pela sidebar."
        )
        st.stop()
    wb = load_wb(default_excel)
    file_label = default_excel.name

obras = sheetnames(wb)
if not obras:
    st.error("Nenhuma aba de obra encontrada no Excel (fora LEIA-ME/README).")
    st.stop()

obra = st.sidebar.selectbox("Obra (aba)", obras)

top_opt = st.sidebar.selectbox("Mostrar Top", ["5", "10", "Todas"], index=0)
top_n = None if top_opt == "Todas" else int(top_opt)

# ----------------------------
# Header com logo
# ----------------------------
ws = wb[obra]

left, right = st.columns([1, 4])
with left:
    logo_path = find_logo_path(obra, LOGOS_DIR)
    if logo_path:
        st.image(logo_path, use_container_width=True)

with right:
    st.title(obra)
    st.caption(f"Arquivo: {file_label}")

st.divider()

# ----------------------------
# Ler blocos
# ----------------------------
resumo = read_resumo_financeiro(ws)
df_idx = read_indice(ws)
df_fin = read_financeiro(ws)
df_prazo = read_prazo(ws)  # layout novo (com planejado mês / comprometido / realizado)
df_acres, df_econ = read_acrescimos_economias(ws)

# ----------------------------
# Cards (Resumo financeiro)
# ----------------------------
c1, c2, c3, c4, c5, c6, c7 = st.columns(7)

with c1:
    st.metric("Orç. Inicial", fmt_brl(resumo.get("ORÇAMENTO INICIAL (R$)")))
with c2:
    st.metric("Orç. Reajust.", fmt_brl(resumo.get("ORÇAMENTO REAJUSTADO (R$)")))
with c3:
    st.metric("Desembolso Acum.", fmt_brl(resumo.get("DESEMBOLSO ACUMULADO (R$)")))
with c4:
    st.metric("A Pagar", fmt_brl(resumo.get("A PAGAR (R$)")))
with c5:
    st.metric("Saldo a Incorrer", fmt_brl(resumo.get("SALDO A INCORRER (R$)")))
with c6:
    st.metric("Custo Final", fmt_brl(resumo.get("CUSTO FINAL (R$)")))
with c7:
    st.metric("Variação", fmt_brl(resumo.get("VARIAÇÃO (R$)")))

st.divider()

# ----------------------------
# Gráficos (Índice + Financeiro)
# ----------------------------
g1, g2 = st.columns(2)

with g1:
    st.subheader("Índice Projetado (baseline 1,000)")
    if df_idx.empty:
        st.info("Não encontrei dados do bloco de ÍndICE.")
    else:
        fig = go.Figure()
        fig.add_trace(
            go.Scatter(
                x=df_idx["MÊS"],
                y=df_idx["ÍNDICE PROJETADO"],
                mode="lines+markers",
                name="Índice Projetado",
            )
        )
        fig.add_hline(y=1.0, line_dash="dash", line_width=1)
        fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)

with g2:
    st.subheader("Desembolso x Medido (mês a mês)")
    if df_fin.empty:
        st.info("Não encontrei dados do bloco FINANCEIRO.")
    else:
        fig = go.Figure()
        fig.add_trace(
            go.Bar(
                x=df_fin["MÊS"],
                y=df_fin["DESEMBOLSO DO MÊS (R$)"],
                name="Desembolso",
            )
        )
        fig.add_trace(
            go.Bar(
                x=df_fin["MÊS"],
                y=df_fin["MEDIDO NO MÊS (R$)"],
                name="Medido",
            )
        )
        fig.update_layout(
            barmode="group",
            height=360,
            margin=dict(l=10, r=10, t=10, b=10),
        )
        st.plotly_chart(fig, use_container_width=True)

st.divider()

# ----------------------------
# Prazo (Curva S + barra deveria/estou)
# ----------------------------
st.subheader("Prazo — Curva S + Progresso")

if df_prazo.empty:
    st.info("Não encontrei dados do bloco PRAZO.")
else:
    # Curva S: planejado acumulado vs real acumulado (se existir)
    fig = go.Figure()

    planned_acum_pct = to_pct_display(df_prazo["PLANEJADO ACUM. (%)"])
    fig.add_trace(
        go.Scatter(
            x=df_prazo["MÊS"],
            y=planned_acum_pct,
            mode="lines+markers",
            name="Planejado Acum (%)",
        )
    )

    if "REAL ACUM. (%)" in df_prazo.columns and df_prazo["REAL ACUM. (%)"].notna().any():
        real_acum_pct = to_pct_display(df_prazo["REAL ACUM. (%)"])
        fig.add_trace(
            go.Scatter(
                x=df_prazo["MÊS"],
                y=real_acum_pct,
                mode="lines+markers",
                name="Real Acum (%)",
            )
        )

    if "COMPROMETIDO ACUM. (%)" in df_prazo.columns and df_prazo["COMPROMETIDO ACUM. (%)"].notna().any():
        comp_acum_pct = to_pct_display(df_prazo["COMPROMETIDO ACUM. (%)"])
        fig.add_trace(
            go.Scatter(
                x=df_prazo["MÊS"],
                y=comp_acum_pct,
                mode="lines+markers",
                name="Comprometido Acum (%)",
            )
        )

    fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), yaxis_title="%")
    st.plotly_chart(fig, use_container_width=True)

    # Último mês preenchido: usamos Planejado Acum como base
    last_planned = df_prazo.dropna(subset=["PLANEJADO ACUM. (%)"]).tail(1)
    planned_val = float(last_planned["PLANEJADO ACUM. (%)"].iloc[0]) if not last_planned.empty else None
    planned_ratio = to_ratio(planned_val)

    real_ratio = None
    if "REAL ACUM. (%)" in df_prazo.columns and df_prazo["REAL ACUM. (%)"].notna().any():
        real_val = float(df_prazo["REAL ACUM. (%)"].dropna().tail(1).iloc[0])
        real_ratio = to_ratio(real_val)

    colA, colB, colC = st.columns([1, 1, 1])

    with colA:
        st.write(f"**Deveria estar (Planejado):** {to_pct_display(pd.Series([planned_val])).iloc[0]:.1f}%".replace(".", ",") if planned_val is not None else "—")
        st.progress(0 if planned_ratio is None else max(0, min(1, planned_ratio)))

    with colB:
        if real_val is not None:
            st.write(f"**Estou (Real):** {to_pct_display(pd.Series([real_val])).iloc[0]:.1f}%".replace(".", ","))
        else:
            st.write("**Estou (Real):** —")
        st.progress(0 if real_ratio is None else max(0, min(1, real_ratio)))

    with colC:
        gap_ratio = None if (planned_ratio is None or real_ratio is None) else (real_ratio - planned_ratio)
        if gap_ratio is None:
            st.write("**Gap:** —")
        else:
            st.write(f"**Gap:** {(gap_ratio*100):.1f}%".replace(".", ",") + " p.p.")
        # (opcional) mostrar tabela do prazo
        with st.expander("Ver tabela de prazo"):
            st.dataframe(df_prazo, use_container_width=True, hide_index=True)

st.divider()

# ----------------------------
# Acréscimos / Economias (Top 5/10/Todas)
# ----------------------------
t1, t2 = st.columns(2)

with t1:
    st.subheader("ACRÉSCIMOS (do mês)")
    if df_acres.empty:
        st.info("Tabela de ACRÉSCIMOS não encontrada.")
    else:
        show = safe_sort_by_variacao(df_acres, ascending=False)  # maiores primeiro
        show = apply_top(show, top_n)
        st.dataframe(show, use_container_width=True, hide_index=True)

with t2:
    st.subheader("ECONOMIAS (do mês)")
    if df_econ.empty:
        st.info("Tabela de ECONOMIAS não encontrada.")
    else:
        show = safe_sort_by_variacao(df_econ, ascending=True)  # mais negativo primeiro
        show = apply_top(show, top_n)
        st.dataframe(show, use_container_width=True, hide_index=True)
