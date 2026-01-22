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
from src.utils import fmt_brl


st.set_page_config(page_title="Prazo & Custo", layout="wide")

LOGOS_DIR = "assets/logos"


# ----------------------------
# Helpers
# ----------------------------
def find_default_excel() -> Path | None:
    for name in ["Excel.xlsm", "Excel.xlsx", "excel.xlsm", "excel.xlsx"]:
        p = Path(name)
        if p.exists():
            return p
    return None


def to_ratio(x) -> float | None:
    """
    Converte valor para razão 0-1, aceitando:
    - 0.0496 (4,96% em Excel)
    - 4.96 (4,96)
    - 49.6 (49,6)
    """
    if x is None:
        return None
    try:
        v = float(x)
    except Exception:
        return None
    if v <= 1.5:
        return v
    return v / 100.0


def clamp01(v: float | None) -> float:
    if v is None:
        return 0.0
    return max(0.0, min(1.0, float(v)))


def brl_compact(v: float | None) -> str:
    """R$ compacto pra caber bonito no card."""
    if v is None:
        return "—"
    n = float(v)
    absn = abs(n)
    if absn >= 1_000_000_000:
        return f"R$ {n/1_000_000_000:.2f} bi".replace(".", ",")
    if absn >= 1_000_000:
        return f"R$ {n/1_000_000:.2f} mi".replace(".", ",")
    if absn >= 1_000:
        return f"R$ {n/1_000:.2f} mil".replace(".", ",")
    return fmt_brl(n)


def currency_style_df(df: pd.DataFrame, cols: list[str]):
    if df.empty:
        return df
    df2 = df.copy()
    for c in cols:
        if c in df2.columns:
            df2[c] = pd.to_numeric(df2[c], errors="coerce")
    return df2


def top_list_card(title: str, items: list[tuple[str, float]], kind: str):
    """
    kind:
      - "good": economias (verde)
      - "bad": desvios (vermelho)
    """
    if kind == "good":
        badge = "✅"
        color = "#22c55e"
        bg = "rgba(34,197,94,0.08)"
        border = "rgba(34,197,94,0.25)"
    else:
        badge = "⚠️"
        color = "#ef4444"
        bg = "rgba(239,68,68,0.08)"
        border = "rgba(239,68,68,0.25)"

    rows_html = ""
    if not items:
        rows_html = '<div style="opacity:0.65; font-size:12px;">Sem dados</div>'
    else:
        for desc, val in items:
            desc_show = (desc[:42] + "…") if len(desc) > 43 else desc
            amount = fmt_brl(abs(val))
            # Para o card: economia mostra positivo, desvio mostra como negativo (visual)
            amount_show = f"{amount}" if kind == "good" else f"- {amount}"
            rows_html += f"""
            <div style="display:flex; justify-content:space-between; align-items:center; padding:10px 0; border-top:1px solid rgba(255,255,255,0.07);">
              <div style="font-size:13px; font-weight:600;">{desc_show}</div>
              <div style="font-size:13px; font-weight:800; color:{color};">{amount_show}</div>
            </div>
            """

    st.markdown(
        f"""
        <div style="
          border:1px solid {border};
          background:{bg};
          border-radius:16px;
          padding:14px 16px;
        ">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
            <div style="font-size:12px; opacity:0.8; font-weight:700; letter-spacing:0.3px;">{title}</div>
            <div style="font-size:12px;">{badge}</div>
          </div>
          {rows_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def progress_card(real_ratio: float | None, planned_ratio: float | None, start_label: str):
    real_ratio = 0.0 if real_ratio is None else float(real_ratio)
    planned_ratio = 0.0 if planned_ratio is None else float(planned_ratio)

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
            <div style="font-size:12px; opacity:0.8; font-weight:700;">Obra vs. Planejado</div>
            <div style="font-size:12px; opacity:0.65;">{start_label}</div>
          </div>

          <div style="margin-top:12px; display:flex; justify-content:space-between; align-items:flex-end;">
            <div>
              <div style="font-size:12px; opacity:0.75;">Progresso Real</div>
              <div style="font-size:28px; font-weight:800; line-height:1;">{real_pct:.0f}%</div>
            </div>
            <div style="text-align:right;">
              <div style="font-size:12px; opacity:0.75;">Previsto</div>
              <div style="font-size:16px; font-weight:800;">{planned_pct:.0f}%</div>
            </div>
          </div>

          <div style="margin-top:12px;">
            <div style="height:10px; background:rgba(255,255,255,0.08); border-radius:999px; position:relative;">
              <div style="width:{max(0,min(100,planned_pct)):.2f}%; height:10px; background:rgba(59,130,246,0.35); border-radius:999px;"></div>
              <div style="width:{max(0,min(100,real_pct)):.2f}%; height:10px; background:rgba(59,130,246,0.95); border-radius:999px; position:absolute; top:0; left:0;"></div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def resumo_variacoes(df_acres: pd.DataFrame, df_econ: pd.DataFrame) -> dict:
    acres_var = pd.to_numeric(df_acres.get("VARIAÇÃO", pd.Series(dtype=float)), errors="coerce").fillna(0)
    econ_var = pd.to_numeric(df_econ.get("VARIAÇÃO", pd.Series(dtype=float)), errors="coerce").fillna(0)

    total_acres = float(acres_var.abs().sum()) if not acres_var.empty else 0.0
    total_econ = float(econ_var.abs().sum()) if not econ_var.empty else 0.0
    saldo = float(acres_var.sum() + econ_var.sum()) if (not acres_var.empty or not econ_var.empty) else 0.0

    # maiores itens
    top_acres = None
    top_econ = None
    if not df_acres.empty and "VARIAÇÃO" in df_acres.columns:
        i = pd.to_numeric(df_acres["VARIAÇÃO"], errors="coerce").abs().idxmax()
        if pd.notna(i):
            top_acres = (str(df_acres.loc[i, "DESCRIÇÃO"]), float(df_acres.loc[i, "VARIAÇÃO"]))
    if not df_econ.empty and "VARIAÇÃO" in df_econ.columns:
        i = pd.to_numeric(df_econ["VARIAÇÃO"], errors="coerce").abs().idxmax()
        if pd.notna(i):
            top_econ = (str(df_econ.loc[i, "DESCRIÇÃO"]), float(df_econ.loc[i, "VARIAÇÃO"]))

    return {
        "total_acres": total_acres,
        "total_econ": total_econ,
        "saldo": saldo,
        "qtd_acres": int(len(df_acres)) if df_acres is not None else 0,
        "qtd_econ": int(len(df_econ)) if df_econ is not None else 0,
        "top_acres": top_acres,
        "top_econ": top_econ,
    }


# ----------------------------
# Excel único (sem upload)
# ----------------------------
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
# Linha 1: Resumo financeiro (cards compactos)
# ----------------------------
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric("Orçamento Reajustado", brl_compact(resumo.get("ORÇAMENTO REAJUSTADO (R$)")))
with c2:
    st.metric("Desembolso Acumulado", brl_compact(resumo.get("DESEMBOLSO ACUMULADO (R$)")))
with c3:
    st.metric("Custo Final", brl_compact(resumo.get("CUSTO FINAL (R$)")))
with c4:
    st.metric("Variação (R$)", fmt_brl(resumo.get("VARIAÇÃO (R$)")))

# ----------------------------
# Linha 2: Índice/Financeiro + Painel lateral (Economias/Desvios + Progresso)
# ----------------------------
left, right = st.columns([2.1, 1])

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
                    name="Índice Projetado",
                )
            )
            fig.add_hline(y=1.0, line_dash="dash", line_width=1)
            fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig, use_container_width=True)

    with g2:
        st.subheader("Desembolso x Medido (mês a mês)")
        if df_fin.empty:
            st.info("Sem dados financeiros.")
        else:
            fig = go.Figure()
            fig.add_trace(go.Bar(x=df_fin["MÊS"], y=df_fin["DESEMBOLSO DO MÊS (R$)"], name="Desembolso"))
            fig.add_trace(go.Bar(x=df_fin["MÊS"], y=df_fin["MEDIDO NO MÊS (R$)"], name="Medido"))
            fig.update_layout(barmode="group", height=320, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig, use_container_width=True)

    st.subheader("Prazo — Curva S (Planejado x Real)")
    if df_prazo.empty:
        st.info("Sem dados de prazo.")
    else:
        # Ignora PLANEJADO ACUM. do Excel. Usa somente mês (%):
        if "PLANEJADO MÊS (%)" not in df_prazo.columns or "REALIZADO Mês (%)" not in df_prazo.columns:
            st.warning("Não achei as colunas: 'PLANEJADO MÊS (%)' e 'REALIZADO Mês (%)' no bloco de prazo.")
        else:
            temp = df_prazo[["MÊS", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)"]].copy()
            temp["PLANEJADO_M"] = temp["PLANEJADO MÊS (%)"].apply(to_ratio)
            temp["REAL_M"] = temp["REALIZADO Mês (%)"].apply(to_ratio)

            temp["PLANEJADO_ACUM"] = temp["PLANEJADO_M"].fillna(0).cumsum()
            temp["REAL_ACUM"] = temp["REAL_M"].fillna(0).cumsum()

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=temp["MÊS"], y=temp["PLANEJADO_ACUM"] * 100,
                mode="lines+markers", name="Planejado (%)"
            ))
            fig.add_trace(go.Scatter(
                x=temp["MÊS"], y=temp["REAL_ACUM"] * 100,
                mode="lines+markers", name="Real (%)"
            ))
            fig.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10), yaxis_title="%")
            st.plotly_chart(fig, use_container_width=True)

with right:
    # Principais economias e desvios do mês (top cards)
    if df_econ is None: df_econ = pd.DataFrame()
    if df_acres is None: df_acres = pd.DataFrame()

    # Ordenação:
    # Economias: pega VARIAÇÃO mais "forte" por magnitude, mas exibe positivo
    econ_sorted = df_econ.copy()
    if not econ_sorted.empty and "VARIAÇÃO" in econ_sorted.columns:
        econ_sorted["__v"] = pd.to_numeric(econ_sorted["VARIAÇÃO"], errors="coerce")
        econ_sorted = econ_sorted.sort_values("__v", ascending=True)  # mais negativo primeiro (normalmente)
    econ_items = []
    if not econ_sorted.empty:
        for _, r in econ_sorted.head(3).iterrows():
            econ_items.append((str(r.get("DESCRIÇÃO", "")), float(r.get("VARIAÇÃO", 0) or 0)))

    acres_sorted = df_acres.copy()
    if not acres_sorted.empty and "VARIAÇÃO" in acres_sorted.columns:
        acres_sorted["__v"] = pd.to_numeric(acres_sorted["VARIAÇÃO"], errors="coerce")
        acres_sorted = acres_sorted.sort_values("__v", ascending=False)  # maior primeiro
    acres_items = []
    if not acres_sorted.empty:
        for _, r in acres_sorted.head(3).iterrows():
            acres_items.append((str(r.get("DESCRIÇÃO", "")), float(r.get("VARIAÇÃO", 0) or 0)))

    top_list_card("PRINCIPAIS ECONOMIAS", econ_items, kind="good")
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    top_list_card("DESVIOS DO MÊS", acres_items, kind="bad")
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    # Progresso card (estilo print)
    real_ratio = None
    planned_ratio = None
    start_label = "—"

    if not df_prazo.empty and "PLANEJADO MÊS (%)" in df_prazo.columns and "REALIZADO Mês (%)" in df_prazo.columns:
        temp = df_prazo[["MÊS", "PLANEJADO MÊS (%)", "REALIZADO Mês (%)"]].copy()
        temp["PLANEJADO_M"] = temp["PLANEJADO MÊS (%)"].apply(to_ratio)
        temp["REAL_M"] = temp["REALIZADO Mês (%)"].apply(to_ratio)
        temp = temp.dropna(subset=["MÊS"])

        if not temp.empty:
            start = temp["MÊS"].iloc[0]
            start_label = f"INÍCIO: {start.strftime('%b/%Y').lower()}"

            temp["PLANEJADO_ACUM"] = temp["PLANEJADO_M"].fillna(0).cumsum()
            temp["REAL_ACUM"] = temp["REAL_M"].fillna(0).cumsum()

            planned_ratio = float(temp["PLANEJADO_ACUM"].iloc[-1])
            real_ratio = float(temp["REAL_ACUM"].iloc[-1])

    progress_card(clamp01(real_ratio), clamp01(planned_ratio), start_label)

# ----------------------------
# Resumo inteligente (análise do mês)
# ----------------------------
st.divider()
st.subheader("Resumo do mês (Economias x Desvios)")

stats = resumo_variacoes(df_acres, df_econ)

saldo = stats["saldo"]
saldo_label = "Economia líquida" if saldo < 0 else "Acréscimo líquido"
saldo_color = "#22c55e" if saldo < 0 else "#ef4444"

r1, r2, r3 = st.columns([1.2, 1, 1])
with r1:
    st.markdown(
        f"""
        <div style="border:1px solid rgba(255,255,255,0.10); background:rgba(255,255,255,0.03); border-radius:16px; padding:14px 16px;">
          <div style="font-size:12px; opacity:0.75; font-weight:700;">{saldo_label}</div>
          <div style="font-size:28px; font-weight:900; color:{saldo_color};">{fmt_brl(abs(saldo))}</div>
          <div style="font-size:12px; opacity:0.65;">(saldo = desvios + economias)</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with r2:
    st.markdown(
        f"""
        <div style="border:1px solid rgba(255,255,255,0.10); background:rgba(255,255,255,0.03); border-radius:16px; padding:14px 16px;">
          <div style="font-size:12px; opacity:0.75; font-weight:700;">Total Economias</div>
          <div style="font-size:24px; font-weight:900; color:#22c55e;">{fmt_brl(stats["total_econ"])}</div>
          <div style="font-size:12px; opacity:0.65;">Itens: {stats["qtd_econ"]}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with r3:
    st.markdown(
        f"""
        <div style="border:1px solid rgba(255,255,255,0.10); background:rgba(255,255,255,0.03); border-radius:16px; padding:14px 16px;">
          <div style="font-size:12px; opacity:0.75; font-weight:700;">Total Desvios</div>
          <div style="font-size:24px; font-weight:900; color:#ef4444;">{fmt_brl(stats["total_acres"])}</div>
          <div style="font-size:12px; opacity:0.65;">Itens: {stats["qtd_acres"]}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

bullets = []
if stats["top_econ"]:
    bullets.append(f"Maior economia: **{stats['top_econ'][0]}** ({fmt_brl(abs(stats['top_econ'][1]))})")
if stats["top_acres"]:
    bullets.append(f"Maior desvio: **{stats['top_acres'][0]}** ({fmt_brl(abs(stats['top_acres'][1]))})")
if not bullets:
    bullets.append("Sem itens suficientes para destacar maiores impactos.")

st.markdown("- " + "\n- ".join(bullets))

# ----------------------------
# Tabelas completas + barras (visual)
# ----------------------------
st.divider()
st.subheader("Detalhamento — Economias e Desvios (tabelas completas)")

t1, t2 = st.columns(2)

with t1:
    st.markdown("### ACRÉSCIMOS / DESVIOS (mês)")
    if df_acres.empty:
        st.info("Sem dados.")
    else:
        show = df_acres.copy()
        show["VARIAÇÃO"] = pd.to_numeric(show["VARIAÇÃO"], errors="coerce")
        show = show.sort_values("VARIAÇÃO", ascending=False)
        if top_n is not None:
            st.caption(f"Top {top_n} (por variação)")
        fig = go.Figure()
        top_bar = show.head(10).iloc[::-1]  # inverte pra ficar bonito
        fig.add_trace(go.Bar(
            x=top_bar["VARIAÇÃO"].abs(),
            y=top_bar["DESCRIÇÃO"],
            orientation="h",
            name="R$"
        ))
        fig.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="R$")
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("Ver tabela completa (Acréscimos)"):
            tbl = show.copy()
            # formatação brl em colunas numéricas
            for col in ["ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO"]:
                if col in tbl.columns:
                    tbl[col] = pd.to_numeric(tbl[col], errors="coerce")
            st.dataframe(
                tbl.style.format({
                    "ORÇAMENTO INICIAL": fmt_brl,
                    "ORÇAMENTO REAJUSTADO": fmt_brl,
                    "CUSTO FINAL": fmt_brl,
                    "VARIAÇÃO": fmt_brl,
                }),
                use_container_width=True
            )

with t2:
    st.markdown("### ECONOMIAS (mês)")
    if df_econ.empty:
        st.info("Sem dados.")
    else:
        show = df_econ.copy()
        show["VARIAÇÃO"] = pd.to_numeric(show["VARIAÇÃO"], errors="coerce")
        show = show.sort_values("VARIAÇÃO", ascending=True)  # normalmente mais negativo = maior economia
        if top_n is not None:
            st.caption(f"Top {top_n} (por variação)")
        fig = go.Figure()
        top_bar = show.head(10).iloc[::-1]
        fig.add_trace(go.Bar(
            x=top_bar["VARIAÇÃO"].abs(),
            y=top_bar["DESCRIÇÃO"],
            orientation="h",
            name="R$"
        ))
        fig.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="R$")
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("Ver tabela completa (Economias)"):
            tbl = show.copy()
            for col in ["ORÇAMENTO INICIAL", "ORÇAMENTO REAJUSTADO", "CUSTO FINAL", "VARIAÇÃO"]:
                if col in tbl.columns:
                    tbl[col] = pd.to_numeric(tbl[col], errors="coerce")
            st.dataframe(
                tbl.style.format({
                    "ORÇAMENTO INICIAL": fmt_brl,
                    "ORÇAMENTO REAJUSTADO": fmt_brl,
                    "CUSTO FINAL": fmt_brl,
                    "VARIAÇÃO": fmt_brl,
                }),
                use_container_width=True
            )
