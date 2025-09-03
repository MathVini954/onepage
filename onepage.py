import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(page_title="ONE_PAGE Dashboard", page_icon="🏗️", layout="wide")

# ===== Estilo CSS =====
st.markdown("""
<style>
    body { background-color: #1e1e2f; color: #fff; }
    .big-title { font-size:38px !important; font-weight:bold; text-align:center; margin-bottom:20px; color:#fff; }
    .card { padding:20px; border-radius:15px; background-color:#2e2e3e; box-shadow:0px 4px 15px rgba(0,0,0,0.5); text-align:center; margin-bottom:20px; }
    .card h4 { margin-bottom:10px; color:#fff; }
    .card p { font-size:16px; color:#fff; font-weight:bold; }
</style>
""", unsafe_allow_html=True)


# ===== Funções auxiliares =====
def format_money(val):
    try:
        return f"R$ {float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(val)

def format_percent(val):
    try:
        return f"{float(val)*100:.1f}%"
    except:
        return str(val)


# ===== Carregar planilha =====
uploaded_file = "ONE_PAGE.xlsx"
if not os.path.exists(uploaded_file):
    st.warning("📥 Coloque o arquivo 'ONE_PAGE.xlsx' no diretório do app.")
else:
    xls = pd.ExcelFile(uploaded_file)

    # Sidebar - filtro de obra
    st.sidebar.image("empresa_logo.png", width=200)
    obra_escolhida = st.sidebar.selectbox("🏢 Selecione a Obra", xls.sheet_names)

    # Logo da obra
    obra_logo = f"{obra_escolhida}.png"
    if os.path.exists(obra_logo):
        st.image(obra_logo, width=200)

    # Ler aba da obra
    df = pd.read_excel(uploaded_file, sheet_name=obra_escolhida, usecols="A:B", header=None, dtype=str)
    df[0] = df[0].str.strip()
    df[1] = df[1].str.strip()

    indicadores = {}
    for i in range(len(df)):
        key = df.iloc[i,0]
        value = df.iloc[i,1]
        try:
            value = float(value)
        except:
            try:
                value = pd.to_datetime(value)
            except:
                pass
        indicadores[key] = value

    # ===== Título =====
    st.markdown(f"<p class='big-title'>📊 Dashboard - {obra_escolhida}</p>", unsafe_allow_html=True)

    # ===== Cards principais =====
    col1, col2, col3, col4 = st.columns(4)
    col1.markdown(f"<div class='card'><h4>AC (m²)</h4><p>{indicadores.get('AC(m²)','-')}</p></div>", unsafe_allow_html=True)
    col2.markdown(f"<div class='card'><h4>AP (m²)</h4><p>{indicadores.get('AP(m²)','-')}</p></div>", unsafe_allow_html=True)
    col3.markdown(f"<div class='card'><h4>Efetivo</h4><p>{format_percent(indicadores.get('Ef',0))}</p></div>", unsafe_allow_html=True)
    col4.markdown(f"<div class='card'><h4>Total Unidades</h4><p>{indicadores.get('Total Unidades','-')}</p></div>", unsafe_allow_html=True)

    # ===== Avanço físico =====
    st.markdown("### 📈 Avanço Físico")
    planejado = indicadores.get("Avanço Físico Planejado",0)
    real = indicadores.get("Avanço Físico Real",0)
    aderencia = indicadores.get("Aderência Física",0)

    st.markdown(
        f"<p>Planejado: {format_percent(planejado)} | "
        f"Real: {format_percent(real)} | "
        f"Aderência: {format_percent(aderencia)}</p>",
        unsafe_allow_html=True
    )

    st.markdown(f"""
    <div style='position: relative; background-color: #555; border-radius: 15px; height: 30px;'>
        <!-- Preenchimento Real -->
        <div style='width:{real*100}%; background-color:#4caf50; height:100%; border-radius:15px; text-align:center; color:white; font-weight:bold; line-height:30px;'>
            {format_percent(real)}
        </div>
        <!-- Marcador Planejado -->
        <div style='position: absolute; left:{planejado*100}%; top:0; bottom:0; width:3px; background-color:red; border-radius:2px;'></div>
    </div>
    """, unsafe_allow_html=True)

    # ===== Financeiro =====
    st.markdown("### 💰 Indicadores Financeiros")
    col1, col2, col3, col4 = st.columns(4)
    col1.markdown(f"<div class='card'><h4>Desvio</h4><p>{format_percent(indicadores.get('Desvio',0))}</p></div>", unsafe_allow_html=True)
    col2.markdown(f"<div class='card'><h4>Desembolso</h4><p>{format_money(indicadores.get('Desembolso',0))}</p></div>", unsafe_allow_html=True)
    col3.markdown(f"<div class='card'><h4>Saldo</h4><p>{format_money(indicadores.get('Saldo',0))}</p></div>", unsafe_allow_html=True)
    col4.markdown(f"<div class='card'><h4>Índice Econômico</h4><p>{format_percent(indicadores.get('Índice Econômico',0))}</p></div>", unsafe_allow_html=True)

    # ===== Gráfico de colunas (exemplo simples) =====
    st.markdown("### 📊 Produção Mensal")
    df_col = pd.DataFrame({
        "Mês": ["Jan","Fev","Mar","Abr","Mai","Jun"],
        "Planejado": [80,85,90,92,95,98],
        "Real": [78,82,87,89,91,92]
    })
    fig = px.bar(df_col, x="Mês", y=["Planejado","Real"], barmode="group",
                 color_discrete_sequence=["#ff4d4d","#4caf50"])
    st.plotly_chart(fig, use_container_width=True)
