import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import os
from datetime import datetime

# -------------------- Configuração da página --------------------
st.set_page_config(
    page_title="ONE PAGE - ENGENHARIA",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# -------------------- Forçar dark mode + CSS geral --------------------
st.markdown("""
<style>
/* -------------------- Paleta da Empresa -------------------- */
body, .stApp {
    background-color: #1A253C; /* Azul Escuro */
    color: #E5E8DD; /* Branco Nuvem */
}

/* -------------------- Sidebar hover -------------------- */
.css-1d391kg {
    width: 3rem; /* minimizada */
    transition: width 0.3s ease;
    overflow-x: hidden;
    background-color: #1A253C; /* Azul Escuro */
    color: #E5E8DD;
}
.css-1d391kg:hover {
    width: 18rem; /* aberta */
}
.css-1d391kg > .block-container {
    opacity: 0;
    transition: opacity 0.3s ease;
}
.css-1d391kg:hover > .block-container {
    opacity: 1;
}

/* -------------------- Títulos de Seção -------------------- */
.sub-header {
    font-size: 1.8rem;
    color: #E5E8DD; /* Branco Nuvem */
    font-weight: 700;
    margin: 2rem 0 1.5rem 0;
    padding-bottom: 0.5rem;
    border-bottom: 2px solid #5DAAAB; /* Azul Céu */
}

/* -------------------- Cards -------------------- */
.metric-card {
    background-color: #1A253C; /* Azul Escuro */
    border-radius: 0.75rem;
    padding: 1rem;
    box-shadow: 0 4px 6px rgba(0,0,0,0.3);
    height: 100%;
    border-left: 5px solid #5DAAAB; /* Azul Céu */
    margin-bottom: 1rem;
}
.metric-title {
    font-size: 1rem;
    color: #5DAAAB; /* Azul Céu */
    font-weight: 600;
    margin-bottom: 0.5rem;
}
.metric-value {
    font-size: 1.6rem;
    color: #E5E8DD; /* Branco Nuvem */
    font-weight: 800;
}

/* -------------------- Seções -------------------- */
.section-container {
    background-color: #1A253C; /* Azul Escuro */
    border-radius: 1rem;
    padding: 1.5rem;
    margin-bottom: 2rem;
    box-shadow: 0 4px 6px rgba(0,0,0,0.4);
}

/* -------------------- Barra de Progresso -------------------- */
.progress-wrapper {
    background-color: #1A253C;
    border-radius: 20px;
    padding: 5px;
    width: 100%;
}
.progress-bar {
    height: 30px;
    border-radius: 20px;
    text-align: center;
    font-weight: bold;
    color: #E5E8DD; /* Branco Nuvem */
    line-height: 30px;
}
</style>
""", unsafe_allow_html=True)

# -------------------- Sidebar --------------------
logo_empresa_path = "empresa_logo.png"
if os.path.exists(logo_empresa_path):
    st.sidebar.image(logo_empresa_path, width=200)

st.sidebar.markdown("### 📂 Selecione a Obra - Jul/25")

file_path = "ONE_PAGE.xlsx"
if not os.path.exists(file_path):
    st.error("❌ Arquivo ONE_PAGE.xlsx não encontrado no diretório!")
    st.stop()

excel_file = pd.ExcelFile(file_path)
sheet_names = excel_file.sheet_names
selected_sheet = st.sidebar.selectbox("Obra:", sheet_names)

# -------------------- Carregar dados --------------------
df = pd.read_excel(file_path, sheet_name=selected_sheet)
df_clean = df.iloc[:, [0, 1]].dropna()
df_clean.columns = ['Metrica', 'Valor']

dados = {str(row['Metrica']).strip(): row['Valor'] for _, row in df_clean.iterrows()}

def get_value(key, default="N/A"):
    return dados.get(key, default)

def format_money(value):
    if isinstance(value, (int, float)):
        return f"R$ {value:,.0f}".replace(',', '.')
    return str(value)

def format_percent(value):
    if isinstance(value, (int, float)) and value <= 1:
        return f"{value*100:.1f}%"
    elif isinstance(value, (int, float)):
        return f"{value:.1f}%"
    return str(value)

def to_float(val):
    if isinstance(val, str):
        try:
            return float(val.replace('R$', '').replace('.', '').replace(',', '.'))
        except:
            return 0
    return val if isinstance(val, (int,float)) else 0

# -------------------- Logo da obra --------------------
obra_logo_path = f"{selected_sheet}.png"
if os.path.exists(obra_logo_path):
    st.image(obra_logo_path, width=350)

# -------------------- Métricas Principais --------------------
st.markdown('<p class="sub-header">📊 Dados do Empreendimento</p>', unsafe_allow_html=True)
cols = st.columns(4)
cols[0].markdown(f'<div class="metric-card"><p class="metric-title">Área Construída (m²)</p><p class="metric-value">{get_value("Área Construída (m²)")}</p></div>', unsafe_allow_html=True)
cols[1].markdown(f'<div class="metric-card"><p class="metric-title">Área Privativa (m²)</p><p class="metric-value">{get_value("Área Privativa (m²)")}</p></div>', unsafe_allow_html=True)
cols[2].markdown(f'<div class="metric-card"><p class="metric-title">Eficiência</p><p class="metric-value">{format_percent(get_value("Eficiência"))}</p></div>',unsafe_allow_html=True)
cols[3].markdown(f'<div class="metric-card"><p class="metric-title">Unidades</p><p class="metric-value">{get_value("Unidades")}</p></div>', unsafe_allow_html=True)

# -------------------- Segunda linha de métricas --------------------
cols2 = st.columns(4)
cols2[0].markdown(f'<div class="metric-card"><p class="metric-title">Rentabilidade Viabilidade</p><p class="metric-value">{format_percent(get_value("Rentabilidade Viabilidade"))}</p></div>', unsafe_allow_html=True)
cols2[1].markdown(f'<div class="metric-card"><p class="metric-title">Rentabilidade Projetada</p><p class="metric-value">{format_percent(get_value("Rentabilidade Projetada"))}</p></div>', unsafe_allow_html=True)
cols2[2].markdown(f'<div class="metric-card"><p class="metric-title">Custo Área Construída</p><p class="metric-value">{format_money(get_value("Custo Área Construída"))}</p></div>', unsafe_allow_html=True)
cols2[3].markdown(f'<div class="metric-card"><p class="metric-title">Custo Área Privativa</p><p class="metric-value">{format_money(get_value("Custo Área Privativa"))}</p></div>', unsafe_allow_html=True)

# -------------------- Análise Financeira --------------------
st.markdown('<p class="sub-header">💰 Análise Financeira</p>', unsafe_allow_html=True)
cols1 = st.columns(3)
cols1[0].markdown(f'<div class="metric-card"><p class="metric-title">Orçamento Base</p><p class="metric-value">{format_money(get_value("Orçamento Base"))}</p></div>', unsafe_allow_html=True)
cols1[1].markdown(f'<div class="metric-card"><p class="metric-title">Orçamento Reajustado</p><p class="metric-value">{format_money(get_value("Orçamento Reajustado"))}</p></div>', unsafe_allow_html=True)
cols1[2].markdown(f'<div class="metric-card"><p class="metric-title">Custo Final</p><p class="metric-value">{format_money(get_value("Custo Final"))}</p></div>', unsafe_allow_html=True)
cols2 = st.columns(4)
cols2[0].markdown(f'<div class="metric-card"><p class="metric-title">Desvio</p><p class="metric-value">{format_money(get_value("Desvio"))}</p></div>', unsafe_allow_html=True)
cols2[1].markdown(f'<div class="metric-card"><p class="metric-title">Desembolso</p><p class="metric-value">{format_money(get_value("Desembolso"))}</p></div>', unsafe_allow_html=True)
cols2[2].markdown(f'<div class="metric-card"><p class="metric-title">Saldo</p><p class="metric-value">{format_money(get_value("Saldo"))}</p></div>', unsafe_allow_html=True)
cols2[3].markdown(f'<div class="metric-card"><p class="metric-title">Índice Econômico</p><p class="metric-value">{get_value("Índice Econômico")}</p></div>', unsafe_allow_html=True)

# -------------------- Avanço Físico --------------------
st.markdown('<p class="sub-header">📅 Avanço Físico</p>', unsafe_allow_html=True)
av_real_num = to_float(get_value("Avanço Físico Real", 0))
av_plan_num = to_float(get_value("Avanço Físico Planejado", 0))
aderencia_num = to_float(get_value("Aderência Física", 0))
if av_real_num <= 1: av_real_num *= 100
if av_plan_num <= 1: av_plan_num *= 100
if aderencia_num <= 1: aderencia_num *= 100
st.markdown(f"""
<div class="progress-wrapper">
    <div class="progress-bar" style="width: {av_real_num}%; background: #5DAAAB;">
        Real: {av_real_num:.1f}%
    </div>
</div>
<p style="color:#5DAAAB;font-weight:600;">Planejado: {av_plan_num:.1f}%</p>
<p style="color:#E5E8DD;font-weight:600;">Aderência: {aderencia_num:.1f}%</p>
""", unsafe_allow_html=True)

# -------------------- Linha do Tempo --------------------
st.markdown('<p class="sub-header">⏰ Linha do Tempo</p>', unsafe_allow_html=True)
meses_pt = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun", 7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}
def format_month_year_pt(date_val):
    try:
        dt = pd.to_datetime(date_val)
        return f"{meses_pt[dt.month]}/{dt.year}"
    except:
        return None
inicio = get_value("Início", None)
tend = get_value("Tendência", None)
prazo_concl = get_value("Prazo Conclusão", None)
prazo_cliente = get_value("Prazo Cliente", None)
cards = [
    {"label": "Início", "date": format_month_year_pt(inicio), "color": "#5DAAAB", "raw": inicio},
    {"label": "Tendência", "date": format_month_year_pt(tend), "color": "#5DAAAB", "raw": tend},
    {"label": "Prazo Conclusão", "date": format_month_year_pt(prazo_concl), "color": "#E5E8DD", "raw": prazo_concl},
    {"label": "Prazo Cliente", "date": format_month_year_pt(prazo_cliente), "color": "#E5E8DD", "raw": prazo_cliente}
]
cols = st.columns(len(cards))
for col, card in zip(cols, cards):
    col.markdown(f"""
        <div style="background-color:{card['color']}; padding: 15px; border-radius: 10px; text-align:center; color:#1A253C;">
            <p style="margin:0; font-weight:bold;">{card['label']}</p>
            <p style="margin:0;">{card['date'] if card['date'] else 'N/A'}</p>
        </div>
    """, unsafe_allow_html=True)

# Timeline Plotly
valid_cards = [c for c in cards if c['raw'] is not None]
if len(valid_cards) >= 2:
    dates = [pd.to_datetime(c['raw']) for c in valid_cards]
    labels = [c['label'] for c in valid_cards]
    colors = [c['color'] for c in valid_cards]
    fig_timeline = go.Figure()
    fig_timeline.add_trace(go.Scatter(x=[min(dates), max(dates)], y=[0,0], mode='lines', line=dict(color='gray', width=3), showlegend=False))
    for date, label, color in zip(dates, labels, colors):
        fig_timeline.add_trace(go.Scatter(x=[date], y=[0], mode='markers+text', marker=dict(size=15, color=color), text=[label], textposition='top center', name=label, textfont=dict(color='white', size=12)))
    fig_timeline.update_layout(title='Cronograma da Obra', showlegend=False, height=200, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', xaxis=dict(showgrid=False, zeroline=False, title=''), yaxis=dict(showgrid=False, zeroline=False, showticklabels=False), font=dict(color='white'))
    st.plotly_chart(fig_timeline, use_container_width=True)
else:
    st.info("Não há datas suficientes para criar a linha do tempo.")

# -------------------- Status Andamento da Obra --------------------
st.markdown('<p class="sub-header">📝 Status Andamento da Obra</p>', unsafe_allow_html=True)
status_rows = df_clean[df_clean['Metrica'].str.strip() == "Status Andamento Obra"]
if not status_rows.empty:
    status_list = status_rows['Valor'].tolist()
    with st.expander("📌 Ver Status Completo", expanded=False):
        for i, status in enumerate(status_list, 1):
            st.markdown(f"**{i}.** {status}")
else:
    st.info("Nenhum status de andamento disponível para esta obra.")
