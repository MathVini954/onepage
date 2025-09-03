import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
import os

# -------------------- Configuração da página --------------------
st.set_page_config(
    page_title="Dashboard de Obras",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------- Estilos CSS --------------------
st.markdown("""
<style>
.main-header {
    font-size: 2.8rem;
    color: #1E3A8A;
    font-weight: 800;
    margin-bottom: 2rem;
    padding-bottom: 0.5rem;
    border-bottom: 3px solid #3B82F6;
    text-align: center;
}
.sub-header {
    font-size: 1.8rem;
    color: #374151;
    font-weight: 700;
    margin: 2rem 0 1.5rem 0;
    padding-bottom: 0.5rem;
    border-bottom: 2px solid #E5E7EB;
}
.metric-card {
    background-color: #1E293B;
    border-radius: 0.75rem;
    padding: 1rem;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
    height: 100%;
    border-left: 5px solid #3B82F6;
    margin-bottom: 1rem;
}
.metric-title {
    font-size: 1rem;
    color: #93C5FD;
    font-weight: 600;
    margin-bottom: 0.5rem;
}
.metric-value {
    font-size: 1.6rem;
    color: #FFFFFF;
    font-weight: 800;
}
.section-container {
    background-color: #0F172A;
    border-radius: 1rem;
    padding: 1.5rem;
    margin-bottom: 2rem;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.4);
}
.progress-wrapper {
    background-color: #1E293B;
    border-radius: 20px;
    padding: 5px;
    width: 100%;
}
.progress-bar {
    height: 30px;
    border-radius: 20px;
    text-align: center;
    font-weight: bold;
    color: white;
    line-height: 30px;
}
</style>
""", unsafe_allow_html=True)

# -------------------- Sidebar (filtro + logo empresa) --------------------
logo_empresa_path = "empresa_logo.png"
if os.path.exists(logo_empresa_path):
    st.sidebar.image(logo_empresa_path, width=200)

st.sidebar.markdown("### 📂 Selecione a Obra")

file_path = "ONE_PAGE.xlsx"
if not os.path.exists(file_path):
    st.error("❌ Arquivo ONE_PAGE.xlsx não encontrado no diretório!")
    st.stop()

excel_file = pd.ExcelFile(file_path)
sheet_names = excel_file.sheet_names
selected_sheet = st.sidebar.selectbox("Obra:", sheet_names)

# -------------------- Cabeçalho principal --------------------
st.markdown('<p class="main-header">🏗️ Dashboard de Obras</p>', unsafe_allow_html=True)

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

# -------------------- Métricas principais --------------------
st.markdown('<p class="sub-header">📊 Métricas Principais</p>', unsafe_allow_html=True)

# Primeira linha
cols = st.columns(4)
cols[0].markdown(f'<div class="metric-card"><p class="metric-title">AC(m²)</p><p class="metric-value">{get_value("AC(m²)")}</p></div>', unsafe_allow_html=True)
cols[1].markdown(f'<div class="metric-card"><p class="metric-title">AP(m²)</p><p class="metric-value">{get_value("AP(m²)")}</p></div>', unsafe_allow_html=True)
cols[2].markdown(f'<div class="metric-card"><p class="metric-title">Ef</p><p class="metric-value">{format_percent(get_value("Ef"))}</p></div>', unsafe_allow_html=True)
cols[3].markdown(f'<div class="metric-card"><p class="metric-title">Total Unidades</p><p class="metric-value">{get_value("Total Unidades")}</p></div>', unsafe_allow_html=True)

# Segunda linha
cols2 = st.columns(4)
cols2[0].markdown(f'<div class="metric-card"><p class="metric-title">Rentab. Viabilidade</p><p class="metric-value">{format_percent(get_value("Rentab. Viabilidade"))}</p></div>', unsafe_allow_html=True)
cols2[1].markdown(f'<div class="metric-card"><p class="metric-title">Rentab. Projetada</p><p class="metric-value">{format_percent(get_value("Rentab. Projetada"))}</p></div>', unsafe_allow_html=True)
cols2[2].markdown(f'<div class="metric-card"><p class="metric-title">Custo Atual AC</p><p class="metric-value">{format_money(get_value("Custo Atual AC"))}</p></div>', unsafe_allow_html=True)
cols2[3].markdown(f'<div class="metric-card"><p class="metric-title">Custo Atual AP</p><p class="metric-value">{format_money(get_value("Custo Atual AP"))}</p></div>', unsafe_allow_html=True)

# -------------------- Análise Financeira --------------------
st.markdown('<p class="sub-header">💰 Análise Financeira</p>', unsafe_allow_html=True)

st.markdown('<div class="section-container">', unsafe_allow_html=True)
cols_fin = st.columns(3)
cols_fin[0].markdown(f'<div class="metric-card"><p class="metric-title">Orçamento Base</p><p class="metric-value">{format_money(get_value("Orçamento Base"))}</p></div>', unsafe_allow_html=True)
cols_fin[1].markdown(f'<div class="metric-card"><p class="metric-title">Orçamento Reajustado</p><p class="metric-value">{format_money(get_value("Orçamento Reajustado"))}</p></div>', unsafe_allow_html=True)
cols_fin[2].markdown(f'<div class="metric-card"><p class="metric-title">Custo Final</p><p class="metric-value">{format_money(get_value("Custo Final"))}</p></div>', unsafe_allow_html=True)

cols_fin2 = st.columns(4)
cols_fin2[0].markdown(f'<div class="metric-card"><p class="metric-title">Desvio</p><p class="metric-value">{get_value("Desvio")}</p></div>', unsafe_allow_html=True)
cols_fin2[1].markdown(f'<div class="metric-card"><p class="metric-title">Desembolso</p><p class="metric-value">{format_money(get_value("Desembolso"))}</p></div>', unsafe_allow_html=True)
cols_fin2[2].markdown(f'<div class="metric-card"><p class="metric-title">Saldo</p><p class="metric-value">{format_money(get_value("Saldo"))}</p></div>', unsafe_allow_html=True)
cols_fin2[3].markdown(f'<div class="metric-card"><p class="metric-title">Índice Econômico</p><p class="metric-value">{get_value("Índice Econômico")}</p></div>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# -------------------- Barra de progresso (Avanço Físico) --------------------
st.markdown('<p class="sub-header">📅 Avanço Físico</p>', unsafe_allow_html=True)

av_real_num = to_float(get_value("Avanço Físico Real", 0))
av_plan_num = to_float(get_value("Avanço Físico Planejado", 0))
aderencia_num = to_float(get_value("Aderência Física", 0))

if av_real_num <= 1: av_real_num *= 100
if av_plan_num <= 1: av_plan_num *= 100
if aderencia_num <= 1: aderencia_num *= 100

st.markdown(f"""
<div class="progress-wrapper">
    <div class="progress-bar" style="width: {av_real_num}%; background: #3B82F6;">
        Real: {av_real_num:.1f}%
    </div>
</div>
<p style="color:#EF4444;font-weight:600;">Planejado: {av_plan_num:.1f}%</p>
<p style="color:#10B981;font-weight:600;">Aderência: {aderencia_num:.1f}%</p>
""", unsafe_allow_html=True)

# -------------------- Timeline --------------------
st.markdown('<p class="sub-header">⏰ Linha do Tempo</p>', unsafe_allow_html=True)
inicio = get_value("Início", "N/A")
tend = get_value("Tend", "N/A")
prazo_concl = get_value("Prazo Concl.", "N/A")
prazo_cliente = get_value("Prazo Cliente", "N/A")

dates = [inicio, tend, prazo_concl, prazo_cliente]
labels = ["Início", "Tendência", "Prazo Conclusão", "Prazo Cliente"]
colors = ["#3B82F6", "#F59E0B", "#10B981", "#EF4444"]

date_values = []
for d in dates:
    if isinstance(d, (datetime, pd.Timestamp)):
        date_values.append(d)
    elif isinstance(d, str) and d != "N/A":
        try:
            date_values.append(pd.to_datetime(d))
        except:
            date_values.append(None)
    else:
        date_values.append(None)

valid_dates = [d for d in date_values if d is not None]
if len(valid_dates) >= 2:
    min_date = min(valid_dates)
    max_date = max(valid_dates)
    fig_timeline = go.Figure()
    fig_timeline.add_trace(go.Scatter(
        x=[min_date, max_date],
        y=[0,0],
        mode='lines',
        line=dict(color='white', width=3),
        showlegend=False
    ))
    for i, (date, label, color) in enumerate(zip(date_values, labels, colors)):
        if date is not None:
            fig_timeline.add_trace(go.Scatter(
                x=[date],
                y=[0],
                mode='markers+text',
                marker=dict(size=15, color=color),
                text=[label],
                textposition="top center",
                name=label,
                textfont=dict(color='white', size=12)
            ))
    fig_timeline.update_layout(
        title='Cronograma da Obra',
        showlegend=True,
        height=300,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        xaxis=dict(showgrid=False, zeroline=False),
        yaxis=dict(showgrid=False, zeroline=False, showticklabels=False)
    )
    st.plotly_chart(fig_timeline, use_container_width=True)
else:
    st.info("Não há datas suficientes para criar a linha do tempo.")

# -------------------- Visualizar dados --------------------
with st.expander("🔍 Visualizar dados carregados"):
    st.dataframe(df_clean, use_container_width=True)

# -------------------- Footer --------------------
st.markdown("---")
st.markdown("<div style='text-align: center; color: #6B7280;'>Dashboard atualizado em tempo real | Dados da obra selecionada</div>", unsafe_allow_html=True)
