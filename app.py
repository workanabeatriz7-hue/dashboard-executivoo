import streamlit as st
import pandas as pd
import plotly.express as px
import os

# CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Dashboard Executivo Stone", layout="wide")

# CORES OFICIAIS
COR_AZUL = '#007d90'
COR_VERDE = '#79ae2b'

@st.cache_data
def carregar_dados():
    desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    caminho = os.path.join(desktop, 'Baserelatoriofaturamentomensaleanual (3).xlsm')
    # Lê a aba BASE pulando 2 linhas (conforme sucesso anterior)
    df = pd.read_excel(caminho, sheet_name='BASE', skiprows=2)
    df.columns = [str(c).strip() for c in df.columns]
    
    df['DATA_EMISSAO_NF'] = pd.to_datetime(df['DATA_EMISSAO_NF'])
    df['Lucro'] = df['FATURAMENTO'] - df['CUSTO_PRODUTO']
    df['Mês'] = df['DATA_EMISSAO_NF'].dt.strftime('%m/%Y')
    return df

try:
    df = carregar_dados()

    # --- MENU LATERAL (FILTROS) ---
    st.sidebar.header("Filtros de Análise")
    divisoes = st.sidebar.multiselect("Selecione a Divisão", options=df['DIVISAO'].unique(), default=df['DIVISAO'].unique())
    df_filtrado = df.query("DIVISAO == @divisoes")

    # --- SEÇÃO 1: SUMÁRIO EXECUTIVO ---
    st.title("📊 Sumário Executivo de Resultados")
    
    # KPIs principais em caixas destacadas
    total_fat = df_filtrado['FATURAMENTO'].sum()
    total_lucro = df_filtrado['Lucro'].sum()
    margem_med = (total_lucro / total_fat) if total_fat != 0 else 0

    kpi1, kpi2, kpi3 = st.columns(3)
    with kpi1:
        st.metric("FATURAMENTO TOTAL", f"R$ {total_fat:,.2f}")
    with kpi2:
        st.metric("LUCRO OPERACIONAL", f"R$ {total_lucro:,.2f}")
    with kpi3:
        st.metric("MARGEM MÉDIA", f"{margem_med:.1%}")

    # Gráfico de Evolução de Vendas
    st.markdown("### Evolução Mensal de Vendas")
    vendas_mes = df_filtrado.groupby(df_filtrado['DATA_EMISSAO_NF'].dt.to_period('M')).agg({'FATURAMENTO': 'sum'}).reset_index()
    vendas_mes['DATA_EMISSAO_NF'] = vendas_mes['DATA_EMISSAO_NF'].dt.to_timestamp()
    
    fig_evol = px.area(vendas_mes, x='DATA_EMISSAO_NF', y='FATURAMENTO', 
                       color_discrete_sequence=[COR_AZUL], template="plotly_white")
    st.plotly_chart(fig_evol, use_container_width=True)

    st.markdown("---")

    # --- SEÇÃO 2: TOP 10 CLIENTES E PRODUTOS ---
    st.title("🏆 Ranking de Performance (Top 10)")
    
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Top 10 Clientes")
        top_c = df_filtrado.groupby('CLIENTE')['FATURAMENTO'].sum().nlargest(10).reset_index()
        fig_c = px.bar(top_c, x='FATURAMENTO', y='CLIENTE', orientation='h', 
                       color_discrete_sequence=[COR_VERDE], template="plotly_white")
        fig_c.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_c, use_container_width=True)

    with col2:
        st.subheader("Top 10 Produtos")
        top_p = df_filtrado.groupby('MATERIAL')['FATURAMENTO'].sum().nlargest(10).reset_index()
        fig_p = px.bar(top_p, x='FATURAMENTO', y='MATERIAL', orientation='h', 
                       color_discrete_sequence=[COR_AZUL], template="plotly_white")
        fig_p.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_p, use_container_width=True)

except Exception as e:
    st.error(f"Erro: {e}")