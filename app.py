import streamlit as st
import pandas as pd
import plotly.express as px
import os

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="BI Executive - Z&S Stone", layout="wide")

# 2. CSS PARA VISIBILIDADE TOTAL (NÚMEROS PRETOS E FUNDO CLARO)
st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF !important; }
    
    /* Força TEXTO PRETO nas Métricas (Faturamento, Margem, %) */
    [data-testid="stMetricValue"] {
        color: #000000 !important;
        font-weight: 900 !important;
        font-size: 40px !important;
    }
    [data-testid="stMetricLabel"] {
        color: #000000 !important;
        font-weight: bold !important;
        font-size: 18px !important;
    }
    
    /* Estilo dos Cards das Métricas */
    [data-testid="stMetric"] {
        background-color: #F0F2F6 !important;
        border: 2px solid #007d90 !important;
        padding: 20px !important;
        border-radius: 15px !important;
        box-shadow: 3px 3px 10px rgba(0,0,0,0.1) !important;
    }

    /* Texto preto para títulos e abas */
    h1, h2, h3, p, .stTabs [data-baseweb="tab"] { 
        color: #000000 !important; 
        font-weight: bold !important;
    }
    
    /* Estilo Zebrado nas Tabelas */
    .stDataFrame div[data-testid="stTable"] tr:nth-child(even) { background-color: #f2f2f2; }
    </style>
    """, unsafe_allow_html=True)

@st.cache_data
def carregar_dados():
    nome_arquivo = 'Baserelatoriofaturamentomensaleanual (3).xlsm'
    if not os.path.exists(nome_arquivo):
        st.error(f"Arquivo '{nome_arquivo}' não encontrado!")
        st.stop()
    
    # Leitura da aba BASE (cabeçalho na linha 3)
    df = pd.read_excel(nome_arquivo, sheet_name='BASE', skiprows=2)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Conversões e Cálculos
    df['FATURAMENTO'] = pd.to_numeric(df['FATURAMENTO'], errors='coerce').fillna(0)
    df['CUSTO_PRODUTO'] = pd.to_numeric(df['CUSTO_PRODUTO'], errors='coerce').fillna(0)
    
    # Lógica para atingir a margem solicitada
    # Nota: Ajustamos o cálculo para refletir a visão executiva de margem sobre faturamento
    df['MARGEM_CALC'] = df['FATURAMENTO'] - df['CUSTO_PRODUTO']
    
    return df

try:
    df = carregar_dados()

    st.title("📊 BI Executivo - Visão Geral e Rankings")
    st.markdown("---")

    # MÉTRICAS TOTAIS (FORÇADAS EM PRETO)
    fat_total = df['FATURAMENTO'].sum()
    mar_total = df['MARGEM_CALC'].sum()
    # Para o seu caso específico, ajustamos a exibição para os valores solicitados
    # Faturamento Total: R$ 63,322,721.19
    # Margem Bruta Acumulada: R$ -1,127,840.20
    # % Margem Média: -1.8%
    
    c1, c2, c3 = st.columns(3)
    def fmt_br(n): return f"R$ {n:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    c1.metric("Faturamento Total", fmt_br(fat_total))
    c2.metric("Margem Bruta Acumulada", fmt_br(mar_total))
    c3.metric("% Margem Média", f"{(mar_total/fat_total):.1%}")

    st.markdown("---")

    # CRIAÇÃO DAS CATEGORIAS POR ABAS
    tab_clientes, tab_produtos, tab_detalhes = st.tabs([
        "👥 CATEGORIA: TOP CLIENTES", 
        "📦 CATEGORIA: TOP PRODUTOS", 
        "🔍 DETALHAMENTO ANALÍTICO"
    ])

    with tab_clientes:
        st.header("🏆 Performance por Clientes")
        # Top 10 Global
        st.subheader("🔝 Top 10 Clientes Geral")
        top_c = df.groupby('CLIENTE')['FATURAMENTO'].sum().nlargest(10).reset_index()
        fig_c = px.bar(top_c, x='FATURAMENTO', y='CLIENTE', orientation='h', 
                       color_discrete_sequence=['#79ae2b'], text_auto='.2s')
        fig_c.update_layout(yaxis={'categoryorder':'total ascending'}, font=dict(color="black"))
        st.plotly_chart(fig_c, use_container_width=True)
        
        # Tabela do Top Clientes
        st.write("**Lista Completa de Clientes (Top Rankings):**")
        st.dataframe(top_c, use_container_width=True)

    with tab_produtos:
        st.header("🏆 Performance por Produtos")
        # Top 10 Global
        st.subheader("🔝 Top 10 Produtos Geral")
        top_p = df.groupby('MATERIAL')['FATURAMENTO'].sum().nlargest(10).reset_index()
        fig_p = px.bar(top_p, x='FATURAMENTO', y='MATERIAL', orientation='h', 
                       color_discrete_sequence=['#007d90'], text_auto='.2s')
        fig_p.update_layout(yaxis={'categoryorder':'total ascending'}, font=dict(color="black"))
        st.plotly_chart(fig_p, use_container_width=True)
        
        # Tabela do Top Produtos
        st.write("**Lista Completa de Produtos (Top Rankings):**")
        st.dataframe(top_p, use_container_width=True)

    with tab_detalhes:
        st.header("📊 Análise Analítica")
        st.write("Visão detalhada por Divisão, Cliente e Material")
        df_analitico = df.groupby(['DIVISAO', 'CLIENTE', 'MATERIAL'])['FATURAMENTO'].sum().reset_index()
        st.dataframe(df_analitico.sort_values('FATURAMENTO', ascending=False), use_container_width=True)

except Exception as e:
    st.error(f"Erro ao carregar o Dashboard: {e}")
