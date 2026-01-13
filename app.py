import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="BI Executive - Z&S Stone", layout="wide")

# PALETA DE CORES
COR_AZUL = '#007d90'
COR_VERDE = '#79ae2b'
COR_CINZA = '#f2f2f2'

# CSS PARA ESTILIZAÇÃO: Zebrado, Cores das Métricas e Design dos Cards
st.markdown(f"""
    <style>
    /* Zebrado nas tabelas */
    .stDataFrame div[data-testid="stTable"] tr:nth-child(even) {{ background-color: {COR_CINZA}; }}
    
    /* Customização dos Cards de KPI */
    [data-testid="stMetricValue"] {{ color: {COR_AZUL} !important; font-weight: bold; }}
    [data-testid="stMetricLabel"] {{ color: #333333 !important; font-size: 16px; font-weight: bold; }}
    [data-testid="stMetric"] {{
        background-color: white;
        border-left: 6px solid {COR_AZUL};
        padding: 15px;
        border-radius: 8px;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
    }}
    .main {{ background-color: #fdfdfd; }}
    </style>
    """, unsafe_allow_html=True)

@st.cache_data
def carregar_dados():
    caminho = 'Baserelatoriofaturamentomensaleanual (3).xlsm'
    # Leitura da aba BASE conforme a estrutura original (linha 3 é o cabeçalho)
    df = pd.read_excel(caminho, sheet_name='BASE', skiprows=2)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Tratamento de Colunas e Cálculos
    df['DATA_EMISSAO_NF'] = pd.to_datetime(df['DATA_EMISSAO_NF'])
    df['Mês/Ano'] = df['DATA_EMISSAO_NF'].dt.strftime('%m/%Y')
    df['Ano'] = df['DATA_EMISSAO_NF'].dt.year
    
    # Garantindo colunas numéricas (Atenção ao espaço no nome da coluna 'MARGEM ')
    df['FATURAMENTO'] = pd.to_numeric(df['FATURAMENTO'], errors='coerce').fillna(0)
    df['CUSTO_PRODUTO'] = pd.to_numeric(df['CUSTO_PRODUTO'], errors='coerce').fillna(0)
    df['MARGEM_VALOR'] = pd.to_numeric(df['MARGEM '], errors='coerce').fillna(0)
    df['MARGEM_PERC'] = (df['MARGEM_VALOR'] / df['FATURAMENTO']).fillna(0)
    
    return df

try:
    df = carregar_dados()

    # --- SIDEBAR (FILTROS) ---
    st.sidebar.title("📊 BI Executive Z&S")
    st.sidebar.markdown("---")
    
    # Filtro de Divisão (Geral por padrão)
    divisoes = sorted(df['DIVISAO'].unique())
    selecao_div = st.sidebar.multiselect("Filtrar por Divisão:", options=divisoes, default=divisoes)
    
    # Filtro de Ano
    anos = sorted(df['Ano'].unique())
    selecao_ano = st.sidebar.multiselect("Filtrar Ano:", options=anos, default=anos)

    # Aplicação dos Filtros
    df_f = df.query("DIVISAO == @selecao_div and Ano == @selecao_ano")

    # --- TÍTULO ---
    st.title("🛡️ Inteligência Comercial: Geral Mensal e Anual")
    
    # --- SUMÁRIO EXECUTIVO (MENSAL E ANUAL) ---
    col_fat, col_mar, col_perc, col_qtde = st.columns(4)
    tot_fat = df_f['FATURAMENTO'].sum()
    tot_mar = df_f['MARGEM_VALOR'].sum()
    
    col_fat.metric("FATURAMENTO BRUTO", f"R$ {tot_fat:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
    col_mar.metric("MARGEM BRUTA", f"R$ {tot_mar:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
    col_perc.metric("% MARGEM MÉDIA", f"{(tot_mar/tot_fat if tot_fat != 0 else 0):.1%}")
    col_qtde.metric("QTDE TOTAL VENDIDA", f"{df_f['QTDE'].sum():,.0f}")

    st.markdown("---")

    # --- AS 10 ABAS DE ANÁLISE (Respeitando as Abas do Excel) ---
    tab_margem, tab_client, tab_prod, tab_origem, tab_top = st.tabs([
        "📉 MARGEM (M/A)", "👥 ANALYTICAL/SYNTHETIC CLIENT", "📦 ANALYTICAL/SYNTHETIC PROD", "🌍 ORIGEM / EXPORT", "🏆 TOP 10 PERFORMANCE"
    ])

    # ABA 1: MARGEM
    with tab_margem:
        st.subheader("Análise Evolutiva de Margens")
        resumo_margem = df_f.groupby('Mês/Ano').agg({'FATURAMENTO':'sum', 'MARGEM_VALOR':'sum'}).reset_index()
        
        fig_evol = go.Figure()
        fig_evol.add_trace(go.Scatter(x=resumo_margem['Mês/Ano'], y=resumo_margem['FATURAMENTO'], name='Faturamento', line=dict(color=COR_AZUL, width=4)))
        fig_evol.add_trace(go.Bar(x=resumo_margem['Mês/Ano'], y=resumo_margem['MARGEM_VALOR'], name='Margem Bruta', marker_color=COR_VERDE, opacity=0.6))
        fig_evol.update_layout(title="Evolução Faturamento vs Margem", template="plotly_white")
        st.plotly_chart(fig_evol, use_container_width=True)
        
        st.write("**Detalhamento de Margens (Zebrado):**")
        st.dataframe(resumo_margem.style.format({'FATURAMENTO': 'R$ {:,.2f}', 'MARGEM_VALOR': 'R$ {:,.2f}'}), use_container_width=True)

    # ABA 2: ANALYTICAL/SYNTHETIC CLIENT
    with tab_client:
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.subheader("Analytical Client (Vendas por Divisão)")
            an_client = df_f.groupby(['CLIENTE', 'DIVISAO']).agg({'FATURAMENTO':'sum', 'MARGEM_VALOR':'sum'}).reset_index().sort_values('FATURAMENTO', ascending=False)
            st.dataframe(an_client, use_container_width=True)
        with col_c2:
            st.subheader("Synthetic Client (Total por Cliente)")
            syn_client = df_f.groupby('CLIENTE').agg({'FATURAMENTO':'sum', 'MARGEM_VALOR':'sum', 'MARGEM_PERC':'mean'}).reset_index().sort_values('FATURAMENTO', ascending=False)
            st.dataframe(syn_client, use_container_width=True)

    # ABA 3: ANALYTICAL/SYNTHETIC PRODUCT
    with tab_prod:
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.subheader("Analytical Product")
            an_prod = df_f.groupby(['MATERIAL', 'DIVISAO']).agg({'FATURAMENTO':'sum', 'QTDE':'sum'}).reset_index().sort_values('FATURAMENTO', ascending=False)
            st.dataframe(an_prod, use_container_width=True)
        with col_p2:
            st.subheader("Synthetic Product")
            syn_prod = df_f.groupby('MATERIAL').agg({'FATURAMENTO':'sum', 'MARGEM_VALOR':'sum', 'MARGEM_PERC':'mean'}).reset_index().sort_values('FATURAMENTO', ascending=False)
            st.dataframe(syn_prod, use_container_width=True)

    # ABA 4: ORIGEM / EXPORT
    with tab_origem:
        col_o, col_e = st.columns(2)
        with col_o:
            st.subheader("Analytical Origin")
            origem = df_f.groupby('OPERACAO')['FATURAMENTO'].sum().reset_index()
            fig_pie = px.pie(origem, values='FATURAMENTO', names='OPERACAO', hole=0.4, color_discrete_sequence=[COR_AZUL, COR_VERDE, '#333333'])
            st.plotly_chart(fig_pie, use_container_width=True)
        with col_e:
            st.subheader("Exportação (EX)")
            # Valida se a coluna EX existe para análise de export
            df_f['EX'] = df_f['EX'].fillna('Não')
            export = df_f.groupby('EX').agg({'FATURAMENTO':'sum', 'QTDE':'sum'}).reset_index()
            st.dataframe(export, use_container_width=True)

    # ABA 5: TOP 10 PERFORMANCE
    with tab_top:
        st.subheader("Ranking de Performance - Top 10 Geral")
        col_t1, col_t2 = st.columns(2)
        
        with col_t1:
            top10_c = df_f.groupby('CLIENTE')['FATURAMENTO'].sum().nlargest(10).reset_index()
            fig_top_c = px.bar(top10_c, x='FATURAMENTO', y='CLIENTE', orientation='h', title="Top 10 Clientes", color_discrete_sequence=[COR_VERDE])
            fig_top_c.update_layout(yaxis={'categoryorder':'total ascending'}, template="plotly_white")
            st.plotly_chart(fig_top_c, use_container_width=True)
            
        with col_t2:
            top10_p = df_f.groupby('MATERIAL')['FATURAMENTO'].sum().nlargest(10).reset_index()
            fig_top_p = px.bar(top10_p, x='FATURAMENTO', y='MATERIAL', orientation='h', title="Top 10 Produtos", color_discrete_sequence=[COR_AZUL])
            fig_top_p.update_layout(yaxis={'categoryorder':'total ascending'}, template="plotly_white")
            st.plotly_chart(fig_top_p, use_container_width=True)

except Exception as e:
    st.error(f"Erro Crítico ao carregar o Modelo BI: {e}")
