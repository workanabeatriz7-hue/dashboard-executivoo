import streamlit as st
import pandas as pd
import plotly.express as px

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="BI Executive - Stone", layout="wide")

# CSS para garantir que os números fiquem pretos e visíveis, com o zebrado
st.markdown("""
    <style>
    /* Força o texto das métricas a ser visível (Preto) */
    [data-testid="stMetricValue"] {
        color: #333333 !important;
        font-weight: bold;
    }
    [data-testid="stMetricLabel"] {
        color: #555555 !important;
    }
    /* Estilo dos Cards */
    [data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #007d90;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 2px 2px 8px rgba(0,0,0,0.1);
    }
    /* Linhas zebradas na tabela */
    .stDataFrame div[data-testid="stTable"] tr:nth-child(even) {
        background-color: #f2f2f2;
    }
    </style>
    """, unsafe_allow_html=True)

COR_AZUL = '#007d90'
COR_VERDE = '#79ae2b'

@st.cache_data
def carregar_dados():
    caminho = 'Baserelatoriofaturamentomensaleanual (3).xlsm'
    # Pula 2 linhas para pegar o cabeçalho correto
    df = pd.read_excel(caminho, sheet_name='BASE', skiprows=2)
    
    # Limpa nomes de colunas
    df.columns = [str(c).strip() for c in df.columns]
    
    # Ajuste de Datas
    df['DATA_EMISSAO_NF'] = pd.to_datetime(df['DATA_EMISSAO_NF'])
    df['Mês/Ano'] = df['DATA_EMISSAO_NF'].dt.strftime('%m/%Y')
    
    # Garantir que valores sejam numéricos
    df['FATURAMENTO'] = pd.to_numeric(df['FATURAMENTO'], errors='coerce').fillna(0)
    # Na sua planilha, a coluna de Margem tem um espaço ou nome específico
    # Vamos usar 'MARGEM' ou 'MARGEM ' (limpamos os espaços acima com strip)
    df['MARGEM'] = pd.to_numeric(df['MARGEM'], errors='coerce').fillna(0)
    
    return df

try:
    df = carregar_dados()

    # --- FILTROS ---
    st.sidebar.title("Configurações")
    todas_div = sorted(df['DIVISAO'].unique())
    selecao = st.sidebar.multiselect("Divisões Ativas", options=todas_div, default=todas_div)
    
    df_f = df.query("DIVISAO == @selecao")

    # --- TÍTULO ---
    st.title("📊 Análise de Margem e Performance Comercial")
    st.markdown("### Geral Mensal e Anual")

    # --- KPIs (Onde os números não apareciam) ---
    fat_total = df_f['FATURAMENTO'].sum()
    mar_total = df_f['MARGEM'].sum()
    margem_perc = (mar_total / fat_total) if fat_total != 0 else 0

    col1, col2, col3 = st.columns(3)
    # Usamos o formato R$ com separador de milhar brasileiro
    col1.metric("FATURAMENTO TOTAL", f"R$ {fat_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
    col2.metric("MARGEM BRUTA ACUM.", f"R$ {mar_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
    col3.metric("% MARGEM MÉDIA", f"{margem_perc:.1%}")

    st.markdown("---")

    # --- ABAS ANALÍTICAS ---
    t_margem, t_client, t_prod, t_origem = st.tabs([
        "📈 Evolução Margem", "👤 Analytic Client", "📦 Analytic Product", "🌍 Origem/Export"
    ])

    with t_margem:
        resumo_mes = df_f.groupby('Mês/Ano').agg({'FATURAMENTO':'sum', 'MARGEM':'sum'}).reset_index()
        fig = px.bar(resumo_mes, x='Mês/Ano', y='FATURAMENTO', 
                     title="Faturamento por Mês", color_discrete_sequence=[COR_AZUL])
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(resumo_mes, use_container_width=True)

    with t_client:
        st.subheader("Análise Analítica de Clientes")
        client_data = df_f.groupby(['CLIENTE', 'DIVISAO']).agg({'FATURAMENTO':'sum', 'MARGEM':'sum'}).reset_index()
        st.dataframe(client_data.sort_values('FATURAMENTO', ascending=False), use_container_width=True)

    with t_prod:
        st.subheader("Análise Analítica de Produtos")
        prod_data = df_f.groupby(['MATERIAL', 'DIVISAO']).agg({'FATURAMENTO':'sum', 'MARGEM':'sum'}).reset_index()
        st.dataframe(prod_data.sort_values('FATURAMENTO', ascending=False), use_container_width=True)

    with t_origem:
        c_o, c_e = st.columns(2)
        with c_o:
            st.write("**Origem da Operação**")
            origem = df_f.groupby('OPERACAO')['FATURAMENTO'].sum().reset_index()
            st.plotly_chart(px.pie(origem, values='FATURAMENTO', names='OPERACAO', color_discrete_sequence=[COR_AZUL, COR_VERDE]), use_container_width=True)
        with c_e:
            st.write("**Exportação**")
            # Verifica a coluna EX da sua base
            exp = df_f.groupby('EX')['FATURAMENTO'].sum().reset_index()
            st.dataframe(exp, use_container_width=True)

except Exception as e:
    st.error(f"Erro ao carregar dashboard: {e}")
