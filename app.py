import streamlit as st
import pandas as pd
import plotly.express as px

# 1. CONFIGURAÇÃO DA PÁGINA WEB
st.set_page_config(page_title="Dashboard Executivo Stone", layout="wide")

# Cores Oficiais
COR_AZUL = '#007d90'
COR_VERDE = '#79ae2b'

# 2. FUNÇÃO DE CARREGAMENTO DE DADOS (VERSÃO NUVEM)
@st.cache_data
def carregar_dados():
    # Na nuvem, o arquivo deve estar na raiz do seu GitHub
    caminho = 'Baserelatoriofaturamentomensaleanual (3).xlsm'
    
    # Lê a aba BASE pulando as 2 linhas de cabeçalho do sistema
    df = pd.read_excel(caminho, sheet_name='BASE', skiprows=2)
    
    # Limpa nomes de colunas (remove espaços e garante nomes limpos)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Mapeamento de colunas e cálculos
    df['DATA_EMISSAO_NF'] = pd.to_datetime(df['DATA_EMISSAO_NF'])
    df['Lucro'] = df['FATURAMENTO'] - df['CUSTO_PRODUTO']
    df['Mês_Ano'] = df['DATA_EMISSAO_NF'].dt.strftime('%m/%Y')
    return df

# 3. EXECUÇÃO DO DASHBOARD
try:
    df = carregar_dados()

    # --- MENU LATERAL (FILTROS) ---
    st.sidebar.header("Filtros Globais")
    divisoes = st.sidebar.multiselect(
        "Selecione as Divisões:", 
        options=sorted(df['DIVISAO'].unique()), 
        default=df['DIVISAO'].unique()
    )
    
    # Aplicando o filtro aos dados
    df_f = df.query("DIVISAO == @divisoes")

    # --- CRIAÇÃO DAS ABAS (TABS) ---
    tab1, tab2 = st.tabs(["📊 Sumário Executivo", "🏆 Top 10 Performance"])

    # --- ABA 1: SUMÁRIO EXECUTIVO ---
    with tab1:
        st.header("Sumário Executivo de Resultados")
        
        # Cálculos de KPI
        fat_total = df_f['FATURAMENTO'].sum()
        lucro_total = df_f['Lucro'].sum()
        margem_med = (lucro_total / fat_total) if fat_total != 0 else 0
        
        # Exibição dos Cartões (KPIs)
        k1, k2, k3 = st.columns(3)
        k1.metric("FATURAMENTO TOTAL", f"R$ {fat_total:,.2f}")
        k2.metric("LUCRO OPERACIONAL", f"R$ {lucro_total:,.2f}")
        k3.metric("MARGEM MÉDIA", f"{margem_med:.1%}")

        st.markdown("---")
        
        # Gráfico de Evolução Mensal
        st.subheader("Tendência Mensal de Faturamento")
        # Agrupando por mês para o gráfico
        df_mes = df_f.groupby(df_f['DATA_EMISSAO_NF'].dt.to_period('M'))['FATURAMENTO'].sum().reset_index()
        df_mes['DATA_EMISSAO_NF'] = df_mes['DATA_EMISSAO_NF'].dt.to_timestamp()
        
        fig_evol = px.area(
            df_mes, x='DATA_EMISSAO_NF', y='FATURAMENTO',
            labels={'FATURAMENTO': 'Faturamento (R$)', 'DATA_EMISSAO_NF': 'Mês'},
            color_discrete_sequence=[COR_AZUL], template="plotly_white"
        )
        st.plotly_chart(fig_evol, use_container_width=True)
        
        # Explicação rápida dos números
        with st.expander("O que esses números indicam?"):
            st.write("**Faturamento:** Volume total de vendas brutas.")
            st.write("**Lucro:** Valor que sobra após descontar o custo do produto.")
            st.write("**Margem:** Percentual de rentabilidade sobre a venda.")

    # --- ABA 2: TOP 10 PERFORMANCE ---
    with tab2:
        st.header("Ranking de Performance (Top 10)")
        
        c1, c2 = st.columns(2)
        
        with c1:
            st.subheader("Top 10 Clientes (R$)")
            top_c = df_f.groupby('CLIENTE')['FATURAMENTO'].sum().nlargest(10).reset_index()
            fig_c = px.bar(
                top_c, x='FATURAMENTO', y='CLIENTE', 
                orientation='h', color_discrete_sequence=[COR_VERDE],
                template="plotly_white"
            )
            fig_c.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_c, use_container_width=True)
            
        with c2:
            st.subheader("Top 10 Produtos (R$)")
            top_p = df_f.groupby('MATERIAL')['FATURAMENTO'].sum().nlargest(10).reset_index()
            fig_p = px.bar(
                top_p, x='FATURAMENTO', y='MATERIAL', 
                orientation='h', color_discrete_sequence=[COR_AZUL],
                template="plotly_white"
            )
            fig_p.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_p, use_container_width=True)

except Exception as e:
    st.error(f"Ocorreu um erro ao carregar o Dashboard: {e}")
    st.info("Verifique se o arquivo Excel está com o nome correto no seu GitHub.")
