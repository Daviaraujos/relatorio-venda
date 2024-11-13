import pandas as pd
import plotly.express as px
import streamlit as st
import requests
from streamlit_autorefresh import st_autorefresh

# Configuração da página
st.set_page_config(page_title="Análise de Vendas", page_icon=":bar_chart:", layout="wide")

# Atualização automática
st_autorefresh(interval=5000, key="data_refresh")

# Título do dashboard
st.title('Análise de Vendas')

# Carregar dados da planilha Excel
try:
    response = requests.get('https://docs.google.com/spreadsheets/d/e/2PACX-1vQt8EOEnxeGbcvhHIz_5ubSFJk9G8ids7B-xW8OpsViI3rQVhMdtKFuXl_Lmrnb8h0jWnaoL0cQK2rR/pub?output=xlsx')
    response.raise_for_status()
    xls = pd.ExcelFile(response.content)
    df = pd.read_excel(xls, sheet_name='Copia de DADOS GERAIS COMERCIAL')

    # Processamento de datas e limpeza
    df['Data da assinatura'] = pd.to_datetime(df['Data da assinatura'], errors='coerce')
    df['Mês'] = df['Data da assinatura'].dt.to_period('M').astype(str)
    meses_disponiveis = ["Todos os Períodos"] + sorted(df['Mês'].unique(), reverse=True)

    # Barra lateral para selecionar o mês
    mes_selecionado = st.sidebar.selectbox("Mês", meses_disponiveis)

    # Filtrar dados com base na seleção
    if mes_selecionado == "Todos os Períodos":
        df_filtrado = df
    else:
        df_filtrado = df[df['Mês'] == mes_selecionado]

    # Exibir principais métricas
    total_leads = df_filtrado['Number_leads'].count()
    leads_qualificados = df_filtrado[df_filtrado['Atende aos requisitos'] != '-']['Number_leads'].count()
    leads_respondidos = df_filtrado[df_filtrado['Respondeu as msgns'] != '-']['Number_leads'].count()
    propostas_aceitas = df_filtrado[df_filtrado['Aceitou'] != '-']['Number_leads'].count()
    assinaturas_finalizadas = df_filtrado['Data da assinatura'].count()

    taxa_conversao = (assinaturas_finalizadas / total_leads) * 100 if total_leads else 0
    taxa_resposta = (leads_respondidos / total_leads) * 100 if total_leads else 0

    # Colunas para métricas
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total de Leads", total_leads)
    col2.metric("Leads Qualificados", leads_qualificados)
    col3.metric("Leads Respondidos", leads_respondidos)
    col4.metric("Propostas Aceitas", propostas_aceitas)
    col5.metric("Assinaturas Finalizadas", assinaturas_finalizadas)

    # Gráficos
    col6, col7, col8 = st.columns(3)

    # Gráfico de vendas por consultor
    vendas_por_consultor = df_filtrado[df_filtrado['Aceitou'] == 'SIM'].groupby('Consultor')['Number_leads'].count().reset_index()
    fig_consultor = px.bar(vendas_por_consultor, x='Consultor', y='Number_leads', title="Vendas por Consultor", color='Consultor')
    col6.plotly_chart(fig_consultor, use_container_width=True)

    # Funil de vendas
    funil_dados = {
        'Etapas': ['Leads', 'Leads Qualificados', 'Leads Respondidos', 'Propostas Aceitas', 'Assinaturas Finalizadas'],
        'Valores': [total_leads, leads_qualificados, leads_respondidos, propostas_aceitas, assinaturas_finalizadas]
    }
    fig_funil = px.funnel(funil_dados, x='Valores', y='Etapas', title="Funil de Vendas")
    col7.plotly_chart(fig_funil, use_container_width=True)

    # Análise de tempo de fechamento
    df_filtrado['Tempo de Fechamento'] = (df_filtrado['Data da assinatura'] - df_filtrado['Data da mensagem']).dt.days
    tempo_medio_fechamento = df_filtrado['Tempo de Fechamento'].mean()
    col8.metric("Tempo Médio de Fechamento (dias)", f"{tempo_medio_fechamento:.2f}" if not pd.isna(tempo_medio_fechamento) else "N/A")

    # Gráfico de distribuição do tempo de fechamento
    fig_tempo = px.histogram(df_filtrado, x='Tempo de Fechamento', nbins=30, title="Distribuição do Tempo de Fechamento")
    st.plotly_chart(fig_tempo, use_container_width=True)

except Exception as e:
    st.error(f"Erro ao carregar a planilha: {e}")
