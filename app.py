import streamlit as st
import pandas as pd
from statsmodels.tsa.api import VAR
from datetime import date
from dateutil.relativedelta import relativedelta
import io
import warnings
import numpy as np

# Ignorar avisos comuns do statsmodels para uma interface mais limpa
warnings.filterwarnings("ignore")

# --- Funções Auxiliares ---

@st.cache_data
def carregar_e_preparar_dados_multivariado(arquivo_carregado):
    """
    Função atualizada para processar múltiplos anos a partir do cabeçalho
    do Excel no formato 'Mês AA' (ex: 'Janeiro 24').
    """
    try:
        df = pd.read_excel(arquivo_carregado, header=[0, 1], index_col=0)
        
        lista_operacoes_ordenada = df.index.drop_duplicates().tolist()
        
        df.columns = [f'{level1}_{level2}' for level1, level2 in df.columns]
        df = df.reset_index()

        nome_antigo_da_coluna_operacoes = df.columns[0]
        df.rename(columns={nome_antigo_da_coluna_operacoes: 'Operações'}, inplace=True)
        
        df_melted = df.melt(id_vars=['Operações'], var_name='MesAno_Metrica', value_name='Valor')
        
        df_melted[['Mes_Ano', 'Metrica']] = df_melted['MesAno_Metrica'].str.rsplit('_', n=1, expand=True)
        df_melted[['Mês', 'Ano_str']] = df_melted['Mes_Ano'].str.split(' ', n=1, expand=True)

        df_filtrado = df_melted[df_melted['Metrica'].isin(['Real', 'FCT'])].copy()
        
        df_filtrado.dropna(subset=['Valor'], inplace=True)
        df_filtrado = df_filtrado[pd.to_numeric(df_filtrado['Valor'], errors='coerce').notna()]
        df_filtrado['Valor'] = df_filtrado['Valor'].astype(float)
        
        df_pivot = df_filtrado.pivot_table(index=['Operações', 'Mês', 'Ano_str'], columns='Metrica', values='Valor').reset_index()
        
        meses_map = {
            'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4, 'Maio': 5, 'Junho': 6,
            'Julho': 7, 'Agosto': 8, 'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12
        }
        
        df_pivot['Ano'] = '20' + df_pivot['Ano_str']
        df_pivot['NumeroMes'] = df_pivot['Mês'].map(meses_map)
        
        df_pivot['Data'] = pd.to_datetime(df_pivot['Ano'].astype(str) + '-' + df_pivot['NumeroMes'].astype(str) + '-01')
        
        for col in ['Real', 'FCT']:
            if col not in df_pivot.columns:
                st.error(f"A coluna '{col}' é necessária mas não foi encontrada. Verifique o arquivo Excel.")
                return None, None
        
        df_final = df_pivot[['Operações', 'Data', 'Real', 'FCT']].sort_values(by=['Operações', 'Data'])
        return df_final, lista_operacoes_ordenada

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
        st.info("Por favor, verifique se o formato do arquivo Excel está correto (cabeçalhos como 'Janeiro 24').")
        return None, None

def convert_df_to_excel(df):
    """Converte um dataframe para um arquivo Excel em memória."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Projeções')
    processed_data = output.getvalue()
    return processed_data

# --- Interface do Streamlit ---

st.set_page_config(layout="wide", page_title="Dashboard de Projeção de Interações")

st.header("📊 Projeção e Análise de Forecast")

st.markdown("""
Use as abas abaixo para navegar entre a **projeção de volumes futuros** e a **análise da acuracidade** das suas projeções históricas.
""")

arquivo_carregado = st.file_uploader("Carregue sua planilha Excel de dados", type=['xlsx'])

if arquivo_carregado:
    df_dados, lista_operacoes = carregar_e_preparar_dados_multivariado(arquivo_carregado)
    
    if df_dados is not None and not df_dados.empty:
        
        tab1, tab2 = st.tabs(["Projeção Futura", "Análise de Acuracidade Histórica"])

        # --- ABA 1: PROJEÇÃO FUTURA ---
        with tab1:
            st.sidebar.header("Parâmetros da Projeção")

            ano_atual = date.today().year
            proximo_mes = date.today() + relativedelta(months=1)
            
            lista_anos = list(range(ano_atual, ano_atual + 5))
            ano_selecionado = st.sidebar.selectbox("Selecione o Ano", options=lista_anos, index=lista_anos.index(proximo_mes.year))

            meses_pt = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            mes_selecionado_nome = st.sidebar.selectbox("Selecione o Mês", options=meses_pt, index=proximo_mes.month - 1)
            
            if st.sidebar.button("🚀 Gerar Projeções"):
                mes_selecionado_numero = meses_pt.index(mes_selecionado_nome) + 1
                data_alvo = date(ano_selecionado, mes_selecionado_numero, 1)

                resultados_projecao = []
                
                progress_bar = st.progress(0, text="Iniciando projeções...")
                
                for i, operacao in enumerate(lista_operacoes):
                    progress_bar.progress((i + 1) / len(lista_operacoes), text=f"Projetando para: {operacao}...")
                    df_operacao = df_dados[df_dados['Operações'] == operacao][['Real', 'FCT']].dropna()
                    
                    if len(df_operacao) >= 5: 
                        ultima_data = df_dados[df_dados['Operações'] == operacao]['Data'].max().date()
                        meses_para_projetar = (data_alvo.year - ultima_data.year) * 12 + (data_alvo.month - ultima_data.month)

                        if meses_para_projetar > 0:
                            try:
                                # --- LÓGICA DE CORREÇÃO DO ERRO ---
                                # Verifica se alguma das colunas é constante
                                is_constant = df_operacao.nunique().min() == 1
                                # Se for constante, ajusta o modelo para não adicionar sua própria constante ('n' = no trend)
                                trend_param = 'n' if is_constant else 'c'
                                
                                # Instancia o modelo com o parâmetro de tendência ajustado
                                modelo = VAR(df_operacao, trend=trend_param)
                                # ------------------------------------

                                resultado_modelo = modelo.fit()
                                projecao = resultado_modelo.forecast(df_operacao.values, steps=meses_para_projetar)
                                valor_projetado = projecao[-1]
                                fct_projetado = valor_projetado[1]
                                resultados_projecao.append({'Operação': operacao, 'Mês Projetado': data_alvo.strftime("%m/%Y"), 'Projeção (FCT)': int(round(fct_projetado, 0))})
                            except Exception as model_error:
                                 resultados_projecao.append({'Operação': operacao, 'Mês Projetado': data_alvo.strftime("%m/%Y"), 'Projeção (FCT)': f"Erro no modelo: {model_error}"})
                        else:
                             resultados_projecao.append({'Operação': operacao, 'Mês Projetado': data_alvo.strftime("%m/%Y"), 'Projeção (FCT)': "Mês alvo está no passado ou presente."})
                    else:
                        resultados_projecao.append({'Operação': operacao, 'Mês Projetado': data_alvo.strftime("%m/%Y"), 'Projeção (FCT)': "Dados insuficientes (mínimo de 5 meses)."})
                
                progress_bar.empty()
                st.success("Projeções concluídas!")

                df_resultados = pd.DataFrame(resultados_projecao)
                st.dataframe(df_resultados, use_container_width=True, height=len(df_resultados)*36+38)
                
                excel_data = convert_df_to_excel(df_resultados)
                st.download_button(
                   label="📥 Baixar Resultados em Excel",
                   data=excel_data,
                   file_name=f'projecao_fct_{data_alvo.strftime("%m-%Y")}.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

        # --- ABA 2: ANÁLISE DE ACURACIDADE ---
        with tab2:
            st.subheader("Análise de Acuracidade do Forecast Histórico")
            st.markdown("""
            Esta análise mede a performance das suas projeções passadas (`FCT`) em comparação com os dados `Reais`.
            * **Bias (Viés):** Mostra se você tem uma tendência a projetar para mais (valor negativo) ou para menos (valor positivo). Um valor perto de zero é ideal.
            * **MAE (Erro Absoluto Médio):** A média do erro em unidades. "Em média, errei por X interações".
            * **MAPE (Erro Percentual Absoluto Médio):** A média do erro em porcentagem. É ótimo para comparar a acuracidade entre operações de tamanhos diferentes.
            """)
            
            resultados_acuracidade = []
            for operacao in lista_operacoes:
                df_op = df_dados[df_dados['Operações'] == operacao].dropna(subset=['Real', 'FCT'])
                
                if not df_op.empty:
                    erro = df_op['Real'] - df_op['FCT']
                    bias = np.mean(erro)
                    mae = np.mean(np.abs(erro))
                    mape_df = df_op[df_op['Real'] != 0]
                    mape = np.mean(np.abs((mape_df['Real'] - mape_df['FCT']) / mape_df['Real'])) * 100 if not mape_df.empty else 0
                    resultados_acuracidade.append({'Operação': operacao, 'Bias (Viés)': int(round(bias,0)), 'MAE': int(round(mae,0)), 'MAPE (%)': f"{mape:.2f}%"})
            
            if resultados_acuracidade:
                df_acuracidade = pd.DataFrame(resultados_acuracidade)
                st.dataframe(df_acuracidade, use_container_width=True)
