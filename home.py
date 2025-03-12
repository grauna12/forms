import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
import io
import os
import gspread
from google.oauth2.service_account import Credentials
import json
import logging
from xlsxwriter import Workbook

# Configuração da página
st.set_page_config(
    page_title="Sistema de Avaliação de Funcionários",
    page_icon="📋",
    layout="wide"
)

# Lista de funcionários predefinidos
FUNCIONARIOS_PREDEFINIDOS = [
    "Alan",
    "Carlos Henrique",
    "Darlysson",
    "Domingos",
    "Douglas",
    "Francivaldo",
    "Hesletti",
    "Joelson",
    "Thiago Victor",
    "Wanderson",
    "Clemerson",
    "Samuel",
    "Yago",
    "João Davi"
]

# Define as colunas padrão do sistema
COLUNAS_PADRAO = [
    "Nome", "Semana", "Box/Uniforme", "Ferramentas", "EPI", "Horário", 
    "Apontamento", "Execução de Serviços", "Uso de celular/Indisciplinas", 
    "Revisão de entrega(Padrão Honda)", "Observações", "Data de Avaliação"
]

# Função para autenticar e conectar ao Google Sheets usando st.secrets
@st.cache_resource
def connect_to_gsheet(sheet_name):
    logging.info(f"Tentando conectar ao Google Sheets. Sheet name: {sheet_name}")
    
    try:
        # Carrega as credenciais do st.secrets
        credentials_info = {
            "type": st.secrets["general"]["type"],
            "project_id": st.secrets["general"]["project_id"],
            "private_key_id": st.secrets["general"]["private_key_id"],
            "private_key": st.secrets["general"]["private_key"],
            "client_email": st.secrets["general"]["client_email"],
            "client_id": st.secrets["general"]["client_id"],
            "auth_uri": st.secrets["general"]["auth_uri"],
            "token_uri": st.secrets["general"]["token_uri"],
            "auth_provider_x509_cert_url": st.secrets["general"]["auth_provider_x509_cert_url"],
            "client_x509_cert_url": st.secrets["general"]["client_x509_cert_url"],
            "universe_domain": st.secrets["general"]["universe_domain"]
        }
        
        credentials = Credentials.from_service_account_info(
            credentials_info, 
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
        )
        logging.info("Credenciais carregadas com sucesso.")
        
        client = gspread.authorize(credentials)
        logging.info("Cliente gspread autorizado.")
        
        try:
            # Verifica se a planilha existe, caso contrário cria uma nova
            try:
                spreadsheet = client.open(sheet_name)
                logging.info(f"Planilha '{sheet_name}' aberta com sucesso.")
            except gspread.exceptions.SpreadsheetNotFound:
                # Cria uma nova planilha se não existir
                spreadsheet = client.create(sheet_name)
                # Compartilha com a conta de serviço para garantir acesso
                spreadsheet.share(credentials_info['client_email'], perm_type='user', role='writer')
                logging.info(f"Planilha '{sheet_name}' criada com sucesso.")
            
            sheet = spreadsheet.sheet1
            logging.info("Primeira aba da planilha acessada.")
            
            # Verifica se a estrutura está correta e inicializa se necessário
            initialize_sheet_structure(sheet)
            
            return sheet
        except gspread.exceptions.WorksheetNotFound:
            logging.error("Primeira aba da planilha não encontrada.")
            st.error("Primeira aba da planilha não encontrada.")
        except gspread.exceptions.APIError as e:
            logging.error(f"Erro da API do Google Sheets: {e}")
            st.error(f"Erro da API do Google Sheets: {e}")
        except Exception as e:
            logging.error(f"Erro inesperado ao acessar a planilha: {e}")
            st.error(f"Erro inesperado ao acessar a planilha: {e}")
    except KeyError as e:
        logging.error(f"Erro ao carregar credenciais de st.secrets: {e}")
        st.error(f"Erro ao carregar credenciais de st.secrets. Verifique se as secrets estão configuradas corretamente: {e}")
    except Exception as e:
        logging.error(f"Erro ao carregar credenciais ou autorizar cliente: {e}")
        st.error(f"Erro ao carregar credenciais ou autorizar cliente: {e}")
    
    return None

# Função para inicializar a estrutura da planilha se necessário
def initialize_sheet_structure(sheet):
    try:
        # Tenta ler os valores da primeira linha (cabeçalho)
        headers = sheet.row_values(1)
        
        # Se estiver vazio ou não contiver todas as colunas necessárias
        if not headers:
            logging.info("Planilha vazia. Inicializando com cabeçalho padrão.")
            sheet.append_row(COLUNAS_PADRAO)
        elif not all(col in headers for col in COLUNAS_PADRAO):
            # Verifica se as colunas padrão estão todas presentes
            logging.info("Cabeçalho da planilha não contém todas as colunas necessárias. Atualizando estrutura.")
            
            # Limpa a primeira linha e adiciona o cabeçalho correto
            sheet.delete_rows(1, 1)
            sheet.append_row(COLUNAS_PADRAO)
        
        logging.info("Estrutura da planilha verificada e inicializada com sucesso.")
    except Exception as e:
        logging.error(f"Erro ao inicializar estrutura da planilha: {e}")
        raise e

# Função para salvar dados
def save_data(df):
    return df.to_csv(index=False).encode('utf-8')

# Função para carregar dados do Google Sheets
def load_data_from_gsheet():
    # Use o nome da planilha de st.secrets se disponível, caso contrário use o padrão
    sheet_name = st.secrets.get("SHEET_NAME", "forms oficina")
    sheet = connect_to_gsheet(sheet_name)
    
    if sheet:
        try:
            # Obtém todos os dados da planilha
            records = sheet.get_all_records()
            if records:
                return pd.DataFrame(records)
            else:
                # Se a planilha estiver vazia, retorna DataFrame vazio
                return pd.DataFrame(columns=COLUNAS_PADRAO)
        except Exception as e:
            st.error(f"Erro ao carregar dados da planilha: {e}")
            return pd.DataFrame(columns=COLUNAS_PADRAO)
    else:
        return pd.DataFrame(columns=COLUNAS_PADRAO)

# Função para salvar novos dados no Google Sheets
def save_to_gsheet(new_data):
    # Use o nome da planilha de st.secrets se disponível, caso contrário use o padrão
    sheet_name = st.secrets.get("SHEET_NAME", "forms oficina")
    sheet = connect_to_gsheet(sheet_name)
    
    if sheet:
        try:
            # Verifica se a planilha tem cabeçalho
            existing_data = sheet.get_all_values()
            if not existing_data:
                # Se a planilha estiver vazia, adiciona cabeçalho
                sheet.append_row(COLUNAS_PADRAO)
            
            # Adiciona nova linha de dados
            row_values = [new_data.get(col, "") for col in COLUNAS_PADRAO]
            sheet.append_row(row_values)
            return True
        except Exception as e:
            st.error(f"Erro ao salvar dados na planilha: {e}")
            import traceback
            st.error(f"Detalhes do erro: {traceback.format_exc()}")
            return False
    else:
        return False

# Função para converter valores string para numéricos para cálculos e gráficos
def converter_valor_numerico(valor):
    if isinstance(valor, str):
        if valor == "Positivo":
            return 1
        elif valor == "Ruim":
            return 0
    return valor

# Título principal
st.title("Sistema de Avaliação de Funcionários")

# Inicializa o estado da sessão para armazenar dados
if 'funcionarios_df' not in st.session_state:
    # Carrega dados existentes do Google Sheets
    st.session_state.funcionarios_df = load_data_from_gsheet()
    
    # Converte valores antigos para novos valores, se necessário
    if not st.session_state.funcionarios_df.empty:
        for criterio in COLUNAS_PADRAO[2:-2]:  # Excluindo Nome, Semana, Observações e Data
            if criterio in st.session_state.funcionarios_df.columns:
                st.session_state.funcionarios_df[criterio] = st.session_state.funcionarios_df[criterio].replace(
                    {"Satisfatório": "Positivo", "Insatisfatório": "Ruim"}
                )

# Sidebar para navegação
st.sidebar.title("Menu")
pagina = st.sidebar.radio("Ir para:", ["Nova Avaliação", "Histórico", "Relatórios"])

# Página de Nova Avaliação
if pagina == "Nova Avaliação":
    st.header("Nova Avaliação de Funcionário")
    
    # Formulário para adicionar/selecionar funcionário
    col1, col2, col3 = st.columns([2, 1, 1])
    
    with col1:
        opcao_funcionario = st.radio("Selecionar opção:", ["Selecionar funcionário existente", "Adicionar novo funcionário"])
        
        if opcao_funcionario == "Selecionar funcionário existente":
            # Solução corrigida - Combinando funcionários existentes e predefinidos
            nomes_existentes = sorted(list(st.session_state.funcionarios_df["Nome"].unique()))
            
            # Combine os nomes existentes com os predefinidos, removendo duplicatas
            todos_funcionarios = sorted(list(set(nomes_existentes + FUNCIONARIOS_PREDEFINIDOS)))
            
            # Selecione da lista combinada
            nome_funcionario = st.selectbox("Selecione o funcionário:", todos_funcionarios)
        else:
            # Opção para adicionar novo funcionário - usa a lista predefinida
            nome_funcionario = st.selectbox(
                "Selecione o funcionário:",
                options=FUNCIONARIOS_PREDEFINIDOS + ["Outro (especificar)"]
            )
            
            if nome_funcionario == "Outro (especificar)":
                nome_funcionario = st.text_input("Digite o nome do funcionário:")
    
    with col2:
        semana = st.number_input("Número da semana:", min_value=1, max_value=5, value=1)
    
    with col3:
        data_avaliacao = st.date_input("Data da avaliação:", datetime.now())
    
    # Critérios de avaliação
    st.subheader("Critérios de Avaliação")
    st.info("Avalie o funcionário em cada critério como 'Positivo' ou 'Ruim'")
    
    criterios = {
        "Box/Uniforme": "Avaliação da organização do espaço de trabalho e uso correto do uniforme",
        "Ferramentas": "Utilização e cuidado com as ferramentas de trabalho",
        "EPI": "Uso correto dos Equipamentos de Proteção Individual",
        "Horário": "Pontualidade e assiduidade",
        "Apontamento": "Registro e documentação das atividades",
        "Execução de Serviços": "Qualidade e eficiência na execução das tarefas",
        "Uso de celular/Indisciplinas": "Comportamento apropriado no ambiente de trabalho",
        "Revisão de entrega(Padrão Honda)": "Conformidade com os padrões de qualidade Honda"
    }
    
    avaliacoes = {}
    col1, col2 = st.columns(2)
    
    for i, (criterio, descricao) in enumerate(criterios.items()):
        with col1 if i % 2 == 0 else col2:
            st.write(f"**{criterio}**")
            st.caption(descricao)
            avaliacoes[criterio] = st.radio(
                f"Avaliação para {criterio}", 
                options=["Positivo", "Ruim"],
                key=f"radio_{criterio}",
                horizontal=True
            )
    
    observacoes = st.text_area("Observações adicionais:", height=100)
    
    # Botão para registrar avaliação
    if st.button("Registrar Avaliação", type="primary"):
        if nome_funcionario and nome_funcionario != "Outro (especificar)":
            nova_avaliacao = {
                "Nome": nome_funcionario,
                "Semana": semana,
                "Box/Uniforme": avaliacoes["Box/Uniforme"],
                "Ferramentas": avaliacoes["Ferramentas"],
                "EPI": avaliacoes["EPI"],
                "Horário": avaliacoes["Horário"],
                "Apontamento": avaliacoes["Apontamento"],
                "Execução de Serviços": avaliacoes["Execução de Serviços"],
                "Uso de celular/Indisciplinas": avaliacoes["Uso de celular/Indisciplinas"],
                "Revisão de entrega(Padrão Honda)": avaliacoes["Revisão de entrega(Padrão Honda)"],
                "Observações": observacoes,
                "Data de Avaliação": data_avaliacao.strftime("%Y-%m-%d")
            }
            
            # Salvar na planilha do Google Sheets
            if save_to_gsheet(nova_avaliacao):
                # Adicionar nova avaliação ao DataFrame local
                st.session_state.funcionarios_df = pd.concat([
                    st.session_state.funcionarios_df, 
                    pd.DataFrame([nova_avaliacao])
                ], ignore_index=True)
                
                st.success(f"Avaliação de {nome_funcionario} na semana {semana} registrada com sucesso!")
            else:
                st.error("Houve um erro ao salvar a avaliação na planilha do Google Sheets.")
        else:
            st.error("Por favor, informe o nome do funcionário.")

# Página de Histórico
elif pagina == "Histórico":
    st.header("Histórico de Avaliações")
    
    # Botão para atualizar dados do Google Sheets
    if st.button("Atualizar dados do Google Sheets"):
        st.session_state.funcionarios_df = load_data_from_gsheet()
        # Converte valores antigos para novos valores, se necessário
        if not st.session_state.funcionarios_df.empty:
            for criterio in COLUNAS_PADRAO[2:-2]:  # Excluindo Nome, Semana, Observações e Data
                if criterio in st.session_state.funcionarios_df.columns:
                    st.session_state.funcionarios_df[criterio] = st.session_state.funcionarios_df[criterio].replace(
                        {"Satisfatório": "Positivo", "Insatisfatório": "Ruim"}
                    )
        st.success("Dados atualizados com sucesso!")
    
    if st.session_state.funcionarios_df.empty:
        st.warning("Não há avaliações registradas.")
    else:
        # Filtros
        col1, col2 = st.columns(2)
        with col1:
            nomes_unicos = sorted(list(st.session_state.funcionarios_df["Nome"].unique()))
            nome_filtro = st.multiselect("Filtrar por funcionário:", nomes_unicos, default=nomes_unicos)
        
        with col2:
            semanas_unicas = sorted(list(st.session_state.funcionarios_df["Semana"].unique()))
            semana_filtro = st.multiselect("Filtrar por semana:", semanas_unicas, default=semanas_unicas)
        
        # Aplicar filtros
        df_filtrado = st.session_state.funcionarios_df.copy()
        if nome_filtro:
            df_filtrado = df_filtrado[df_filtrado["Nome"].isin(nome_filtro)]
        if semana_filtro:
            df_filtrado = df_filtrado[df_filtrado["Semana"].isin(semana_filtro)]
        
        # Exibir dados filtrados
        if not df_filtrado.empty:
            st.dataframe(df_filtrado, use_container_width=True)
            
            # Botões para download
            st.download_button(
                label="Download CSV",
                data=save_data(df_filtrado),
                file_name=f"avaliacao_funcionarios_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
            
            # Botão para download Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_filtrado.to_excel(writer, index=False, sheet_name='Avaliações')
            
            st.download_button(
                label="Download Excel",
                data=buffer.getvalue(),
                file_name=f"avaliacao_funcionarios_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.ms-excel"
            )
        else:
            st.warning("Nenhum registro encontrado com os filtros aplicados.")

# Página de Relatórios
elif pagina == "Relatórios":
    st.header("Relatórios e Análises")
    
    if st.session_state.funcionarios_df.empty:
        st.warning("Não há dados para gerar relatórios.")
    else:
        # Seleção de funcionário para relatório detalhado
        nomes_unicos = sorted(list(st.session_state.funcionarios_df["Nome"].unique()))
        nome_relatorio = st.selectbox("Selecione o funcionário para relatório detalhado:", nomes_unicos)
        
        # Filtrar dados do funcionário selecionado
        df_funcionario = st.session_state.funcionarios_df[st.session_state.funcionarios_df["Nome"] == nome_relatorio]
        
        if not df_funcionario.empty:
            # Critérios de avaliação para análise
            criterios_avaliacao = [
                "Box/Uniforme", "Ferramentas", "EPI", "Horário", "Apontamento",
                "Execução de Serviços", "Uso de celular/Indisciplinas", "Revisão de entrega(Padrão Honda)"
            ]
            
            # Converter valores de texto para números para cálculos
            df_numerica = df_funcionario.copy()
            for criterio in criterios_avaliacao:
                df_numerica[criterio] = df_numerica[criterio].apply(lambda x: 1 if x == "Positivo" else 0)
            
            # Calcular percentual de "Positivo" por critério
            percentuais = {}
            for criterio in criterios_avaliacao:
                total_avaliacoes = len(df_numerica)
                positivos = df_numerica[criterio].sum()
                percentuais[criterio] = (positivos / total_avaliacoes * 100) if total_avaliacoes > 0 else 0
            
            # Calcular percentual geral de avaliações positivas
            percentual_geral = sum(percentuais.values()) / len(percentuais) if percentuais else 0
            
            # Exibir resumo
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.subheader(f"Resumo de {nome_relatorio}")
                st.metric("Percentual Positivo", f"{percentual_geral:.1f}%")
                st.metric("Total de Avaliações", len(df_funcionario))
                
                # Determinar status baseado no percentual
                if percentual_geral >= 80:
                    status = "✅ Excelente"
                elif percentual_geral >= 60:
                    status = "✓ Positivo"
                elif percentual_geral >= 40:
                    status = "⚠️ Precisa Melhorar"
                else:
                    status = "❌ Ruim"
                
                st.metric("Status", status)
            
            with col2:
                # Gráfico de barras para todos os critérios
                fig, ax = plt.subplots(figsize=(10, 6))
                criterios_curtos = [c[:15] + '...' if len(c) > 15 else c for c in criterios_avaliacao]
                valores = [percentuais[c] for c in criterios_avaliacao]
                
                # Colorir barras com base no valor
                cores = ['green' if v >= 60 else 'red' for v in valores]
                
                ax.bar(criterios_curtos, valores, color=cores)
                ax.set_ylabel('% Positivo')
                ax.set_ylim(0, 100)
                ax.set_title(f'Desempenho por Critério - {nome_relatorio}')
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()
                
                st.pyplot(fig)
            
            # Tabela de desempenho por critério
            st.subheader("Desempenho por Critério")
            df_desempenho = pd.DataFrame({
                'Critério': criterios_avaliacao,
                'Percentual Positivo': [f"{percentuais[c]:.1f}%" for c in criterios_avaliacao],
                'Status': ["Positivo" if percentuais[c] >= 60 else "Ruim" for c in criterios_avaliacao]
            })
            st.dataframe(df_desempenho.sort_values('Percentual Positivo', ascending=False), use_container_width=True)
            
            # Evolução ao longo do tempo (por semana)
            st.subheader("Evolução ao Longo do Tempo")
            
            if len(df_funcionario) > 1:  # Só faz sentido se houver mais de uma avaliação
                # Agrupar por semana e calcular percentual de critérios positivos
                df_evolucao = df_numerica.copy()
                evolucao_semanas = {}
                
                for _, grupo in df_evolucao.groupby('Semana'):
                    semana = grupo['Semana'].iloc[0]
                    percentual = grupo[criterios_avaliacao].mean().mean() * 100
                    evolucao_semanas[semana] = percentual
                
                # Ordenar semanas
                semanas_ordenadas = sorted(evolucao_semanas.keys())
                percentuais_ordenados = [evolucao_semanas[s] for s in semanas_ordenadas]
                
                # Plotar gráfico de linha
                fig, ax = plt.subplots(figsize=(10, 5))
                ax.plot(semanas_ordenadas, percentuais_ordenados, marker='o', linestyle='-', linewidth=2)
                ax.set_xlabel('Semana')
                ax.set_ylabel('% Positivo')
                ax.set_ylim(0, 100)
                ax.grid(True, linestyle='--', alpha=0.7)
                plt.title(f"Evolução de Desempenho de {nome_relatorio} por Semana")
                
                st.pyplot(fig)
            else:
                st.info("São necessárias múltiplas avaliações para mostrar a evolução ao longo do tempo.")
            
            # Áreas para melhoria
            st.subheader("Áreas para Melhoria")
            areas_melhoria = [(c, percentuais[c]) for c in criterios_avaliacao]
            areas_melhoria.sort(key=lambda x: x[1])
            
            for criterio, percentual in areas_melhoria[:3]:
                if percentual < 60:
                    st.warning(f"**{criterio}**: {percentual:.1f}% Positivo - Precisa de atenção especial")
                else:
                    st.info(f"**{criterio}**: {percentual:.1f}% Positivo - Potencial para melhoria")
                    
            # Download do relatório em Excel
            if st.button("Exportar Relatório Completo"):
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    # Dados individuais
                    df_funcionario.to_excel(writer, index=False, sheet_name='Avaliações')
                    
                    # Percentuais por critério
                    pd.DataFrame({
                        'Critério': criterios_avaliacao,
                        'Percentual Positivo': [percentuais[c] for c in criterios_avaliacao],
                        'Status': ["Positivo" if percentuais[c] >= 60 else "Ruim" for c in criterios_avaliacao]
                    }).to_excel(writer, index=False, sheet_name='Percentuais por Critério')
                    
                    # Evolução por semana (se disponível)
                    if len(df_funcionario) > 1 and evolucao_semanas:
                        pd.DataFrame({
                            'Semana': semanas_ordenadas,
                            '% Positivo': percentuais_ordenados
                        }).to_excel(writer, index=False, sheet_name='Evolução Temporal')
                
                st.download_button(
                    label="Download Relatório Excel",
                    data=buffer.getvalue(),
                    file_name=f"relatorio_{nome_relatorio}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.ms-excel"
                )
        else:
            st.warning(f"Não há avaliações registradas para {nome_relatorio}.")

# Configurações
with st.sidebar.expander("Configurações"):
    # Opção para configurar a planilha
    if st.checkbox("Mostrar configurações de planilha"):
        sheet_name = st.text_input("Nome da planilha Google Sheets:", value=st.secrets.get("SHEET_NAME", "forms oficina"))
        
        if st.button("Testar conexão"):
            sheet = connect_to_gsheet(sheet_name)
            if sheet:
                st.success("Conexão com a planilha estabelecida com sucesso!")
            else:
                st.error("Falha na conexão com a planilha. Verifique as credenciais.")

        # Adicionar seção para ajudar na depuração
        st.subheader("Ajuda para depuração")
        st.info("""
        Se você estiver tendo problemas com a conexão, verifique:
        1. Que as secrets do Streamlit estão configuradas corretamente (na seção Secrets do dashboard do Streamlit)
        2. Que a conta de serviço tem acesso à planilha
        """)
    
    # Opção para importar funcionários em massa
    st.subheader("Importar Funcionários em Massa")
    importar_funcionarios = st.checkbox("Mostrar importação em massa")
    
    if importar_funcionarios:
        st.info("Cole a lista de funcionários abaixo, um por linha, para importar todos de uma vez.")
        
        funcionarios_input = st.text_area(
            "Lista de funcionários:",
            "\n".join(FUNCIONARIOS_PREDEFINIDOS),
            height=200
        )
        
        if st.button("Importar Todos os Funcionários"):
            funcionarios_lista = [nome.strip() for nome in funcionarios_input.split("\n") if nome.strip()]
            if funcionarios_lista:
                # Verificar quais funcionários já existem no sistema
                nomes_existentes = set(st.session_state.funcionarios_df["Nome"].unique())
                novos_funcionarios = []
                
                for nome in funcionarios_lista:
                    if nome not in nomes_existentes:
                        novos_funcionarios.append(nome)
                        # Adicionar funcionário com uma avaliação inicial (opcional)
                        nova_avaliacao = {
                            "Nome": nome,
                            "Semana": 1,
                            "Box/Uniforme": "Positivo",
                            "Ferramentas": "Positivo",
                            "EPI": "Positivo",
                            "Horário": "Positivo",
                            "Apontamento": "Positivo",
                            "Execução de Serviços": "Positivo",
                            "Uso de celular/Indisciplinas": "Positivo",
                            "Revisão de entrega(Padrão Honda)": "Positivo",
                            "Observações": "Importação inicial",
                            "Data de Avaliação": datetime.now().strftime("%Y-%m-%d")
                        }
                        save_to_gsheet(nova_avaliacao)
                
                if novos_funcionarios:
                    # Recarregar dados
                    st.session_state.funcionarios_df = load_data_from_gsheet()
                    st.success(f"{len(novos_funcionarios)} funcionários importados com sucesso!")
                else:
                    st.info("Todos os funcionários já existem no sistema.")

# Rodapé
st.markdown("---")
st.caption("Sistema de Avaliação de Funcionários - Versão 1.1")