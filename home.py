# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
# import numpy as np # numpy não está sendo usado diretamente
from datetime import datetime
import io
# import os # os não está sendo usado
import gspread
from google.oauth2.service_account import Credentials
# import json # json não está sendo usado
import logging
# from xlsxwriter import Workbook # Usado implicitamente por pd.ExcelWriter

# --- Configuração de Logging ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Configuração da página e Colunas Padrão ---
st.set_page_config(
    page_title="Sistema de Avaliação de Funcionários",
    page_icon="📋",
    layout="wide"
)

COLUNAS_PADRAO = [
    "Nome", "Semana", "Box/Uniforme", "Ferramentas", "EPI", "Horário",
    "Apontamento", "Execução de Serviços", "Uso de celular/Indisciplinas",
    "Quantidade de Retorno", "Observações", "Data de Avaliação"
]
CRITERIOS_AVALIACAO_COLS = COLUNAS_PADRAO[2:-2] # Define as colunas de critérios dinamicamente

# --- Funções para Google Sheets (sem alterações significativas) ---
@st.cache_resource
def get_gspread_client():
    logging.info("Tentando obter cliente gspread.")
    try:
        # Assegure que todas essas chaves existem em st.secrets.general
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
        client = gspread.authorize(credentials)
        logging.info("Cliente gspread autorizado com sucesso.")
        return client
    except KeyError as e:
        logging.error(f"Erro ao carregar credenciais de st.secrets: Chave ausente '{e}'")
        st.error(f"Erro ao carregar credenciais: A chave '{e}' não foi encontrada nos secrets. Verifique a configuração do Streamlit Cloud.")
        return None
    except Exception as e:
        logging.error(f"Erro ao carregar credenciais ou autorizar cliente: {e}")
        st.error(f"Erro ao carregar credenciais ou autorizar cliente: {e}")
        return None

@st.cache_resource
def get_spreadsheet(_client, sheet_name):
    logging.info(f"Tentando abrir/criar planilha: {sheet_name}")
    if not _client:
        logging.error("Tentativa de obter planilha com cliente Nulo.")
        st.error("Erro interno: Conexão com Google Sheets não estabelecida.")
        return None
    try:
        spreadsheet = _client.open(sheet_name)
        logging.info(f"Planilha '{sheet_name}' aberta.")
        return spreadsheet
    except gspread.exceptions.SpreadsheetNotFound:
        logging.info(f"Planilha '{sheet_name}' não encontrada. Criando...")
        try:
            spreadsheet = _client.create(sheet_name)
            # Compartilha com o email de serviço para garantir permissão de escrita
            spreadsheet.share(st.secrets["general"]["client_email"], perm_type='user', role='writer')
            logging.info(f"Planilha '{sheet_name}' criada e compartilhada.")
            # Tenta adicionar abas padrão se a planilha for nova
            try:
                 get_worksheet(spreadsheet, FUNCIONARIOS_SHEET_NAME, headers=FUNCIONARIOS_HEADER)
                 get_worksheet(spreadsheet, AVALIACOES_SHEET_NAME, headers=COLUNAS_PADRAO)
            except Exception as e:
                 logging.warning(f"Não foi possível criar/configurar abas padrão na nova planilha: {e}")
            return spreadsheet
        except Exception as e:
            logging.error(f"Erro ao criar planilha '{sheet_name}': {e}")
            st.error(f"Erro ao criar planilha '{sheet_name}'. Verifique as permissões da conta de serviço no Google Drive.")
            return None
    except Exception as e:
        logging.error(f"Erro ao abrir planilha '{sheet_name}': {e}")
        st.error(f"Erro ao abrir planilha '{sheet_name}'. Verifique o nome e as permissões.")
        return None


def get_worksheet(spreadsheet, worksheet_name, headers=None):
    logging.info(f"Tentando obter/criar aba: {worksheet_name}")
    if not spreadsheet:
        logging.error("Tentativa de obter worksheet com spreadsheet Nulo.")
        return None
    try:
        worksheet = spreadsheet.worksheet(worksheet_name)
        logging.info(f"Aba '{worksheet_name}' encontrada.")
        # Verifica e adiciona cabeçalho se a aba estiver vazia
        try:
            # Usar cell(1,1).value é mais leve que get_all_values ou row_values(1) para verificar se está vazia
            if worksheet.cell(1, 1).value is None and headers:
                logging.info(f"Aba '{worksheet_name}' vazia. Adicionando cabeçalho: {headers}")
                worksheet.append_row(headers)
        except Exception as e:
            logging.warning(f"Não foi possível verificar/adicionar cabeçalho na aba '{worksheet_name}'. Erro: {e}")
        return worksheet
    except gspread.exceptions.WorksheetNotFound:
        logging.info(f"Aba '{worksheet_name}' não encontrada. Criando...")
        try:
            col_count = len(headers) if headers else 10 # Define um número razoável de colunas
            worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows="100", cols=col_count)
            if headers:
                worksheet.append_row(headers)
            logging.info(f"Aba '{worksheet_name}' criada com cabeçalho.")
            return worksheet
        except Exception as e:
            logging.error(f"Erro ao criar aba '{worksheet_name}': {e}")
            st.error(f"Erro ao criar aba '{worksheet_name}'.")
            return None
    except Exception as e:
        logging.error(f"Erro ao obter aba '{worksheet_name}': {e}")
        st.error(f"Erro ao obter aba '{worksheet_name}'.")
        return None


# --- Funções para Carregar/Salvar Funcionários ---
FUNCIONARIOS_SHEET_NAME = "Funcionarios"
FUNCIONARIOS_HEADER = ["Nome"]

# @st.cache_data(ttl=300) # Cache de dados por 5 minutos
# Cache Resource é mais apropriado aqui pois o cliente é um recurso
@st.cache_data(ttl=300, show_spinner="Carregando lista de funcionários...")
def load_funcionarios_from_gsheet():
    logging.info("Carregando lista de funcionários do Google Sheets.")
    funcionarios = []
    try:
        client = get_gspread_client()
        if not client:
            return []
        # Usar um nome padrão mais genérico ou buscar do secret
        sheet_name = st.secrets.get("general", {}).get("sheet_name", "forms oficina") # Exemplo de nome
        spreadsheet = get_spreadsheet(client, sheet_name)
        if not spreadsheet:
            return []

        worksheet = get_worksheet(spreadsheet, FUNCIONARIOS_SHEET_NAME, headers=FUNCIONARIOS_HEADER)

        if worksheet:
            # Pega todos os valores da primeira coluna, exceto o cabeçalho
            all_names = worksheet.col_values(1)[1:]
            # Limpa (remove espaços extras) e filtra nomes não vazios, remove duplicados e ordena
            funcionarios = sorted(list(set(name.strip() for name in all_names if name and name.strip())))
            logging.info(f"{len(funcionarios)} funcionários únicos carregados.")
        else:
             logging.warning(f"Aba '{FUNCIONARIOS_SHEET_NAME}' não pôde ser acessada.")

    except gspread.exceptions.APIError as e:
         logging.error(f"Erro de API do Google ao carregar funcionários: {e}", exc_info=True)
         st.error(f"Erro de API ao conectar com Google Sheets: {e}. Verifique as permissões da API e da conta de serviço.")
    except Exception as e:
        logging.error(f"Erro detalhado ao carregar funcionários: {e}", exc_info=True)
        st.error("Ocorreu um erro inesperado ao carregar a lista de funcionários.")
    return funcionarios


def add_funcionario_to_gsheet(nome_funcionario):
    nome_funcionario = nome_funcionario.strip()
    if not nome_funcionario:
        st.error("O nome do funcionário não pode ser vazio.")
        return False

    logging.info(f"Tentando adicionar funcionário: {nome_funcionario}")
    try:
        client = get_gspread_client()
        if not client:
            return False # Mensagem de erro já mostrada por get_gspread_client
        sheet_name = st.secrets.get("general", {}).get("sheet_name", "forms oficina")
        spreadsheet = get_spreadsheet(client, sheet_name)
        if not spreadsheet:
            return False # Mensagem de erro já mostrada por get_spreadsheet

        worksheet = get_worksheet(spreadsheet, FUNCIONARIOS_SHEET_NAME, headers=FUNCIONARIOS_HEADER)

        if worksheet:
            # Verifica se o funcionário já existe (case-insensitive)
            current_funcionarios = worksheet.col_values(1)[1:] # Ignora cabeçalho
            current_funcionarios_lower = [f.strip().lower() for f in current_funcionarios if f and f.strip()]

            if nome_funcionario.lower() in current_funcionarios_lower:
                st.warning(f"Funcionário '{nome_funcionario}' já existe na lista.")
                return False # Não é um erro, mas a operação não foi realizada

            # Adiciona o novo funcionário
            worksheet.append_row([nome_funcionario])
            logging.info(f"Funcionário '{nome_funcionario}' adicionado com sucesso.")
            load_funcionarios_from_gsheet.clear() # Limpa o cache para refletir a mudança
            return True
        else:
            st.error(f"Não foi possível acessar a aba '{FUNCIONARIOS_SHEET_NAME}' para adicionar o funcionário.")
            return False
    except gspread.exceptions.APIError as e:
         logging.error(f"Erro de API do Google ao adicionar funcionário: {e}", exc_info=True)
         st.error(f"Erro de API ao salvar no Google Sheets: {e}. Verifique as permissões.")
         return False
    except Exception as e:
        logging.error(f"Erro ao adicionar funcionário '{nome_funcionario}': {e}", exc_info=True)
        st.error(f"Ocorreu um erro inesperado ao tentar adicionar o funcionário '{nome_funcionario}'.")
        return False

# --- Funções para Carregar/Salvar Avaliações ---
AVALIACOES_SHEET_NAME = "Avaliações"

# @st.cache_data(ttl=300) # Cache de dados por 5 minutos
@st.cache_data(ttl=300, show_spinner="Carregando avaliações...")
def load_data_from_gsheet():
    logging.info("Carregando dados de avaliações do Google Sheets.")
    df = pd.DataFrame(columns=COLUNAS_PADRAO) # Retorna df vazio em caso de erro
    try:
        client = get_gspread_client()
        if not client:
            return df
        sheet_name = st.secrets.get("general", {}).get("sheet_name", "forms oficina")
        spreadsheet = get_spreadsheet(client, sheet_name)
        if not spreadsheet:
            return df

        sheet = get_worksheet(spreadsheet, AVALIACOES_SHEET_NAME, headers=COLUNAS_PADRAO)

        if sheet:
            data = sheet.get_all_values()
            if len(data) > 1: # Verifica se há dados além do cabeçalho
                header = data[0]
                # Corrigir cabeçalho se necessário (comparar com COLUNAS_PADRAO)
                if header != COLUNAS_PADRAO:
                     logging.warning(f"Cabeçalho na planilha {header} difere do esperado {COLUNAS_PADRAO}. Usando o esperado.")
                     # Tenta remapear colunas se possível, ou usa COLUNAS_PADRAO
                     try:
                        df = pd.DataFrame(data[1:], columns=header)
                        # Adiciona colunas faltantes e reordena
                        for col in COLUNAS_PADRAO:
                             if col not in df.columns:
                                 df[col] = pd.NA
                        df = df[COLUNAS_PADRAO]
                     except Exception as e:
                          logging.error(f"Erro ao ajustar colunas do DataFrame: {e}. Usando cabeçalho da planilha.")
                          # Se falhar, usa o cabeçalho da planilha e avisa
                          df = pd.DataFrame(data[1:], columns=header)
                          st.warning("O cabeçalho da planilha de Avaliações parece diferente do esperado. Alguns dados podem não ser exibidos corretamente.")
                else:
                     df = pd.DataFrame(data[1:], columns=COLUNAS_PADRAO)


                logging.info(f"{len(df)} registros de avaliação brutos carregados.")

                # --- Processamento de Tipos ---
                # Garante que colunas padrão existam
                for col in COLUNAS_PADRAO:
                    if col not in df.columns:
                        df[col] = pd.NA # Adiciona coluna se faltar

                df = df[COLUNAS_PADRAO].copy() # Garante a ordem e cria cópia

                # Converter Semana para numérico Int64 (suporta NA)
                df["Semana"] = pd.to_numeric(df["Semana"], errors='coerce').astype('Int64')

                # Converter Data de Avaliação para datetime
                if "Data de Avaliação" in df.columns:
                    original_dates = df["Data de Avaliação"].copy()
                    # Tenta converter com múltiplos formatos comuns, incluindo YYYY-MM-DD
                    df["Data de Avaliação"] = pd.to_datetime(
                        df["Data de Avaliação"],
                        errors='coerce',
                        format=None, # Tenta inferir o formato
                        # dayfirst=True # Descomente se o formato DD/MM/YYYY for mais comum
                    )
                    failed_conversions = original_dates.notna() & df["Data de Avaliação"].isna()
                    if failed_conversions.any():
                         unique_failed = original_dates[failed_conversions].unique()
                         logging.warning(f"Não foi possível converter algumas datas para datetime. Exemplos: {unique_failed[:5]}. Verifique o formato na planilha.")
                         st.warning(f"Algumas datas ('{unique_failed[0]}', ...) não puderam ser lidas corretamente. Verifique o formato na planilha (ex: YYYY-MM-DD ou DD/MM/YYYY).", icon="⚠️")
                else:
                     logging.warning("Coluna 'Data de Avaliação' não encontrada.")
                     df["Data de Avaliação"] = pd.NaT # Adiciona como NaT se não existir


                # Limpar e padronizar colunas de critérios
                for criterio in CRITERIOS_AVALIACAO_COLS:
                    if criterio in df.columns:
                        # Converte para string, remove espaços, converte para minúsculas
                        df[criterio] = df[criterio].astype(str).str.strip().str.lower()
                        # Mapeia valores conhecidos e substitui vazios por NA do Pandas
                        df[criterio] = df[criterio].replace({
                            'satisfatório': 'positivo',
                            'insatisfatório': 'ruim',
                            '': pd.NA, # String vazia vira NA
                            'nan': pd.NA, # String 'nan' vira NA
                            'na': pd.NA,  # String 'na' vira NA
                            '<na>': pd.NA # String '<NA>' vira NA
                        })
                        # Converte para tipo 'string' do pandas que suporta NA explicitamente
                        df[criterio] = df[criterio].astype('string')
                    else:
                         df[criterio] = pd.NA # Adiciona como NA se não existir
                         df[criterio] = df[criterio].astype('string')

                logging.info(f"Dados de avaliação processados. {len(df)} linhas.")
                return df

            else:
                logging.info("Aba de avaliações vazia ou só com cabeçalho.")
                return df # Retorna DataFrame vazio com colunas padrão
        else:
            logging.warning(f"Aba '{AVALIACOES_SHEET_NAME}' não pôde ser acessada.")
            return df # Retorna DataFrame vazio

    except gspread.exceptions.APIError as e:
         logging.error(f"Erro de API do Google ao carregar avaliações: {e}", exc_info=True)
         st.error(f"Erro de API ao conectar com Google Sheets para ler avaliações: {e}. Verifique as permissões.")
         return df # Retorna DataFrame vazio
    except Exception as e:
        logging.error(f"Erro ao carregar dados de avaliações: {e}", exc_info=True)
        st.error("Ocorreu um erro inesperado ao carregar os dados de avaliações.")
        return df # Retorna DataFrame vazio


def save_to_gsheet(new_data):
    logging.info(f"Tentando salvar avaliação para: {new_data.get('Nome')}")
    try:
        client = get_gspread_client()
        if not client:
            return False
        sheet_name = st.secrets.get("general", {}).get("sheet_name", "forms oficina")
        spreadsheet = get_spreadsheet(client, sheet_name)
        if not spreadsheet:
            return False
        sheet = get_worksheet(spreadsheet, AVALIACOES_SHEET_NAME, headers=COLUNAS_PADRAO)

        if sheet:
            row_values = []
            for col in COLUNAS_PADRAO:
                value = new_data.get(col, "") # Pega valor ou string vazia

                # Formatação específica para tipos antes de enviar
                if isinstance(value, (datetime, pd.Timestamp)):
                    # Formato ISO 8601 (YYYY-MM-DD) é universalmente aceito por Sheets como data
                    value = value.strftime('%Y-%m-%d')
                elif isinstance(value, pd.NaT.__class__): # Verifica se é NaT
                     value = "" # Envia string vazia para NaT
                elif hasattr(value, 'isoformat'): # Para objetos date
                     value = value.isoformat()
                # Garantir que outros tipos sejam strings (ou números)
                elif not isinstance(value, (str, int, float, bool)) and value is not None:
                    value = str(value)
                elif value is None:
                     value = "" # None vira string vazia

                row_values.append(value)

            # Usa USER_ENTERED para que o Sheets interprete os valores (datas, números)
            sheet.append_row(row_values, value_input_option='USER_ENTERED')
            logging.info("Avaliação salva com sucesso na planilha.")
            load_data_from_gsheet.clear() # Limpa o cache para refletir a nova linha
            return True
        else:
            st.error(f"Não foi possível acessar a aba '{AVALIACOES_SHEET_NAME}' para salvar a avaliação.")
            return False
    except gspread.exceptions.APIError as e:
         logging.error(f"Erro de API do Google ao salvar avaliação: {e}", exc_info=True)
         st.error(f"Erro de API ao salvar no Google Sheets: {e}. Verifique as permissões.")
         return False
    except Exception as e:
        logging.error(f"Erro ao salvar avaliação: {e}", exc_info=True)
        st.error("Ocorreu um erro inesperado ao tentar salvar a avaliação.")
        return False

# --- Função save_data para CSV (Mantida) ---
def save_data_csv(df_to_save):
    df_csv = df_to_save.copy()
    # Formatar datas para DD/MM/YYYY no CSV
    if 'Data de Avaliação' in df_csv.columns and pd.api.types.is_datetime64_any_dtype(df_csv['Data de Avaliação']):
         df_csv['Data de Avaliação'] = df_csv['Data de Avaliação'].dt.strftime('%d/%m/%Y')

    # Mapear para 'Positivo'/'Ruim' no CSV
    for criterio in CRITERIOS_AVALIACAO_COLS:
        if criterio in df_csv.columns:
             map_dict = {'positivo': 'Positivo', 'ruim': 'Ruim'}
             #astype(str) antes de map para evitar erros com tipos mistos se NA não foi tratado antes
             df_csv[criterio] = df_csv[criterio].astype(str).map(map_dict).fillna(df_csv[criterio])

    # Substituir NA por strings vazias para o CSV
    df_csv.fillna('', inplace=True)

    return df_csv.to_csv(index=False, sep=';', decimal=',', encoding='utf-8-sig') # utf-8-sig para Excel ler acentos corretamente

# --- Inicialização do Estado da Sessão ---
# Carrega funcionários na primeira execução ou se ainda não estiverem no estado
if 'lista_funcionarios' not in st.session_state:
    st.session_state.lista_funcionarios = load_funcionarios_from_gsheet()
    logging.info(f"Lista inicial de funcionários carregada no estado da sessão: {len(st.session_state.lista_funcionarios)} nomes.")

# Carrega dados de avaliação na primeira execução ou se ainda não estiverem no estado
if 'funcionarios_df' not in st.session_state:
    # Chama a função que carrega E processa os tipos/valores
    st.session_state.funcionarios_df = load_data_from_gsheet()
    logging.info(f"DataFrame inicial de avaliações carregado no estado da sessão com {len(st.session_state.funcionarios_df)} linhas.")

# --- Interface Streamlit ---

st.title("Sistema de Avaliação de Funcionários")

# --- Sidebar ---
st.sidebar.title("Menu")
pagina = st.sidebar.radio("Ir para:", ["Nova Avaliação", "Histórico", "Relatórios"])

st.sidebar.markdown("---")
st.sidebar.subheader("Gerenciar Funcionários")

# Formulário para adicionar funcionário na sidebar
with st.sidebar.form("add_funcionario_form", clear_on_submit=True):
    novo_funcionario_nome = st.text_input("Nome do Novo Funcionário")
    submitted = st.form_submit_button("Adicionar Funcionário")
    if submitted:
        if novo_funcionario_nome:
            success = add_funcionario_to_gsheet(novo_funcionario_nome)
            if success:
                st.success(f"Funcionário '{novo_funcionario_nome}' adicionado com sucesso!")
                # Atualiza a lista no estado da sessão imediatamente após adicionar
                st.session_state.lista_funcionarios = load_funcionarios_from_gsheet()
                st.rerun() # Recarrega a página para refletir a mudança na lista
        else:
            st.warning("Por favor, digite o nome do funcionário.")

# --- Página de Nova Avaliação ---
if pagina == "Nova Avaliação":
    st.header("Nova Avaliação de Funcionário")
    with st.form("avaliacao_form", clear_on_submit=True):
        col1, col2, col3 = st.columns([2, 1, 1])

        with col1:
            # Usa a lista do estado da sessão
            lista_func_opts = [""] + st.session_state.get('lista_funcionarios', []) # Pega do estado ou lista vazia
            nome_funcionario = st.selectbox(
                "Selecione o funcionário:",
                options=lista_func_opts,
                index=0,
                format_func=lambda x: "Selecione..." if x == "" else x,
                # help="Selecione o funcionário que está sendo avaliado."
            )
            st.caption("Para adicionar um novo funcionário, use a opção na barra lateral.")

        with col2:
            # Sugere a semana atual, mas permite alteração. Limitado a 1-5 como no código original.
            current_week = datetime.now().isocalendar()[1] # Semana ISO
            # Se a lógica de semana 1-5 for específica, manter o max_value=5
            semana = st.number_input("Semana (1-5):", min_value=1, max_value=5, value=1, step=1) # Ajustar max_value se necessário (e.g., 53)

        with col3:
            # Data da avaliação, padrão para hoje
            data_avaliacao = st.date_input("Data da avaliação:", datetime.now().date())

        st.subheader("Critérios de Avaliação")
        
        # Lista dos critérios excluindo "Quantidade de Retorno" que será tratado separadamente
        criterios_avaliacao = [c for c in CRITERIOS_AVALIACAO_COLS if c != "Quantidade de Retorno"]
        
        st.info("Avalie o funcionário em cada critério como 'Positivo' ou 'Ruim'")

        criterios_descricoes = {
            "Box/Uniforme": "Organização do espaço e uso do uniforme",
            "Ferramentas": "Utilização e cuidado com as ferramentas",
            "EPI": "Uso correto dos EPIs",
            "Horário": "Pontualidade e assiduidade",
            "Apontamento": "Registro e documentação das atividades",
            "Execução de Serviços": "Qualidade e eficiência nas tarefas",
            "Uso de celular/Indisciplinas": "Comportamento apropriado",
            "Quantidade de Retorno": "Quantas vezes um cliente retornou com problemas",
        }

        avaliacoes = {}
        # Divide os critérios em colunas para melhor layout
        num_cols_crit = 2
        cols_crit = st.columns(num_cols_crit)

        # Iteramos sobre os critérios normais (excluindo "Quantidade de Retorno")
        for i, criterio in enumerate(criterios_avaliacao):
            with cols_crit[i % num_cols_crit]:
                st.markdown(f"**{criterio}**") # Usar markdown para negrito
                st.caption(criterios_descricoes.get(criterio, "Sem descrição")) # Descrição abaixo
                # Usar índice 0 ('Positivo') como padrão
                avaliacoes[criterio] = st.radio(
                    f"Avaliação para {criterio}", # Label é importante para acessibilidade
                    options=["Positivo", "Ruim"],
                    key=f"radio_{criterio}", # Chave única para cada radio
                    horizontal=True,
                    index=0, # Padrão para 'Positivo'
                    label_visibility="collapsed" # Esconde o label principal do radio
                )
        
        # Campo específico para Quantidade de Retorno como número inteiro
        st.markdown("**Quantidade de Retorno**")
        st.caption(criterios_descricoes.get("Quantidade de Retorno", "Sem descrição"))
        quantidade_retorno = st.number_input(
            "Quantidade de Retorno", 
            min_value=0, 
            value=0, 
            step=1, 
            help="Informe quantas vezes o cliente retornou com problemas"
        )

        observacoes = st.text_area("Observações adicionais:", height=100)

        submitted_avaliacao = st.form_submit_button("Registrar Avaliação", type="primary")

        if submitted_avaliacao:
            if nome_funcionario and nome_funcionario != "":
                # Monta o dicionário com os dados da avaliação
                nova_avaliacao = {
                    "Nome": nome_funcionario,
                    "Semana": semana,
                    "Data de Avaliação": data_avaliacao, # Já é um objeto date
                    "Observações": observacoes,
                    "Quantidade de Retorno": quantidade_retorno  # Salva o valor inteiro
                }
                # Adiciona as avaliações dos critérios
                # Converte para minúsculo ao salvar para padronizar ('positivo', 'ruim')
                for crit, aval in avaliacoes.items():
                     nova_avaliacao[crit] = aval.lower()

                # Salva no Google Sheets
                if save_to_gsheet(nova_avaliacao):
                    st.success(f"Avaliação de {nome_funcionario} na semana {semana} registrada com sucesso!")
                    # Limpa o cache de dados após salvar com sucesso
                    load_data_from_gsheet.clear()
                    # Atualiza o dataframe no estado da sessão
                    st.session_state.funcionarios_df = load_data_from_gsheet()
                    # Não precisa de rerun aqui, o form já limpa e a próxima interação recarrega se necessário.
                # else: # A função save_to_gsheet já mostra o st.error
                #     st.error("Falha ao salvar a avaliação.")
            else:
                st.error("Por favor, selecione um funcionário antes de registrar a avaliação.")

# --- Página de Histórico (VERSÃO CORRIGIDA COM FILTRO DE DATA) ---
elif pagina == "Histórico":
    st.header("Histórico de Avaliações")

    # Botão para forçar atualização dos dados
    if st.button("🔄 Atualizar Dados", key="update_hist"):
        # Limpa caches e recarrega dados no estado da sessão
        load_funcionarios_from_gsheet.clear()
        load_data_from_gsheet.clear()
        st.session_state.lista_funcionarios = load_funcionarios_from_gsheet()
        st.session_state.funcionarios_df = load_data_from_gsheet()
        st.success("Dados atualizados com sucesso!")
        st.rerun()

    # Verifica se o DataFrame existe e não está vazio no estado da sessão
    if 'funcionarios_df' not in st.session_state or st.session_state.funcionarios_df.empty:
        st.warning("Não há avaliações registradas ou os dados ainda não foram carregados.")
        # Botão para tentar carregar se estiver vazio
        if st.button("Tentar Carregar Dados Agora"):
            st.session_state.funcionarios_df = load_data_from_gsheet()
            st.rerun()
    else:
        # Usa o DataFrame do estado da sessão
        df_original = st.session_state.funcionarios_df.copy()

        # --- Filtros ---
        st.subheader("Filtros")
        col1_filter, col2_filter = st.columns(2)

        with col1_filter:
            # Pega nomes únicos do DataFrame carregado
            nomes_unicos_avaliacoes = sorted(list(df_original["Nome"].unique()))
            nome_filtro = st.multiselect("Funcionário:", nomes_unicos_avaliacoes, default=[], placeholder="Selecione um ou mais")

        with col2_filter:
            # Pega semanas únicas, tratando NAs e convertendo para int
            semanas_unicas = sorted(list(df_original["Semana"].dropna().unique().astype(int)))
            semana_filtro = st.multiselect("Semana (1-5):", semanas_unicas, default=[], placeholder="Selecione uma ou mais")

        # --- Filtro por Data ---
        st.markdown("---") # Separador visual
        col_data1, col_data2 = st.columns(2)

        # Verifica se a coluna de data existe e é do tipo datetime antes de calcular min/max
        min_date_allowed = pd.Timestamp.min.date() # Data mínima global permitida
        max_date_allowed = datetime.now().date()  # Data máxima global permitida
        default_start_date = None # Inicializa como None
        default_end_date = None   # Inicializa como None

        date_column = "Data de Avaliação"
        date_column_exists = date_column in df_original.columns
        date_column_is_datetime = date_column_exists and pd.api.types.is_datetime64_any_dtype(df_original[date_column])

        if date_column_is_datetime:
            valid_dates = df_original[date_column].dropna()
            if not valid_dates.empty:
                min_data_df = valid_dates.min().date()
                max_data_df = valid_dates.max().date()
                default_start_date = min_data_df # Usa min dos dados como padrão inicial
                default_end_date = max_data_df   # Usa max dos dados como padrão final
                min_date_allowed = min_data_df # Ajusta limites permitidos aos dados existentes
                max_date_allowed = max_data_df # Ajusta limites permitidos aos dados existentes
            else:
                 st.caption("Datas válidas não encontradas para definir o intervalo.")
        elif date_column_exists:
             st.caption(f"Coluna '{date_column}' existe, mas não é do tipo data. Filtro de data desabilitado.")
        else:
             st.caption(f"Coluna '{date_column}' não encontrada. Filtro de data desabilitado.")

        # Define valores padrão para date_input mesmo se não houver dados (evita erro)
        if default_start_date is None: default_start_date = datetime.now().date()
        if default_end_date is None: default_end_date = datetime.now().date()


        with col_data1:
            data_inicio = st.date_input(
                "Data de Início:",
                value=default_start_date,
                min_value=min_date_allowed,
                max_value=max_date_allowed,
                disabled=not date_column_is_datetime # Desabilita se a coluna não for datetime
             )
        with col_data2:
             # Garante que a data fim mínima seja a data início selecionada
            min_val_fim = data_inicio if data_inicio else min_date_allowed
            data_fim = st.date_input(
                 "Data de Fim:",
                 value=default_end_date,
                 min_value=min_val_fim,
                 max_value=max_date_allowed,
                 disabled=not date_column_is_datetime # Desabilita se a coluna não for datetime
            )


        # --- Aplicar Filtros ---
        df_filtrado = df_original.copy() # Começa com todos os dados

        # Aplica filtro de nome se houver seleção
        if nome_filtro:
            df_filtrado = df_filtrado[df_filtrado["Nome"].isin(nome_filtro)]

        # Aplica filtro de semana se houver seleção
        if semana_filtro:
             # Converte a coluna para Int64 antes de comparar para lidar com NAs corretamente
            df_filtrado = df_filtrado[df_filtrado["Semana"].astype('Int64').isin(semana_filtro)]

        # Aplicar filtro de data APENAS se a coluna for datetime e o widget estiver habilitado
        if date_column_is_datetime:
            if data_inicio and data_fim: # Verifica se as datas são válidas
                if data_inicio > data_fim:
                    st.error("Erro: A data de início não pode ser posterior à data de fim.")
                else:
                    # Converte st.date_input (date) para Timestamp para comparar com datetime64[ns]
                    start_datetime = pd.to_datetime(data_inicio)
                    end_datetime = pd.to_datetime(data_fim)

                    # Filtra o DataFrame: >= início E <= fim. Exclui NAs na data.
                    df_filtrado = df_filtrado[
                        df_filtrado[date_column].notna() &
                        (df_filtrado[date_column] >= start_datetime) &
                        (df_filtrado[date_column] <= end_datetime)
                    ]
            # else: # Se as datas não forem válidas (None), não filtra por data
                 # logging.debug("Datas de início ou fim inválidas, não aplicando filtro de data.")


        # --- Exibição do DataFrame Filtrado ---
        st.markdown("---")
        st.subheader("Resultados Filtrados")

        if df_filtrado.empty:
            st.warning("Nenhum registro encontrado com os filtros aplicados.")
        else:
            # Preparar DataFrame para exibição (formatar data, mapear valores)
            df_filtrado_display = df_filtrado.copy()

            # Formatar Data para DD/MM/YYYY apenas para exibição
            if date_column_is_datetime:
                df_filtrado_display[date_column] = df_filtrado_display[date_column].dt.strftime('%d/%m/%Y').fillna('N/A')
            elif date_column_exists: # Se existe mas não é datetime, apenas converte para string
                 df_filtrado_display[date_column] = df_filtrado_display[date_column].astype(str).fillna('N/A')
            else: # Se nem existe
                 df_filtrado_display[date_column] = 'N/A'

            # Mapear 'positivo'/'ruim' para 'Positivo'/'Ruim' para exibição
            map_dict_display = {'positivo': 'Positivo', 'ruim': 'Ruim'}
            for criterio in CRITERIOS_AVALIACAO_COLS:
                if criterio in df_filtrado_display.columns:
                    # Aplica o mapeamento, mantendo outros valores como estão (astype('string') lida com NA)
                    df_filtrado_display[criterio] = df_filtrado_display[criterio].map(map_dict_display).fillna(df_filtrado_display[criterio])

            # Substitui NAs restantes por string vazia para exibição mais limpa
            df_filtrado_display.fillna('N/A', inplace=True)

            st.dataframe(df_filtrado_display, use_container_width=True, hide_index=True)

            # --- Botões de Download (usar df_filtrado original ANTES da formatação de display) ---
            st.markdown("---")
            st.subheader("Download dos Dados Filtrados")
            col_dl1, col_dl2 = st.columns(2)

            with col_dl1:
                # Usa o df_filtrado (com dados mais brutos) para o CSV
                try:
                    csv_data = save_data_csv(df_filtrado) # Função já formata datas e valores para CSV
                    st.download_button(
                        label="Download CSV",
                        data=csv_data,
                        file_name=f"avaliacao_filtrada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv"
                    )
                except Exception as e:
                     st.error(f"Erro ao gerar CSV: {e}")

            with col_dl2:
                try:
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        # Preparar dados para Excel (similar ao display, mas talvez com menos N/A)
                        df_export_excel = df_filtrado.copy()

                        # Formatar data para DD/MM/YYYY no Excel
                        if date_column_is_datetime:
                           df_export_excel[date_column] = df_export_excel[date_column].dt.strftime('%d/%m/%Y')
                        elif date_column_exists:
                           df_export_excel[date_column] = df_export_excel[date_column].astype(str)

                        # Mapear para 'Positivo'/'Ruim' no Excel
                        map_dict_excel = {'positivo': 'Positivo', 'ruim': 'Ruim'}
                        for criterio in CRITERIOS_AVALIACAO_COLS:
                            if criterio in df_export_excel.columns:
                                 df_export_excel[criterio] = df_export_excel[criterio].map(map_dict_excel).fillna(df_export_excel[criterio])

                        # Substituir NA/None por string vazia para Excel
                        df_export_excel.fillna('', inplace=True)

                        df_export_excel.to_excel(writer, index=False, sheet_name='Avaliações Filtradas')
                        # Auto-ajuste de colunas (opcional)
                        workbook = writer.book
                        worksheet = writer.sheets['Avaliações Filtradas']
                        for i, col in enumerate(df_export_excel.columns):
                             try: # Adiciona try-except para evitar erro em coluna vazia
                                 column_len = max(
                                     df_export_excel[col].astype(str).map(len).max(),
                                     len(col)
                                 ) + 2
                                 worksheet.set_column(i, i, min(column_len, 50)) # Limita largura máx
                             except Exception:
                                  worksheet.set_column(i, i, len(col) + 2) # Usa largura do cabeçalho se erro

                    st.download_button(
                        label="Download Excel",
                        data=buffer.getvalue(),
                        file_name=f"avaliacao_filtrada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Erro ao gerar Excel: {e}")


# --- Página de Relatórios (COM CORREÇÃO DO KEYERROR MANTIDA) ---
elif pagina == "Relatórios":
    st.header("Relatórios e Análises de Desempenho")

    # Botão para atualizar dados especificamente nesta página
    if st.button("🔄 Atualizar Dados para Relatórios", key="update_rel"):
        load_data_from_gsheet.clear()
        st.session_state.funcionarios_df = load_data_from_gsheet() # Recarrega e processa
        st.success("Dados atualizados com sucesso para relatórios!")
        st.rerun()

    # Verifica se há dados no estado da sessão
    if 'funcionarios_df' not in st.session_state or st.session_state.funcionarios_df.empty:
        st.warning("Não há dados para gerar relatórios. Registre avaliações ou clique em 'Atualizar dados'.")
        if st.button("Tentar Carregar Dados Agora", key="load_rel"):
             st.session_state.funcionarios_df = load_data_from_gsheet()
             st.rerun()
    else:
        df_relatorio_base = st.session_state.funcionarios_df.copy()

        # Filtro por funcionário para o relatório
        nomes_unicos_avaliacoes = sorted(["Todos"] + list(df_relatorio_base["Nome"].unique()))
        nome_relatorio = st.selectbox("Selecione o funcionário para análise:", nomes_unicos_avaliacoes)

        # Filtra pelo funcionário selecionado (ou usa todos)
        if nome_relatorio != "Todos":
            df_relatorio_filtrado = df_relatorio_base[df_relatorio_base["Nome"] == nome_relatorio].copy()
        else:
            df_relatorio_filtrado = df_relatorio_base.copy()

        if df_relatorio_filtrado.empty:
            st.warning(f"Nenhuma avaliação encontrada para '{nome_relatorio}'.")
        else:
            # --- Análise das Últimas Semanas ---
            # Garante que 'Semana' seja numérica Int64
            df_relatorio_filtrado['Semana'] = pd.to_numeric(df_relatorio_filtrado['Semana'], errors='coerce').astype('Int64')
            semanas_disponiveis = sorted(df_relatorio_filtrado['Semana'].dropna().unique())

            if not semanas_disponiveis:
                st.warning(f"Nenhuma semana de avaliação encontrada para '{nome_relatorio}'.")
            else:
                # Define o período de análise (ex: últimas 5 semanas com dados)
                num_semanas_analise = 5
                semanas_recentes = semanas_disponiveis[-num_semanas_analise:] # Pega as últimas N semanas
                df_last_weeks = df_relatorio_filtrado[df_relatorio_filtrado['Semana'].isin(semanas_recentes)].copy()

                if df_last_weeks.empty:
                     st.warning(f"Nenhum dado encontrado nas últimas {num_semanas_analise} semanas de avaliação ({min(semanas_recentes)}-{max(semanas_recentes)}) para '{nome_relatorio}'.")
                else:
                    num_semanas_analisadas = len(semanas_recentes)
                    st.info(f"Análise focada nas últimas {num_semanas_analisadas} semanas de avaliação disponíveis (Semanas: {min(semanas_recentes)} a {max(semanas_recentes)}) para '{nome_relatorio}'.")

                    # --- Cálculos Agregados ---
                    contagem = {}
                    for criterio in CRITERIOS_AVALIACAO_COLS:
                        if criterio in df_last_weeks.columns:
                            # value_counts() já ignora NA por padrão
                            counts = df_last_weeks[criterio].value_counts()
                            contagem[criterio] = {
                                'Positivo': counts.get('positivo', 0), # Usa chave minúscula
                                'Ruim': counts.get('ruim', 0)          # Usa chave minúscula
                            }
                        else:
                            contagem[criterio] = {'Positivo': 0, 'Ruim': 0}

                    df_contagem = pd.DataFrame(contagem).T
                    # Garante que as colunas existam mesmo que não haja dados
                    if 'Positivo' not in df_contagem.columns: df_contagem['Positivo'] = 0
                    if 'Ruim' not in df_contagem.columns: df_contagem['Ruim'] = 0

                    # Calcula o total de avaliações por critério (Positivo + Ruim)
                    df_contagem['Total'] = df_contagem['Positivo'] + df_contagem['Ruim']

                    # Calcula totais gerais
                    total_positivos_geral = df_contagem['Positivo'].sum()
                    total_ruins_geral = df_contagem['Ruim'].sum()
                    total_avaliacoes_geral = df_contagem['Total'].sum()

                    # Calcula percentual geral (evita divisão por zero)
                    percentual_geral = (total_positivos_geral / total_avaliacoes_geral * 100) if total_avaliacoes_geral > 0 else 0
                    total_registros_periodo = len(df_last_weeks) # Nº de linhas no período

                    st.markdown("---")
                    st.subheader(f"Resumo do Desempenho - {nome_relatorio} (Semanas {min(semanas_recentes)}-{max(semanas_recentes)})")

                    col_resumo1, col_resumo2 = st.columns(2)

                    # --- COLUNA 1: RESUMO GERAL COM MÉTRICAS ---
                    with col_resumo1:
                        st.write(f"**Desempenho Geral (Últimas {num_semanas_analisadas} Semanas)**")
                        if total_avaliacoes_geral > 0:
                            st.metric("Percentual Positivo Geral", f"{percentual_geral:.1f}%")

                            # Define o status com base no percentual
                            if percentual_geral >= 80:
                                status = "✅ Excelente"
                                help_text = "Acima de 80% de avaliações positivas."
                            elif percentual_geral >= 60:
                                status = "👍 Bom"
                                help_text = "Entre 60% e 79.9% de avaliações positivas."
                            elif percentual_geral >= 40:
                                status = "⚠️ Atenção"
                                help_text = "Entre 40% e 59.9% de avaliações positivas. Requer acompanhamento."
                            else:
                                status = "❌ Crítico"
                                help_text = "Abaixo de 40% de avaliações positivas. Requer ação."
                            st.metric("Status Geral", status, help=help_text)

                            st.metric("Total de Critérios Avaliados", f"{total_avaliacoes_geral} (Pos: {total_positivos_geral}, Ruim: {total_ruins_geral})")
                            st.metric("Nº de Registros de Avaliação", f"{total_registros_periodo}")
                        else:
                            st.info("Não há avaliações ('Positivo' ou 'Ruim') registradas no período para calcular o desempenho.")

                    # --- COLUNA 2: GRÁFICO DE TENDÊNCIA ---
                    with col_resumo2:
                        st.write("**Tendência de Desempenho Semanal**")
                        try:
                             # Derrete o DF para facilitar o agrupamento por semana e critério
                            df_melted = df_last_weeks.melt(
                                id_vars=['Semana'],
                                value_vars=CRITERIOS_AVALIACAO_COLS,
                                var_name='Criterio',
                                value_name='Avaliacao'
                            )
                            # Filtra apenas avaliações válidas ('positivo' ou 'ruim')
                            df_melted_valid = df_melted[df_melted['Avaliacao'].isin(['positivo', 'ruim'])]

                            # Agrupa por semana e calcula total e positivos
                            df_weekly_summary = df_melted_valid.groupby('Semana')['Avaliacao'].agg(
                                total= 'count',
                                positivo=lambda x: (x == 'positivo').sum()
                            ).reset_index()

                            # Calcula percentual positivo, tratando divisão por zero
                            df_weekly_summary['Percent_Positivo'] = (
                                df_weekly_summary['positivo'] / df_weekly_summary['total'].replace(0, pd.NA) * 100
                            ).fillna(0) # Preenche NA com 0 se o total for 0

                            if not df_weekly_summary.empty:
                                if len(df_weekly_summary) > 1 : # Precisa de pelo menos 2 pontos para uma linha
                                    fig_line, ax_line = plt.subplots(figsize=(7, 4)) # Ajusta tamanho
                                    ax_line.plot(df_weekly_summary['Semana'].astype(str), df_weekly_summary['Percent_Positivo'], marker='o', linestyle='-', color='royalblue') # Usa semana como string para eixo X
                                    ax_line.set_title("% Positivas por Semana", fontsize=11)
                                    ax_line.set_xlabel("Semana do Ano", fontsize=9)
                                    ax_line.set_ylabel("% Positivo", fontsize=9)
                                    ax_line.set_ylim(0, 105) # De 0 a 105 para dar espaço ao 100%
                                    ax_line.grid(True, axis='y', linestyle=':', alpha=0.7)
                                    ax_line.tick_params(axis='both', which='major', labelsize=8)
                                    # ax_line.set_xticks(df_weekly_summary['Semana'].unique()) # Pode ficar lotado, deixar automático
                                    plt.tight_layout()
                                    st.pyplot(fig_line)
                                elif len(df_weekly_summary) == 1:
                                     # Mostra métrica se houver apenas uma semana
                                     semana_unica = df_weekly_summary['Semana'].iloc[0]
                                     perc_unico = df_weekly_summary['Percent_Positivo'].iloc[0]
                                     st.metric(f"Semana {semana_unica}", f"{perc_unico:.1f}% Positivo")
                                else: # Caso df_weekly_summary esteja vazio após filtros
                                     st.info("Não há dados semanais suficientes para gerar o gráfico de tendência.")

                            else:
                                st.info("Não há avaliações ('Positivo' ou 'Ruim') nas semanas selecionadas para o gráfico.")

                        except Exception as e:
                            logging.error(f"Erro ao gerar gráfico de tendência: {e}", exc_info=True)
                            st.error("Erro ao gerar o gráfico de tendência semanal.")

                    st.markdown("---")
                    st.subheader("Desempenho por Critério")

                    # Mostra a tabela de contagem
                    df_contagem_display = df_contagem[['Positivo', 'Ruim', 'Total']].copy()
                    # Calcula percentual por critério
                    df_contagem_display['% Positivo'] = (
                        df_contagem_display['Positivo'] / df_contagem_display['Total'].replace(0, pd.NA) * 100
                    ).fillna(0).round(1) # Arredonda para 1 casa decimal
                    st.dataframe(df_contagem_display.sort_values(by='% Positivo', ascending=True), use_container_width=True)

                    # Gráfico de Barras - Pontos fortes e fracos
                    if not df_contagem_display.empty:
                        try:
                            fig_bar, ax_bar = plt.subplots(figsize=(10, max(5, len(df_contagem_display)*0.5))) # Altura dinâmica
                            df_contagem_display.sort_values(by='% Positivo', ascending=True, inplace=True) # Ordena para o gráfico
                            ax_bar.barh(df_contagem_display.index, df_contagem_display['% Positivo'], color='skyblue')
                            ax_bar.set_xlabel('% Positivo')
                            ax_bar.set_title('Percentual Positivo por Critério de Avaliação')
                            ax_bar.set_xlim(0, 105)
                            ax_bar.set_xlabel("")
                            ax_bar.xaxis.set_visible(False)
                            for spine in ax_bar.spines.values():
                                 spine.set_visible(False)
                            # Adiciona os valores nas barras
                            for index, value in enumerate(df_contagem_display['% Positivo']):
                                 ax_bar.text(value + 1, index, f'{value:.1f}%', va='center', fontsize=9)
                            plt.tight_layout()
                            st.pyplot(fig_bar)
                        except Exception as e:
                             logging.error(f"Erro ao gerar gráfico de barras: {e}", exc_info=True)
                             st.error("Erro ao gerar o gráfico de desempenho por critério.")

                    # Exibir Observações Relevantes (ex: onde a avaliação foi 'ruim')
                    st.markdown("---")
                    st.subheader("Observações Relevantes (Avaliações 'Ruim')")
                    df_ruins = df_last_weeks # Começa com os dados das últimas semanas
                    # Filtra linhas onde PELO MENOS UM critério foi 'ruim'
                    mask_ruim = df_ruins[CRITERIOS_AVALIACAO_COLS].eq('ruim').any(axis=1)
                    df_obs_ruim = df_ruins.loc[mask_ruim, ['Nome', 'Semana', 'Data de Avaliação', 'Observações'] + CRITERIOS_AVALIACAO_COLS]
                    date_column_is_datetime = False
                    if not df_obs_ruim.empty:
                        # Formatar data para exibição
                        if date_column_is_datetime: # Usa a flag já definida
                             df_obs_ruim['Data de Avaliação'] = df_obs_ruim['Data de Avaliação'].dt.strftime('%d/%m/%Y').fillna('N/A')
                        else:
                              df_obs_ruim['Data de Avaliação'] = df_obs_ruim['Data de Avaliação'].astype(str).fillna('N/A')

                        # Destaca quais critérios foram ruins na linha
                        def highlight_ruim(row):
                            highlighted_obs = f"Obs: {row['Observações'] or 'Nenhuma'}. \nCritérios Ruins: "
                            ruins = [col for col in CRITERIOS_AVALIACAO_COLS if row[col] == 'ruim']
                            return highlighted_obs + (", ".join(ruins) if ruins else "Nenhum")

                        df_obs_ruim['Detalhes'] = df_obs_ruim.apply(highlight_ruim, axis=1)
                        st.dataframe(df_obs_ruim[['Nome', 'Semana', 'Data de Avaliação', 'Detalhes']], hide_index=True, use_container_width=True)
                    else:
                        st.info("Nenhuma observação associada a avaliações 'Ruim' encontrada no período selecionado.")

# --- Fim do Script ---