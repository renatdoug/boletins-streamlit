import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import os
import re
import json
import traceback

# Constantes
SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
SHEET_NAME = "Boletins"
WORKSHEET_NOTAS = "Notas_Tabela"
WORKSHEET_CONTROLE = "Controle_Liberacao"

# Funções auxiliares


def authenticate_gsheets():
    """Autentica com Google Sheets usando credenciais JSON."""
    try:
        # Localmente, tenta carregar credenciais.json
        credentials_path = os.path.join(
            os.path.dirname(__file__), "credenciais.json")
        if os.path.exists(credentials_path):
            with open(credentials_path, "r", encoding="utf-8") as f:
                credentials_info = json.load(f)
            credentials = Credentials.from_service_account_info(
                credentials_info, scopes=SCOPE)
        # No Streamlit Cloud, usa st.secrets
        elif "GOOGLE_CREDENTIALS" in st.secrets:
            credentials_info = st.secrets["GOOGLE_CREDENTIALS"]
            credentials = Credentials.from_service_account_info(
                credentials_info, scopes=SCOPE)
        else:
            st.error(
                "Credenciais do Google Sheets não encontradas. "
                "Verifique se 'credenciais.json' está no diretório do projeto "
                "ou se 'GOOGLE_CREDENTIALS' está configurado no secrets.toml do Streamlit Cloud."
            )
            st.stop()
        return gspread.authorize(credentials)
    except json.JSONDecodeError as e:
        st.error(
            f"Erro ao ler o arquivo credenciais.json: Formato JSON inválido. {e}")
        st.stop()
    except Exception as e:
        st.error(
            f"Erro ao autenticar com Google Sheets: {e}\n{traceback.format_exc()}")
        st.stop()


def clean_nota_value(value):
    """Converte valores de nota, tratando vírgulas, datas e outros formatos."""
    if pd.isna(value):
        return 0.0
    value = str(value).strip()
    value = value.replace(',', '.')
    if re.match(r'^\d{1,2}/\d{1,2}$', value):
        try:
            parts = value.split('/')
            value = f"{parts[0]}.{parts[1]}"
        except:
            return 0.0
    value = re.sub(r'[^\d.]', '', value)
    parts = value.split('.')
    if len(parts) > 2:
        value = parts[0] + '.' + ''.join(parts[1:])
    try:
        return float(value) if value else 0.0
    except ValueError:
        return 0.0


@st.cache_data(show_spinner=False, ttl=300)
def load_data(worksheet_name, _cache_version=0):
    """Carrega dados de uma planilha como DataFrame."""
    try:
        client = st.session_state["client"]
        sheet = client.open(SHEET_NAME).worksheet(worksheet_name)
        data = pd.DataFrame(sheet.get_all_records())
        headers = sheet.row_values(1)
        # Normalizar colunas de texto
        for col in ['Matrícula', 'Série', 'Componente Curricular', 'Bimestre', 'Tipo de Avaliação', 'Mat_Professor']:
            if col in data.columns:
                data[col] = data[col].astype(str).str.strip().str.upper()
        # Converte a coluna 'Nota'
        if 'Nota' in data.columns:
            data['Nota'] = data['Nota'].apply(clean_nota_value)
            data['Nota'] = pd.to_numeric(
                data['Nota'], errors='coerce').fillna(0.0)
        # Adiciona índice da linha (1-based, considerando cabeçalho)
        data['row_index'] = data.index + 2
        return data, sheet, headers
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Planilha {worksheet_name} não encontrada.")
        return pd.DataFrame(), None, []
    except Exception as e:
        st.error(
            f"Erro ao carregar planilha {worksheet_name}: {e}\n{traceback.format_exc()}")
        st.stop()


def validate_period(bimestre, df_periodo, today):
    """Valida se o período de lançamento está liberado."""
    bimestre = str(bimestre).strip().upper()
    periodo_ok = df_periodo[df_periodo['Bimestre'].str.strip(
    ).str.upper() == bimestre]
    if periodo_ok.empty:
        return False, "Lançamento não autorizado para este período. Consulte o gestor."
    try:
        inicio = datetime.strptime(
            periodo_ok['Data Início'].values[0], "%d/%m/%Y").date()
        fim = datetime.strptime(
            periodo_ok['Data Fim'].values[0], "%d/%m/%Y").date()
        if not (inicio <= today <= fim):
            return False, f"Lançamento permitido apenas entre {inicio.strftime('%d/%m/%Y')} e {fim.strftime('%d/%m/%Y')}"
        return True, ""
    except ValueError as e:
        return False, f"Erro no formato das datas: {e}"


def validate_professor(mat_prof, df):
    """Verifica se a matrícula do professor é válida."""
    return str(mat_prof).strip().upper() in df['Mat_Professor'].str.strip().str.upper().values


def logout():
    """Limpa a autenticação do professor e parâmetros."""
    for key in list(st.session_state.keys()):
        if key not in ["client", "df", "sheet_notas", "df_periodo", "headers_notas", "cache_version"]:
            del st.session_state[key]
    st.success("Deslogado com sucesso!")
    st.rerun()


# Inicialização
if "client" not in st.session_state:
    st.session_state["client"] = authenticate_gsheets()
if "cache_version" not in st.session_state:
    st.session_state["cache_version"] = 0
if "df" not in st.session_state:
    st.session_state["df"], st.session_state["sheet_notas"], st.session_state["headers_notas"] = load_data(
        WORKSHEET_NOTAS, _cache_version=st.session_state["cache_version"])
    st.session_state["df_periodo"], _, _ = load_data(
        WORKSHEET_CONTROLE, _cache_version=st.session_state["cache_version"])

client = st.session_state["client"]
df = st.session_state["df"]
sheet_notas = st.session_state["sheet_notas"]
df_periodo = st.session_state["df_periodo"]
headers_notas = st.session_state["headers_notas"]

# Encontra a coluna 'Nota'
nota_column_idx = headers_notas.index(
    "Nota") + 1 if "Nota" in headers_notas else 8
nota_column_letter = chr(64 + nota_column_idx)

# Interface
st.title("📘 Lançamento de Notas por Professor 2025")

# Botão de logout
if st.session_state.get("prof_autenticado"):
    if st.button("Deslogar"):
        logout()

# 1. Autenticação do Professor
if not st.session_state.get("prof_autenticado"):
    with st.form("auth_form"):
        st.subheader("1. Identificação do Professor")
        nome_prof = st.text_input("Nome do Professor")
        mat_prof = st.text_input("Matrícula do Professor")
        submit_prof = st.form_submit_button("Confirmar")

    if submit_prof:
        if not nome_prof.strip() or not mat_prof.strip():
            st.error("Por favor, preencha nome e matrícula com valores válidos.")
        elif not validate_professor(mat_prof, df):
            st.error("Matrícula inválida ou sem permissão.")
        else:
            st.session_state["prof_autenticado"] = True
            st.session_state["nome_prof"] = nome_prof
            st.session_state["mat_prof"] = mat_prof
            st.success("Autenticação realizada com sucesso!")
else:
    # 2. Parâmetros do Lançamento
    nome_prof = st.session_state["nome_prof"]
    mat_prof = st.session_state["mat_prof"]

    st.subheader("2. Parâmetros do Lançamento")
    series_disponiveis = df[df['Mat_Professor'].str.strip().str.upper(
    ) == mat_prof.strip().upper()]['Série'].unique().tolist()
    if not series_disponiveis:
        st.error("Nenhuma série associada a esta matrícula.")
        st.stop()

    serie = st.selectbox(
        "Série", options=[""] + series_disponiveis, index=0, key="serie")
    componentes = df[(df['Mat_Professor'].str.strip().str.upper() == mat_prof.strip()) &
                     (df['Série'].str.strip().str.upper() == str(serie).strip().upper())]['Componente Curricular'].unique() if serie else []
    componente = st.selectbox("Componente Curricular", options=[
                              ""] + list(componentes) if len(componentes) > 0 else [""], index=0, key="componente")
    bimestre = st.selectbox(
        "Bimestre/Período", options=["", "1º", "2º", "3º", "4º", "Final"], index=0, key="bimestre")
    tipo_avaliacao = st.selectbox("Tipo de Avaliação", options=[
                                  "", "MENSAL", "BIMESTRAL", "RECUPERAÇÃO"], index=0, key="tipo_avaliacao")

    if not serie or not componente or not bimestre or not tipo_avaliacao:
        st.warning("Por favor, selecione todos os parâmetros de lançamento.")
        st.stop()

    if componente == "":
        st.warning("Nenhum componente curricular disponível para esta série.")
        st.stop()

    # Validação do período
    hoje = datetime.today().date()
    periodo_valido, mensagem = validate_period(bimestre, df_periodo, hoje)
    if not periodo_valido:
        st.error(f"❌ {mensagem}")
        st.stop()

    # Carrega alunos
    alunos_serie = df[(df['Série'].str.strip().str.upper() == str(serie).strip().upper())][
        ['Nome do Aluno', 'Matrícula', 'Turno']].drop_duplicates(subset=['Matrícula']).sort_values(by='Nome do Aluno')

    if alunos_serie.empty:
        st.warning("Nenhum aluno encontrado para esta série.")
        st.stop()

    # 3. Lançamento de Notas
    st.subheader("3. Lançamento de Notas")
    with st.form("form_lote_notas"):
        notas = {}
        for idx, row in alunos_serie.iterrows():
            nome = row['Nome do Aluno']
            matricula = row['Matrícula']
            col_id = f"nota_{matricula}_{serie}_{componente}_{bimestre}_{tipo_avaliacao}_{idx}"

            # Normalizar valores para o filtro
            matricula_norm = str(matricula).strip().upper()
            serie_norm = str(serie).strip().upper()
            componente_norm = str(componente).strip().upper()
            bimestre_norm = str(bimestre).strip().upper()
            tipo_avaliacao_norm = str(tipo_avaliacao).strip().upper()

            # Busca nota existente
            cond = (
                (df['Matrícula'].str.strip().str.upper() == matricula_norm) &
                (df['Série'].str.strip().str.upper() == serie_norm) &
                (df['Componente Curricular'].str.strip().str.upper() == componente_norm) &
                (df['Bimestre'].str.strip().str.upper() == bimestre_norm) &
                (df['Tipo de Avaliação'].str.strip(
                ).str.upper() == tipo_avaliacao_norm)
            )
            nota_existente = float(
                df[cond]['Nota'].values[0]) if not df[cond].empty else 0.0

            cols = st.columns([3, 1])
            cols[0].markdown(f"**{nome} ({matricula})**")
            notas[matricula] = cols[1].number_input(
                "", min_value=0.0, max_value=10.0, step=0.1, value=nota_existente, key=col_id)

        sobrescrever = st.checkbox(
            "🔁 Sobrescrever notas existentes", key="sobrescrever")
        submitted = st.form_submit_button("Salvar Notas")

        if submitted:
            with st.spinner("Salvando notas..."):
                registros = []
                erros = []
                atualizados = []
                batch_updates = []

                for _, row in alunos_serie.iterrows():
                    nome = row['Nome do Aluno']
                    matricula = row['Matrícula']
                    turno = row['Turno']
                    nota_valor = notas[matricula]

                    # Normalizar valores para o filtro
                    matricula_norm = str(matricula).strip().upper()
                    serie_norm = str(serie).strip().upper()
                    componente_norm = str(componente).strip().upper()
                    bimestre_norm = str(bimestre).strip().upper()
                    tipo_avaliacao_norm = str(tipo_avaliacao).strip().upper()

                    # Busca nota existente
                    cond = (
                        (df['Matrícula'].str.strip().str.upper() == matricula_norm) &
                        (df['Série'].str.strip().str.upper() == serie_norm) &
                        (df['Componente Curricular'].str.strip().str.upper() == componente_norm) &
                        (df['Bimestre'].str.strip().str.upper() == bimestre_norm) &
                        (df['Tipo de Avaliação'].str.strip().str.upper() == tipo_avaliacao_norm)
                    )
                    nota_existente = float(df[cond]['Nota'].values[0]) if not df[cond].empty else 0.0

                    # Ignorar se a nota não foi alterada
                    if nota_valor == nota_existente or (nota_valor == 0.0 and nota_existente == 0.0):
                        continue

                    nova_linha = [
                        nome, matricula, serie, turno, componente,
                        bimestre, tipo_avaliacao, f"{nota_valor:.2f}", nome_prof, mat_prof
                    ]

                    existe = df[cond]
                    if not existe.empty:
                        if sobrescrever:
                            try:
                                row_idx = existe['row_index'].values[0]
                                batch_updates.append({
                                    "range": f"{nota_column_letter}{row_idx}",
                                    "values": [[f"{nota_valor:.2f}"]]
                                })
                                atualizados.append(
                                    f"Nota atualizada para {nome} ({matricula}): {nota_valor:.2f}")
                            except Exception as e:
                                erros.append(f"Erro ao atualizar nota para {nome} ({matricula}): {e}")
                        else:
                            erros.append(
                                f"Nota já existe para {nome} ({matricula}). Marque 'Sobrescrever' para atualizar.")
                    else:
                        registros.append(nova_linha)

                # Reinicializar o cliente antes de atualizar
                try:
                    client = authenticate_gsheets()
                    st.session_state["client"] = client
                    sheet_notas = client.open(SHEET_NAME).worksheet(WORKSHEET_NOTAS)
                    st.session_state["sheet_notas"] = sheet_notas
                except Exception as e:
                    st.error(f"Erro ao reinicializar autenticação: {e}\n{traceback.format_exc()}")
                    st.stop()

                # Executar atualizações em lote
                if batch_updates:
                    try:
                        sheet_notas.batch_update(batch_updates)
                        for msg in atualizados:
                            st.success(msg)
                    except Exception as e:
                        st.error(f"Erro ao atualizar notas em lote: {e}\n{traceback.format_exc()}")
                        st.stop()

                # Adicionar novos registros
                if registros:
                    try:
                        sheet_notas.append_rows(registros, value_input_option="USER_ENTERED")
                        for reg in registros:
                            st.success(f"Nota lançada para {reg[0]} ({reg[1]}): {reg[7]}")
                    except Exception as e:
                        st.error(f"Erro ao adicionar novas notas: {e}\n{traceback.format_exc()}")
                        st.stop()

                # Exibir erros, se houver
                for erro in erros:
                    st.error(erro)

                # Verificar se houve alterações
                if not atualizados and not registros:
                    st.info("Nenhuma nota foi alterada ou adicionada.")
                else:
                    # Atualizar cache
                    st.session_state["cache_version"] += 1
                    st.session_state["df"], st.session_state["sheet_notas"], st.session_state["headers_notas"] = load_data(
                        WORKSHEET_NOTAS, _cache_version=st.session_state["cache_version"])
                    st.success("Notas processadas com sucesso!")