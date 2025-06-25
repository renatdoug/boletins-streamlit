import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import re
import traceback
import os
import json

# Constantes
SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
SHEET_NAME = "Boletins"
WORKSHEET_NOTAS = "Notas_Tabela"

# Funções auxiliares


def authenticate_gsheets():
    try:
        if os.path.exists("credenciais.json"):
            credentials = Credentials.from_service_account_file(
                "credenciais.json", scopes=SCOPE)
        elif "GOOGLE_CREDENTIALS" in st.secrets:
            credentials_info = st.secrets["google_credentials"]
            credentials = Credentials.from_service_account_info(
                credentials_info, scopes=SCOPE)
        else:
            st.error("Credenciais do Google Sheets não encontradas.")
            st.stop()
        return gspread.authorize(credentials)
    except Exception as e:
        st.error(f"Erro ao autenticar com Google Sheets: {e}")
        st.stop()


client = authenticate_gsheets()


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
    return float(value) if value else 0.0


@st.cache_data(show_spinner=False, ttl=300)
def load_data(worksheet_name, _cache_version=0):
    """Carrega dados da planilha."""
    try:
        client = st.session_state["client"]
        sheet = client.open(SHEET_NAME).worksheet(worksheet_name)
        df = pd.DataFrame(sheet.get_all_records())
        if df.empty:
            st.error("Planilha vazia.")
            st.stop()
        required_cols = ['Série', 'Nome do Aluno', 'Matrícula', 'Bimestre',
                         'Componente Curricular', 'Tipo de Avaliação', 'Nota']
        if not all(col in df.columns for col in required_cols):
            st.error("Colunas obrigatórias ausentes na planilha.")
            st.stop()
        # Normalizar colunas de texto
        for col in required_cols[:-1]:  # Exceto 'Nota'
            df[col] = df[col].astype(str).str.strip().str.upper()
        df['Nota'] = df['Nota'].apply(clean_nota_value)
        return df
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Planilha {worksheet_name} não encontrada.")
        st.stop()
    except Exception as e:
        st.error(f"Erro ao acessar planilha: {e}\n{traceback.format_exc()}")
        st.stop()


def validate_matricula(nome, matricula, alunos_serie):
    """Valida a matrícula do aluno."""
    return not alunos_serie[
        (alunos_serie['Nome do Aluno'].str.upper() == nome.upper()) &
        (alunos_serie['Matrícula'].astype(
            str).str.strip() == matricula.strip())
    ].empty


def calculate_media(resultado):
    """Calcula a média entre MENSAL e BIMESTRAL para cada componente curricular."""
    medias = {}
    mensal_rows = resultado[resultado['Tipo de Avaliação'] == 'MENSAL']
    bimestral_rows = resultado[resultado['Tipo de Avaliação'] == 'BIMESTRAL']

    for comp in resultado['Componente Curricular'].unique():
        mensal = mensal_rows[mensal_rows['Componente Curricular'] ==
                             comp]['Nota'].iloc[0] if not mensal_rows[mensal_rows['Componente Curricular'] == comp].empty else 0.0
        bimestral = bimestral_rows[bimestral_rows['Componente Curricular'] ==
                                   comp]['Nota'].iloc[0] if not bimestral_rows[bimestral_rows['Componente Curricular'] == comp].empty else 0.0
        if mensal > 0.0 or bimestral > 0.0:
            medias[comp] = (mensal + bimestral) / 2
        else:
            medias[comp] = 0.0
    return medias


def check_recuperacao(medias):
    """Verifica se recuperação é necessária para médias < 8."""
    recuperacao_needed = []
    for comp, media in medias.items():
        if media < 8:
            recuperacao_needed.append(f"{comp} (Média: {media:.2f})")
    return recuperacao_needed


def check_recuperacao_final(resultado, medias):
    """Verifica o resultado da recuperação apenas para componentes com média < 8."""
    recuperacao_rows = resultado[resultado['Tipo de Avaliação']
                                 == 'RECUPERAÇÃO']
    resultados = []
    # Considera apenas componentes que precisaram de recuperação (média < 8)
    componentes_recuperacao = [comp for comp,
                               media in medias.items() if media < 8]
    for comp in componentes_recuperacao:
        nota_rec = recuperacao_rows[recuperacao_rows['Componente Curricular'] ==
                                    comp]['Nota'].iloc[0] if not recuperacao_rows[recuperacao_rows['Componente Curricular'] == comp].empty else 0.0
        status = "Aprovado" if nota_rec >= 8 else "Reprovado"
        resultados.append(
            f"{comp} (Nota de Recuperação: {nota_rec:.2f} - {status})")
    return resultados


def display_boletim(resultado):
    """Exibe o boletim com estilização, cálculo de média e mensagens de recuperação."""
    desired_order = ['MENSAL', 'BIMESTRAL',
                     'MEDIA', 'RECUPERAÇÃO', 'RECUPERAÇÃO FINAL']
    available_types = resultado['Tipo de Avaliação'].unique()
    ordered_types = [t for t in desired_order if t in available_types]

    boletim = (
        resultado.pivot_table(
            index='Componente Curricular',
            columns='Tipo de Avaliação',
            values='Nota',
            aggfunc='first'
        )
        .reindex(columns=ordered_types)
        .reset_index()
    )
    boletim.columns.name = None
    boletim = boletim.rename(columns={
        "MENSAL": "Men",
        "BIMESTRAL": "Bim",
        "MEDIA": "Med",
        "RECUPERAÇÃO": "Rec",
        "RECUPERAÇÃO FINAL": "Rec Final"
    })

    # Calcular médias
    medias = calculate_media(resultado)
    boletim['Med'] = [medias.get(comp, 0.0)
                      for comp in boletim['Componente Curricular']]

    def colorir_nota(val):
        if isinstance(val, (int, float)):
            return 'background-color: #ffd6d6; color: black; font-weight: bold; text-align: center' if val < 7 else 'background-color: #d6ecff; color: black; font-weight: bold; text-align: center'
        return ''

    st.success("Notas encontradas:")
    st.dataframe(
        boletim.style
        .applymap(colorir_nota, subset=boletim.columns[1:])
        .format("{:.2f}", subset=boletim.columns[1:], na_rep="-")
    )

    # Mensagens de recuperação necessária
    recuperacao_needed = check_recuperacao(medias)
    if recuperacao_needed:
        st.warning("Recuperação necessária para: " +
                   ", ".join(recuperacao_needed))
        st.warning("O ALUNO DEVERÁ FAZER PROVA DE RECUPERAÇÃO NA(S) SEGUINTE(S) DISCIPLINA(S): " +
                   ", ".join([comp.split(" (")[0] for comp in recuperacao_needed]))

    # Mensagens de resultado da recuperação
    recuperacao_resultados = check_recuperacao_final(resultado, medias)
    if recuperacao_resultados:
        st.info("Resultado da recuperação: " +
                ", ".join(recuperacao_resultados))


# Inicialização
if "client" not in st.session_state:
    st.session_state["client"] = authenticate_gsheets()
if "cache_version" not in st.session_state:
    st.session_state["cache_version"] = 0

# Carregar dados
df = load_data(WORKSHEET_NOTAS,
               _cache_version=st.session_state["cache_version"])

# Título
st.title("Consulta de Notas 2025")

# Botão de nova consulta
if "consultado" in st.session_state and st.button("Nova consulta"):
    for key in list(st.session_state.keys()):
        if key not in ["client", "cache_version"]:
            del st.session_state[key]
    st.session_state["cache_version"] += 1
    st.rerun()

# 1️⃣ Selecionar Série
series = sorted(df["Série"].dropna().unique().tolist())
serie_selecionada = st.selectbox(
    "Selecione a série:", [""] + series, key="serie")

# 2️⃣ Selecionar Aluno
if serie_selecionada:
    alunos_serie = df[df["Série"] == serie_selecionada][[
        "Nome do Aluno", "Matrícula"]].drop_duplicates()
    nomes = sorted(alunos_serie["Nome do Aluno"].tolist())
    nome_selecionado = st.selectbox(
        "Selecione o aluno:", [""] + nomes, key="nome")

    # 3️⃣ Selecionar Bimestre
    if nome_selecionado:
        bimestres = sorted(df[df["Nome do Aluno"] == nome_selecionado]
                           ["Bimestre"].dropna().unique().tolist())
        bimestre = st.selectbox(
            "Selecione o bimestre/período:", [""] + bimestres + ["Final"], key="bimestre")

        # 4️⃣ Digitar matrícula
        matricula_input = st.text_input(
            "Digite a matrícula do aluno", type="password", key="matricula")

        # 5️⃣ Botão para consultar
        if st.button("Consultar"):
            if not matricula_input:
                st.error("Por favor, digite a matrícula.")
            elif validate_matricula(nome_selecionado, matricula_input, alunos_serie):
                resultado = df[
                    (df['Nome do Aluno'].str.upper() == nome_selecionado.upper()) &
                    (df['Matrícula'].astype(str).str.strip() == matricula_input.strip()) &
                    (df['Série'] == serie_selecionada) &
                    (df['Bimestre'] == bimestre)
                ]
                if not resultado.empty:
                    display_boletim(resultado)
                    st.session_state["consultado"] = True
                    csv = resultado.to_csv(index=False)
                    st.download_button(
                        "Baixar Boletim", csv, f"boletim_{nome_selecionado}_{bimestre}.csv", "text/csv")
                else:
                    st.warning(
                        "Nenhuma nota lançada para esse bimestre/período.")
            else:
                st.error("Matrícula incorreta para o aluno selecionado.")
