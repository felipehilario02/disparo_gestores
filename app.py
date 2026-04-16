import streamlit as st
import os
from datetime import datetime

from comentarios import carregar_base_comentarios
from config import (
    BASE_DIR,
    COMENTARIOS_FILE,
    COMENTARIOS_SHEET,
    DATE_FMT,
    FORCE_RESP_COL_NAME,
    FULL_BASE_RECIPIENTS,
    INPUT_FILE,
    SHEET_NAME,
)
from io_excel import read_excel_robusto
from helpers import mes_anterior_pt_e_fy
from processor import (
    ExecutionResult,
    process_base_completa,
    process_gestor_normal,
)

st.set_page_config(
    page_title="Envio de Relatórios de Despesas",
    page_icon="📊",
    layout="centered",
)

st.title("📊 Envio de Relatórios de Despesas")
st.caption("Execução manual e controlada do envio de relatórios financeiros.")

st.warning(
    "⚠️ Execute **apenas um gestor por vez**. "
    "Esse modo garante a entrega correta conforme as políticas de e-mail."
)

@st.cache_data(show_spinner=True)
def carregar_base():
    base_path = os.path.join(BASE_DIR, INPUT_FILE)
    df = read_excel_robusto(base_path, sheet_name=SHEET_NAME, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df


df = carregar_base()

conta_comments = carregar_base_comentarios(
    alt_file=COMENTARIOS_FILE,
    sheet_name=COMENTARIOS_SHEET,
)

col_resp = (
    FORCE_RESP_COL_NAME
    if FORCE_RESP_COL_NAME in df.columns
    else df.columns[9]
)

gestores_base = (
    df[col_resp]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
    .tolist()
)

gestores = sorted(set(gestores_base) | set(FULL_BASE_RECIPIENTS))

st.subheader("1️⃣ Seleção do Gestor")

gestor = st.selectbox(
    "Selecione o gestor para envio",
    options=gestores,
)

if gestor in FULL_BASE_RECIPIENTS:
    st.success(
        f"✅ **{gestor}** receberá a **BASE COMPLETA**, "
        
    )
else:
    st.info(
        f"ℹ️ **{gestor}** receberá apenas os dados da sua área."
    )

st.subheader("2️⃣ Confirmação")
confirmar = st.checkbox(
    f"Confirmo o envio do relatório para **{gestor}**"
)

if st.button("🚀 Processar e Enviar", type="primary", disabled=not confirmar):

    with st.spinner("Processando relatório e enviando e-mail..."):
        try:
            today = datetime.today().strftime(DATE_FMT)
            mes_anterior_pt, fy = mes_anterior_pt_e_fy()
            result = ExecutionResult()

            # ✅ REGRA DE NEGÓCIO SEMPRE VEM PRIMEIRO
            if gestor in FULL_BASE_RECIPIENTS:
                process_base_completa(
                    df,
                    gestor,
                    today,
                    mes_anterior_pt,
                    fy,
                    result=result,
                    conta_comment_map=conta_comments,
                )
            else:
                df_resp = df[df[col_resp] == gestor]

                if df_resp.empty:
                    st.error(
                        f"❌ O gestor **{gestor}** não foi encontrado na base "
                        "e não pertence ao grupo de base completa."
                    )
                    st.stop()

                process_gestor_normal(
                    df_resp,
                    gestor,
                    today,
                    mes_anterior_pt,
                    fy,
                    result=result,
                    conta_comment_map=conta_comments,
                )

            st.success("✅ Envio concluído com sucesso!")
            st.code(result.resumo(), language="text")

        except Exception as exc:
            st.error("❌ Ocorreu um erro durante o processamento.")
            st.exception(exc)

st.divider()
st.caption("Uso interno Financeiro • Desenvolvido por Yago Costa")