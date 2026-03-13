import streamlit as st
import pandas as pd

def run():

    st.title("Conciliação de Extrato")
    st.caption("Reconciliação entre extrato financeiro e razão contábil")

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        extrato = st.file_uploader("Upload Extrato Bancário", type=["xlsx","csv"])

    with col2:
        razao = st.file_uploader("Upload Razão Financeira", type=["xlsx","csv"])

    if extrato and razao:

        df_extrato = pd.read_excel(extrato) if extrato.name.endswith("xlsx") else pd.read_csv(extrato)
        df_razao = pd.read_excel(razao) if razao.name.endswith("xlsx") else pd.read_csv(razao)

        st.success("Arquivos carregados com sucesso")

        st.subheader("Extrato")
        st.dataframe(df_extrato)

        st.subheader("Razão")
        st.dataframe(df_razao)

        st.info("Aqui será executado o motor de conciliação de extrato.")

