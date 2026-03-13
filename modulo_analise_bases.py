import streamlit as st
import pandas as pd

def run():

    st.title("Análise de Bases")
    st.caption("Comparação estruturada entre duas bases de dados")

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        base_a = st.file_uploader("Upload Base A", type=["xlsx","csv"])

    with col2:
        base_b = st.file_uploader("Upload Base B", type=["xlsx","csv"])

    if base_a and base_b:

        df_a = pd.read_excel(base_a) if base_a.name.endswith("xlsx") else pd.read_csv(base_a)
        df_b = pd.read_excel(base_b) if base_b.name.endswith("xlsx") else pd.read_csv(base_b)

        st.success("Bases carregadas com sucesso")

        st.subheader("Base A")
        st.dataframe(df_a)

        st.subheader("Base B")
        st.dataframe(df_b)

        st.info("Aqui será executado o motor de comparação de bases.")

