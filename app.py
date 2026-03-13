import streamlit as st
from pathlib import Path
import runpy

BASE_DIR = Path(__file__).resolve().parent

st.set_page_config(
    page_title="Central de Conciliações",
    layout="wide",
    initial_sidebar_state="expanded"
)

with st.sidebar:
    st.markdown("# Central de Conciliações")
    visao = st.radio(
        "Selecione a visão",
        ["Análise de Bases", "Conciliação de Extrato"],
        index=0,
    )
    st.markdown("---")
    st.caption("A plataforma foi separada em visões distintas, preservando a lógica específica de cada conciliação.")

if visao == "Análise de Bases":
    runpy.run_path(str(BASE_DIR / "modulo_analise_bases.py"), run_name="__main__")
else:
    runpy.run_path(str(BASE_DIR / "modulo_conciliacao_extrato.py"), run_name="__main__")
