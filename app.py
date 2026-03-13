import streamlit as st
from pathlib import Path

st.set_page_config(page_title="Central de Conciliações", layout="wide")

BASE_DIR = Path(__file__).resolve().parent


def _run_child(module_path: Path):
    namespace = {"__name__": "__main__", "__file__": str(module_path)}
    code = compile(module_path.read_text(encoding="utf-8"), str(module_path), "exec")
    exec(code, namespace)


with st.sidebar:
    st.markdown("# Central de Conciliações")
    visao = st.radio(
        "Selecione a visão",
        ["Análise de Bases", "Conciliação de Extrato"],
        index=0,
    )
    st.markdown("---")
    st.caption("A plataforma agora foi separada em visões distintas, preservando a lógica específica de cada conciliação.")

if visao == "Análise de Bases":
    _run_child(BASE_DIR / "modulo_analise_bases.py")
else:
    _run_child(BASE_DIR / "modulo_conciliacao_extrato.py")
