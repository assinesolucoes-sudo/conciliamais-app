# ============================================================
# CONCILIAMAIS
# Conferência de Extrato Bancário
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="ConciliaMais — Conferência de Extrato Bancário",
    layout="wide"
)

# ============================================================
# TÍTULO
# ============================================================

st.title("ConciliaMais — Conferência de Extrato Bancário")
st.caption(
    "Upload do Extrato Financeiro + Razão Contábil → Match automático → Divergências → Painel de fechamento"
)

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def normalize_money(x):

    if pd.isna(x):
        return np.nan

    if isinstance(x, (int,float)):
        return float(x)

    s = str(x)

    s = s.replace("R$","")
    s = s.replace(".","")
    s = s.replace(",",".")
    s = s.strip()

    try:
        return float(s)
    except:
        return np.nan


def extract_doc_key(text):

    if pd.isna(text):
        return ""

    t = str(text)

    nums = re.findall(r"\d{6,}", t)

    if nums:
        return nums[-1]

    return ""


def extract_doc_from_history(text):

    if pd.isna(text):
        return ""

    t = str(text)

    match = re.search(r"/\s*(\d+)", t)

    if match:
        return match.group(1)

    nums = re.findall(r"\d{6,}", t)

    if nums:
        return nums[-1]

    return ""


def read_table(uploaded):

    name = uploaded.name.lower()

    if name.endswith(".csv"):
        return pd.read_csv(uploaded, sep=None, engine="python")

    xl = pd.ExcelFile(uploaded)

    best = None

    for sh in xl.sheet_names:

        tmp = xl.parse(sh)

        if best is None or tmp.shape[1] > best.shape[1]:
            best = tmp

    return best


# ============================================================
# UPLOAD
# ============================================================

c1,c2 = st.columns(2)

with c1:

    st.subheader("Extrato Financeiro")

    fin_file = st.file_uploader(
        "Faça o upload da planilha do Extrato Financeiro",
        type=["xlsx","csv"]
    )

with c2:

    st.subheader("Razão Contábil")

    led_file = st.file_uploader(
        "Faça o upload da planilha do Razão Contábil",
        type=["xlsx","csv"]
    )

if not fin_file or not led_file:

    st.info("Faça o upload dos dois arquivos para liberar o processamento.")

    st.stop()

# ============================================================
# LEITURA
# ============================================================

fin_df = read_table(fin_file)
led_df = read_table(led_file)

# ============================================================
# NORMALIZAÇÃO
# ============================================================

fin_df["VALOR"] = fin_df.iloc[:,-1].map(normalize_money)
led_df["VALOR"] = led_df.iloc[:,-1].map(normalize_money)

fin_df["DOC_KEY"] = fin_df.astype(str).apply(
    lambda r: extract_doc_key(" ".join(r)), axis=1
)

led_df["DOC_KEY"] = led_df.astype(str).apply(
    lambda r: extract_doc_key(" ".join(r)), axis=1
)

# ============================================================
# MATCH
# ============================================================

ledger_map = {}

for i,r in led_df.iterrows():

    key = (round(r["VALOR"],2), r["DOC_KEY"])

    ledger_map.setdefault(key, []).append(i)

used = set()

matches = {}

for i,r in fin_df.iterrows():

    key = (round(r["VALOR"],2), r["DOC_KEY"])

    if key in ledger_map:

        for li in ledger_map[key]:

            if li not in used:

                used.add(li)

                matches[i] = li

                break

# ============================================================
# DIVERGÊNCIAS
# ============================================================

fin_only = fin_df.loc[~fin_df.index.isin(matches.keys())].copy()
led_only = led_df.loc[~led_df.index.isin(used)].copy()

fin_only["ORIGEM"] = "Somente Financeiro"
led_only["ORIGEM"] = "Somente Contábil"

led_only["DOCUMENTO"] = led_only.iloc[:,0].map(extract_doc_from_history)

div = pd.concat([fin_only,led_only])

# remover zeros
div = div[div["VALOR"].abs() > 0]

# ============================================================
# INDICADORES
# ============================================================

fin_total = len(fin_df)
fin_match = len(matches)

fin_pend = len(fin_only)
led_pend = len(led_only)

fin_val = fin_only["VALOR"].sum()
led_val = led_only["VALOR"].sum()

impacto = fin_val - led_val

# ============================================================
# PAINEL
# ============================================================

c1,c2,c3,c4 = st.columns(4)

with c1:

    st.metric(
        "MATCH",
        f"{fin_match}/{fin_total}",
        f"{round(fin_match/fin_total*100,1)}%"
    )

with c2:

    st.metric(
        "Pendentes",
        f"Fin: {fin_pend} | Cont: {led_pend}"
    )

with c3:

    st.metric(
        "Impacto Pendentes",
        f"{impacto:,.2f}"
    )

with c4:

    st.metric(
        "Conferência",
        "0,00"
    )

# ============================================================
# GRÁFICO
# ============================================================

st.subheader("Divergências Financeiro vs Contábil")

chart_df = pd.DataFrame({
    "Tipo":["Financeiro","Contábil"],
    "Valor":[fin_val,led_val]
})

st.bar_chart(chart_df.set_index("Tipo"))

# ============================================================
# FILTROS
# ============================================================

st.subheader("Divergências")

c1,c2,c3,c4 = st.columns([2,2,4,2])

with c1:

    origem = st.selectbox(
        "Filtrar origem",
        ["Todas","Somente Financeiro","Somente Contábil"]
    )

with c2:

    ordenar = st.selectbox(
        "Ordenar",
        ["VALOR","DATA"]
    )

with c3:

    busca = st.text_input(
        "Buscar documento ou histórico"
    )

df = div.copy()

if origem != "Todas":

    df = df[df["ORIGEM"] == origem]

if busca:

    mask = df.astype(str).apply(
        lambda r: busca.lower() in " ".join(r).lower(),
        axis=1
    )

    df = df[mask]

if ordenar in df.columns:

    df = df.sort_values(ordenar)

# ============================================================
# TOTAL FILTRADO
# ============================================================

total_filtrado = df["VALOR"].sum()

st.write(
    f"**Total divergências filtradas:** {total_filtrado:,.2f}"
)

# ============================================================
# TABELA
# ============================================================

st.dataframe(
    df,
    use_container_width=True,
    height=450
)

# ============================================================
# EXPORTAÇÃO
# ============================================================

st.subheader("Exportação")

csv = df.to_csv(index=False).encode()

st.download_button(
    "Exportar Divergências (Excel)",
    csv,
    file_name=f"divergencias_{datetime.now().strftime('%Y%m%d')}.csv"
)

# resumo

resumo = pd.DataFrame({

    "Indicador":[
        "Match",
        "Pendentes Financeiro",
        "Pendentes Contábil",
        "Impacto Pendentes"
    ],

    "Valor":[
        fin_match,
        fin_pend,
        led_pend,
        impacto
    ]

})

csv2 = resumo.to_csv(index=False).encode()

st.download_button(
    "Exportar Relatório Resumo",
    csv2,
    file_name=f"resumo_conciliacao_{datetime.now().strftime('%Y%m%d')}.csv"
)
