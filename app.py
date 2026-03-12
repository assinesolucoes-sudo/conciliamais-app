# V26 – Ajuste estrutural do relatório executivo e ponte de conciliação
# Alterações desta versão
# 1) Resumo Executivo agora possui DOIS BLOCOS separados
#    - Bloco 1: Fechamento Global
#    - Bloco 2: Diferença por Agrupador (tabela filtrável)
# 2) Removidos cabeçalhos duplicados
# 3) Ponte de Conciliação também separada em dois blocos
#    - Total Geral (fora do filtro)
#    - Por Agrupador (com filtro + total filtrado)
# 4) Layout Excel com aparência de relatório

import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Concilia Mais", layout="wide")

# =====================================================
# FUNÇÕES AUXILIARES
# =====================================================

def to_number(col):
    s = col.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    s = s.str.replace("R$", "", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0)


def moeda(x):
    return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# =====================================================
# ENGINE DE CONCILIAÇÃO
# =====================================================

def conciliar(dfA, dfB, chaveA, chaveB, valoresA, valoresB, nomeA, nomeB):

    dfA["_KEY"] = dfA[chaveA].astype(str)
    dfB["_KEY"] = dfB[chaveB].astype(str)

    merged = dfA.merge(dfB, left_on="_KEY", right_on="_KEY", how="outer", indicator=True)

    resultados = []

    for a, b in zip(valoresA, valoresB):
        colA = to_number(merged[a])
        colB = to_number(merged[b])

        merged[f"VAL_A_{a}"] = colA
        merged[f"VAL_B_{b}"] = colB
        merged[f"DIF_{a}"] = colA - colB

        resultados.append((a, b))

    merged["MOTIVO"] = np.select(
        [
            merged["_merge"] == "left_only",
            merged["_merge"] == "right_only",
            merged[[c for c in merged.columns if c.startswith("DIF_")]].abs().sum(axis=1) > 0
        ],
        [
            f"Chave só na {nomeA}",
            f"Chave só na {nomeB}",
            f"Valor divergente entre {nomeA} e {nomeB}"
        ],
        default="Conciliado"
    )

    return merged, resultados

# =====================================================
# RELATÓRIOS
# =====================================================

def gerar_resumo_global(df, campos):

    linhas = []

    for a,b in campos:

        totalA = df[f"VAL_A_{a}"].sum()
        totalB = df[f"VAL_B_{b}"].sum()

        linhas.append({
            "Campo confrontado": a,
            "Total Base A": totalA,
            "Total Base B": totalB,
            "Diferença total": totalA-totalB
        })

    return pd.DataFrame(linhas)


def gerar_resumo_agrupador(df, agrupador, campos):

    tabela = []

    grupos = df.groupby(agrupador)

    for g, dados in grupos:

        linha = {"Agrupador": g}

        for a,b in campos:

            linha[f"{a} A"] = dados[f"VAL_A_{a}"].sum()
            linha[f"{a} B"] = dados[f"VAL_B_{b}"].sum()
            linha[f"Dif {a}"] = linha[f"{a} A"] - linha[f"{a} B"]

        motivo = dados[dados["MOTIVO"]!="Conciliado"]["MOTIVO"]

        if len(motivo)>0:
            linha["Motivo predominante"] = motivo.value_counts().index[0]
        else:
            linha["Motivo predominante"] = "Sem diferença"

        tabela.append(linha)

    return pd.DataFrame(tabela)


def gerar_ponte_total(df, campos):

    linhas = []

    for a,b in campos:

        dif = df[f"VAL_A_{a}"].sum() - df[f"VAL_B_{b}"].sum()

        linhas.append({
            "Campo":a,
            "Componente":"Diferença final",
            "Valor":dif
        })

    return pd.DataFrame(linhas)


def gerar_ponte_agrupador(df, agrupador, campos):

    linhas = []

    grupos = df.groupby([agrupador,"MOTIVO"])

    for (g,m),dados in grupos:

        for a,b in campos:

            val = dados[f"VAL_A_{a}"].sum() - dados[f"VAL_B_{b}"].sum()

            if val!=0:

                linhas.append({
                    "Agrupador":g,
                    "Campo":a,
                    "Motivo":m,
                    "Valor":val
                })

    return pd.DataFrame(linhas)

# =====================================================
# EXPORTAÇÃO EXCEL
# =====================================================

def exportar_excel(resumo_global,resumo_agrup,ponte_total,ponte_agrup,detalhe):

    buffer = BytesIO()

    with pd.ExcelWriter(buffer,engine="xlsxwriter") as writer:

        wb = writer.book

        fmt_title = wb.add_format({'bold':True,'font_size':14})
        fmt_head = wb.add_format({'bold':True,'bg_color':'#D9EAF7'})

        # =====================================================
        # RESUMO EXECUTIVO
        # =====================================================

        ws = wb.add_worksheet("RESUMO_EXECUTIVO")

        ws.write("A1","RESUMO EXECUTIVO",fmt_title)

        ws.write("A3","Fechamento global",fmt_head)

        resumo_global.to_excel(writer,sheet_name="RESUMO_EXECUTIVO",startrow=3,index=False)

        linha = len(resumo_global)+6

        ws.write(linha-1,0,"Diferença por agrupador",fmt_head)

        resumo_agrup.to_excel(writer,sheet_name="RESUMO_EXECUTIVO",startrow=linha,index=False)

        # =====================================================
        # PONTE
        # =====================================================

        ws2 = wb.add_worksheet("PONTE_CONCILIACAO")

        ws2.write("A1","PONTE DA CONCILIAÇÃO",fmt_title)

        ws2.write("A3","Total Geral",fmt_head)

        ponte_total.to_excel(writer,sheet_name="PONTE_CONCILIACAO",startrow=3,index=False)

        linha2 = len(ponte_total)+6

        ws2.write(linha2-1,0,"Diferença por agrupador",fmt_head)

        ponte_agrup.to_excel(writer,sheet_name="PONTE_CONCILIACAO",startrow=linha2,index=False)

        # =====================================================
        # DETALHE
        # =====================================================

        detalhe.to_excel(writer,sheet_name="DETALHE_DIFERENCAS",index=False)

    return buffer.getvalue()

# =====================================================
# APP
# =====================================================

st.title("Concilia Mais")

st.write("Ferramenta de conciliação entre duas bases")

nomeA = st.text_input("Nome Base 1","RM")
nomeB = st.text_input("Nome Base 2","Protheus")

arqA = st.file_uploader("Base 1",type=["xlsx","csv"])
arqB = st.file_uploader("Base 2",type=["xlsx","csv"])

if arqA and arqB:

    dfA = pd.read_excel(arqA)
    dfB = pd.read_excel(arqB)

    colA = st.selectbox("Chave Base 1",dfA.columns)
    colB = st.selectbox("Chave Base 2",dfB.columns)

    valA = st.multiselect("Campos valor Base 1",dfA.columns)
    valB = st.multiselect("Campos valor Base 2",dfB.columns)

    agrup = st.selectbox("Agrupar por",dfA.columns)

    if st.button("Executar conciliação"):

        df,campos = conciliar(dfA,dfB,colA,colB,valA,valB,nomeA,nomeB)

        resumo_global = gerar_resumo_global(df,campos)
        resumo_agrup = gerar_resumo_agrupador(df,agrup,campos)

        ponte_total = gerar_ponte_total(df,campos)
        ponte_agrup = gerar_ponte_agrupador(df,agrup,campos)

        detalhe = df[df["MOTIVO"]!="Conciliado"]

        st.dataframe(resumo_global)

        excel = exportar_excel(resumo_global,resumo_agrup,ponte_total,ponte_agrup,detalhe)

        st.download_button("Baixar Excel",excel,"conciliacao.xlsx")
