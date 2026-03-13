import pandas as pd
import numpy as np
import streamlit as st
import io

SEMANTIC_TYPES = ["texto","numero","moeda","percentual","data"]


# ------------------------------------------------------------
# SEMANTIC ENGINE
# ------------------------------------------------------------

def detect_semantic(series):

    if pd.api.types.is_datetime64_any_dtype(series):
        return "data"

    if pd.api.types.is_numeric_dtype(series):

        name = str(series.name).lower()

        if "valor" in name or "total" in name or "saldo" in name:
            return "moeda"

        if "perc" in name or "%" in name:
            return "percentual"

        return "numero"

    return "texto"


def default_excel_format(tipo):

    if tipo == "moeda":
        return "R$ #,##0.00"

    if tipo == "numero":
        return "0.00"

    if tipo == "percentual":
        return "0.00%"

    if tipo == "data":
        return "dd/mm/yyyy"

    return ""


# ------------------------------------------------------------
# MATCH ENGINE
# ------------------------------------------------------------

def build_key(df, keys):

    return df[keys].astype(str).agg("|".join, axis=1)


def conciliar(df1, df2, keys, valores):

    df1 = df1.copy()
    df2 = df2.copy()

    df1["__key"] = build_key(df1, keys["base1"])
    df2["__key"] = build_key(df2, keys["base2"])

    merged = df1.merge(
        df2,
        on="__key",
        how="outer",
        suffixes=("_b1","_b2"),
        indicator=True
    )

    divergencias = []
    ponte = []

    for val in valores:

        c1 = val["base1"]
        c2 = val["base2"]
        nome = val["label"]

        v1 = merged[c1+"_b1"].fillna(0)
        v2 = merged[c2+"_b2"].fillna(0)

        diff = v1 - v2

        ponte.append(pd.DataFrame({
            "Agrupador": merged["__key"],
            "Campo confrontado": nome,
            "Componente":"Valor divergente entre Base 1 e Base 2",
            "Valor": diff
        }))

        diverg = merged[diff != 0]

        divergencias.append(pd.DataFrame({
            "Agrupador": diverg["__key"],
            "Campo confrontado": nome,
            "Base1": diverg[c1+"_b1"],
            "Base2": diverg[c2+"_b2"],
            "Diferença": diverg[c1+"_b1"] - diverg[c2+"_b2"]
        }))

    ponte = pd.concat(ponte)
    divergencias = pd.concat(divergencias)

    return merged, divergencias, ponte


# ------------------------------------------------------------
# EXECUTIVE INDICATORS
# ------------------------------------------------------------

def gerar_indicadores(resultado, divergencias, ponte, campos_confrontados):

    total_b1 = resultado["_merge"].isin(["left_only","both"]).sum()
    total_b2 = resultado["_merge"].isin(["right_only","both"]).sum()

    diverg = len(divergencias)

    dup = resultado["__key"].duplicated().sum()

    so_b1 = (resultado["_merge"]=="left_only").sum()
    so_b2 = (resultado["_merge"]=="right_only").sum()

    conciliados = (resultado["_merge"]=="both").sum()

    impacto = ponte["Valor"].abs().sum()

    valor_b1 = ponte.loc[ponte["Valor"]>0,"Valor"].sum()
    valor_b2 = ponte.loc[ponte["Valor"]<0,"Valor"].sum()

    diff_liq = valor_b1 + valor_b2

    taxa_conc = conciliados / max(total_b1,total_b2)
    taxa_div = diverg / max(total_b1,total_b2)

    indicadores = pd.DataFrame({

        "Indicador":[
            "Campos confrontados",
            "Registros Base 1",
            "Registros Base 2",
            "Itens em divergência",
            "Qtd. em duplicidade",
            "Qtd. só na Base 1",
            "Qtd. só na Base 2",
            "Qtd. conciliados",
            "Impacto absoluto",
            "Valor só na Base 1",
            "Valor só na Base 2",
            "Diferença líquida",
            "Taxa de conciliação",
            "Taxa de divergência"
        ],

        "Valor":[
            campos_confrontados,
            total_b1,
            total_b2,
            diverg,
            dup,
            so_b1,
            so_b2,
            conciliados,
            impacto,
            valor_b1,
            valor_b2,
            diff_liq,
            taxa_conc,
            taxa_div
        ]
    })

    return indicadores


# ------------------------------------------------------------
# EXCEL EXPORT
# ------------------------------------------------------------

def exportar_excel(indicadores, divergencias, ponte):

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        indicadores.to_excel(writer, sheet_name="Resumo", index=False)
        divergencias.to_excel(writer, sheet_name="Divergencias", index=False)
        ponte.to_excel(writer, sheet_name="Ponte_Conciliacao", index=False)

        wb = writer.book

        fmt_moeda = wb.add_format({"num_format":"R$ #,##0.00"})
        fmt_num = wb.add_format({"num_format":"0.00"})
        fmt_perc = wb.add_format({"num_format":"0.00%"})

        ws = writer.sheets["Resumo"]

        for row in range(1,len(indicadores)+1):

            ind = indicadores.iloc[row-1]["Indicador"]

            if "Taxa" in ind:
                ws.write_number(row,1,indicadores.iloc[row-1]["Valor"],fmt_perc)

            elif "Valor" in ind or "Impacto" in ind or "Diferença" in ind:
                ws.write_number(row,1,indicadores.iloc[row-1]["Valor"],fmt_moeda)

            else:
                ws.write_number(row,1,indicadores.iloc[row-1]["Valor"],fmt_num)

    output.seek(0)

    return output


# ------------------------------------------------------------
# STREAMLIT UI
# ------------------------------------------------------------

st.title("Conciliador Inteligente de Bases")

file1 = st.file_uploader("Base 1")
file2 = st.file_uploader("Base 2")

if file1 and file2:

    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    st.write("Base 1",df1.head())
    st.write("Base 2",df2.head())

    keys1 = st.multiselect("Chaves Base 1",df1.columns)
    keys2 = st.multiselect("Chaves Base 2",df2.columns)

    valores1 = st.multiselect("Valores Base 1",df1.columns)
    valores2 = st.multiselect("Valores Base 2",df2.columns)

    if st.button("Conciliar"):

        valores = []

        for i in range(min(len(valores1),len(valores2))):

            valores.append({
                "base1":valores1[i],
                "base2":valores2[i],
                "label":valores1[i]
            })

        merged, divergencias, ponte = conciliar(
            df1,
            df2,
            {"base1":keys1,"base2":keys2},
            valores
        )

        indicadores = gerar_indicadores(
            merged,
            divergencias,
            ponte,
            len(valores)
        )

        excel = exportar_excel(indicadores,divergencias,ponte)

        st.download_button(
            "Baixar Excel",
            excel,
            "conciliacao.xlsx"
        )
