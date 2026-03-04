
import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(page_title="ConciliaMais — Módulo 1 (MVP)", layout="wide")

# ----------------------------
# Helpers
# ----------------------------
def _to_date_series(s):
    # Try multiple date formats and excel dates
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s).dt.date
    out = pd.to_datetime(s, errors="coerce", dayfirst=True)
    # If many NaT, try ISO
    if out.notna().mean() < 0.6:
        out = pd.to_datetime(s, errors="coerce")
    return out.dt.date

def extract_doc_key(text):
    if pd.isna(text):
        return ""
    t = str(text)
    # Prefer sequences length >= 6 (NF/IMP/etc.)
    nums = re.findall(r"\d{6,}", t)
    if nums:
        # take the longest, then the last (often right-aligned)
        nums = sorted(nums, key=lambda x: (len(x), t.rfind(x)))
        return nums[-1]
    return ""

def normalize_money(x):
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip()
    if s == "":
        return 0.0
    # Brazilian formats: 1.234,56
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def read_table(uploaded):
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded, sep=None, engine="python")
        return df
    # xlsx
    xl = pd.ExcelFile(uploaded)
    # Try to pick the first sheet with > 5 columns
    best = None
    for sh in xl.sheet_names:
        tmp = xl.parse(sh)
        if best is None or tmp.shape[1] > best.shape[1]:
            best = tmp
    return best

def auto_detect_financial(df):
    cols = {c.lower(): c for c in df.columns}
    def pick(*cands):
        for c in cands:
            if c in cols:
                return cols[c]
        return None
    date = pick("data", "dt", "data mov", "data_mov")
    entradas = pick("entradas", "entrada", "credito", "crédito")
    saidas = pick("saidas", "saídas", "saida", "saída", "debito", "débito")
    saldo = pick("saldo atual", "saldo", "saldo_atual")
    operacao = pick("operacao", "operação", "historico", "histórico")
    documento = pick("documento", "doc", "num documento", "número do documento")
    prefixo = pick("prefixo/titulo", "prefixo/título", "prefixo", "titulo", "título")
    return dict(date=date, entradas=entradas, saidas=saidas, saldo=saldo, operacao=operacao, documento=documento, prefixo=prefixo)

def auto_detect_ledger(df):
    cols = {c.lower(): c for c in df.columns}
    def pick(*cands):
        for c in cands:
            if c in cols:
                return cols[c]
        return None
    date = pick("data", "dt", "data lanc", "data_lanc")
    debito = pick("debito", "débito", "debit")
    credito = pick("credito", "crédito", "credit")
    saldo = pick("saldo atual", "saldo", "saldo_atual")
    historico = pick("historico", "histórico", "operacao", "operação", "descricao", "descrição")
    doc = pick("lote/sub/doc/linha", "documento", "doc", "num documento")
    conta = pick("conta")
    return dict(date=date, debito=debito, credito=credito, saldo=saldo, historico=historico, doc=doc, conta=conta)

def reconcile(fin_df, led_df, cfg, date_tol_days=0):
    # Build normalized finance
    f = fin_df.copy()
    f["__date"] = _to_date_series(f[cfg["fin_date"]])
    f["__entradas"] = f[cfg["fin_entradas"]].map(normalize_money) if cfg["fin_entradas"] else 0.0
    f["__saidas"] = f[cfg["fin_saidas"]].map(normalize_money) if cfg["fin_saidas"] else 0.0
    if cfg.get("fin_amount") and cfg["fin_amount"]:
        f["__amount"] = f[cfg["fin_amount"]].map(normalize_money)
    else:
        f["__amount"] = f["__entradas"] - f["__saidas"]
    f["__saldo"] = f[cfg["fin_saldo"]].map(normalize_money) if cfg["fin_saldo"] else np.nan
    op_col = cfg.get("fin_operacao")
    doc_col = cfg.get("fin_documento")
    pre_col = cfg.get("fin_prefixo")
    f["__text"] = (
        (f[op_col].astype(str) if op_col else "") + " " +
        (f[doc_col].astype(str) if doc_col else "") + " " +
        (f[pre_col].astype(str) if pre_col else "")
    )
    f["__doc_key"] = f["__text"].map(extract_doc_key)
    f["__idx"] = np.arange(len(f))

    # Build normalized ledger
    l = led_df.copy()
    l["__date"] = _to_date_series(l[cfg["led_date"]])
    l["__deb"] = l[cfg["led_debito"]].map(normalize_money) if cfg["led_debito"] else 0.0
    l["__cred"] = l[cfg["led_credito"]].map(normalize_money) if cfg["led_credito"] else 0.0
    if cfg.get("led_amount") and cfg["led_amount"]:
        l["__amount"] = l[cfg["led_amount"]].map(normalize_money)
    else:
        l["__amount"] = l["__deb"] - l["__cred"]
    l["__saldo"] = l[cfg["led_saldo"]].map(normalize_money) if cfg["led_saldo"] else np.nan
    hist_col = cfg.get("led_historico")
    doc_col2 = cfg.get("led_doc")
    conta_col = cfg.get("led_conta")
    l["__text"] = (
        (l[hist_col].astype(str) if hist_col else "") + " " +
        (l[doc_col2].astype(str) if doc_col2 else "") + " " +
        (l[conta_col].astype(str) if conta_col else "")
    )
    l["__doc_key"] = l["__text"].map(extract_doc_key)
    l["__idx"] = np.arange(len(l))

    # Index ledger candidates by (amount, doc_key)
    ledger_used = set()
    fin_match = {}
    led_match = {}

    def try_match(fin_row_idx, candidates):
        for li in candidates:
            if li in ledger_used:
                continue
            ledger_used.add(li)
            fin_match[fin_row_idx] = li
            led_match[li] = fin_row_idx
            return True
        return False

    # Primary: amount + doc_key
    key_to_led = {}
    for i, r in l.iterrows():
        key = (round(r["__amount"], 2), r["__doc_key"])
        key_to_led.setdefault(key, []).append(r["__idx"])

    for i, r in f.iterrows():
        if r["__doc_key"]:
            key = (round(r["__amount"], 2), r["__doc_key"])
            if key in key_to_led:
                try_match(r["__idx"], key_to_led[key])

    # Secondary: amount + date (with tolerance)
    # Build dict for ledger by amount
    amt_to_led = {}
    for i, r in l.iterrows():
        amt_to_led.setdefault(round(r["__amount"], 2), []).append(r["__idx"])

    # Quick map from idx to row for ledger
    l_by_idx = l.set_index("__idx", drop=False)

    for i, r in f.iterrows():
        fi = r["__idx"]
        if fi in fin_match:
            continue
        amt = round(r["__amount"], 2)
        if amt not in amt_to_led:
            continue
        # filter by date tol
        fdate = r["__date"]
        cands = []
        for li in amt_to_led[amt]:
            if li in ledger_used:
                continue
            ldate = l_by_idx.loc[li, "__date"]
            if pd.isna(fdate) or pd.isna(ldate):
                continue
            if abs((pd.to_datetime(fdate) - pd.to_datetime(ldate)).days) <= date_tol_days:
                cands.append(li)
        if cands:
            try_match(fi, cands)

    # Build divergences table
    # Finance-only
    fin_only = f[~f["__idx"].isin(fin_match.keys())].copy()
    led_only = l[~l["__idx"].isin(led_match.keys())].copy()

    # Add status columns to original frames
    f_out = fin_df.copy()
    f_out["CONCILIADO?"] = f["__idx"].map(lambda x: "S" if x in fin_match else "N")
    f_out["PAREADO_COM_IDX_CONTABIL"] = f["__idx"].map(lambda x: fin_match.get(x, ""))

    l_out = led_df.copy()
    l_out["CONCILIADO?"] = l["__idx"].map(lambda x: "S" if x in led_match else "N")
    l_out["PAREADO_COM_IDX_FINANCEIRO"] = l["__idx"].map(lambda x: led_match.get(x, ""))

    def build_div(df, side):
        # side: FIN or LED
        rows = []
        if side == "FIN":
            for _, r in df.iterrows():
                rows.append({
                    "DATA": r["__date"],
                    "CHAVE_DOC": r["__doc_key"],
                    "HISTÓRICO/OPERAÇÃO": str(r["__text"]).strip(),
                    "VALOR": round(r["__amount"], 2),
                    "ORIGEM": "Somente Financeiro",
                    "IDX_ORIGEM": int(r["__idx"]),
                    "SALDO_NA_LINHA": r["__saldo"],
                    "PAREADO?": "N",
                })
        else:
            for _, r in df.iterrows():
                rows.append({
                    "DATA": r["__date"],
                    "CHAVE_DOC": r["__doc_key"],
                    "HISTÓRICO/OPERAÇÃO": str(r["__text"]).strip(),
                    "VALOR": round(r["__amount"], 2),
                    "ORIGEM": "Somente Contábil",
                    "IDX_ORIGEM": int(r["__idx"]),
                    "SALDO_NA_LINHA": r["__saldo"],
                    "PAREADO?": "N",
                })
        return pd.DataFrame(rows)

    div = pd.concat([build_div(fin_only, "FIN"), build_div(led_only, "LED")], ignore_index=True)

    # Summary / bridge values
    fin_total = round(f["__amount"].sum(), 2)
    led_total = round(l["__amount"].sum(), 2)
    fin_unmatched = round(fin_only["__amount"].sum(), 2)
    led_unmatched = round(led_only["__amount"].sum(), 2)
    diff = round(fin_total - led_total, 2)

    summary = {
        "Total Financeiro (soma movimentos)": fin_total,
        "Total Contábil (soma lançamentos)": led_total,
        "Diferença Financeiro - Contábil": diff,
        "Soma pendente (Somente Financeiro)": fin_unmatched,
        "Soma pendente (Somente Contábil)": led_unmatched,
        "Pendentes (FIN) - (LED)": round(fin_unmatched - led_unmatched, 2),
        "Fechamento (deve bater com Diferença)": round((fin_unmatched - led_unmatched) - diff, 2),
    }

    return f_out, l_out, div, summary

def build_excel(fin_out, led_out, div, summary):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        fin_out.to_excel(writer, index=False, sheet_name="Extrato_Financeiro")
        led_out.to_excel(writer, index=False, sheet_name="Extrato_Contabil")
        div.to_excel(writer, index=False, sheet_name="Divergencias")
        pd.DataFrame(list(summary.items()), columns=["Metrica", "Valor"]).to_excel(
            writer, index=False, sheet_name="Ponte_Resumo"
        )
        # Basic formatting
        wb = writer.book
        fmt_money = wb.add_format({"num_format": "#,##0.00"})
        fmt_hdr = wb.add_format({"bold": True})
        for sh in ["Extrato_Financeiro", "Extrato_Contabil", "Divergencias", "Ponte_Resumo"]:
            ws = writer.sheets[sh]
            ws.set_row(0, None, fmt_hdr)
            ws.autofilter(0, 0, max(0, (writer.sheets[sh].dim_rowmax)), max(0, (writer.sheets[sh].dim_colmax)))
            ws.freeze_panes(1, 0)
        # Apply money format heuristically
        for sh, df in [("Extrato_Financeiro", fin_out), ("Extrato_Contabil", led_out), ("Divergencias", div), ("Ponte_Resumo", None)]:
            ws = writer.sheets[sh]
            if df is not None:
                for j, c in enumerate(df.columns):
                    if any(k in str(c).lower() for k in ["valor", "saldo", "deb", "cred", "entrada", "saida"]):
                        ws.set_column(j, j, 18, fmt_money)
                    else:
                        ws.set_column(j, j, 22)
            else:
                ws.set_column(0, 0, 38)
                ws.set_column(1, 1, 18, fmt_money)
    output.seek(0)
    return output

# ----------------------------
# UI
# ----------------------------
st.title("ConciliaMais — Módulo 1 (MVP)")
st.caption("Upload do Extrato Financeiro + Razão Contábil → Match automático → Divergências → Ponte de fechamento")

c1, c2 = st.columns(2)
with c1:
    fin_file = st.file_uploader("Upload — Extrato Financeiro (.xlsx ou .csv)", type=["xlsx", "csv"], key="fin")
with c2:
    led_file = st.file_uploader("Upload — Razão Contábil (.xlsx ou .csv)", type=["xlsx", "csv"], key="led")

st.divider()

if fin_file and led_file:
    fin_df = read_table(fin_file)
    led_df = read_table(led_file)

    st.subheader("Mapeamento de colunas (auto-detectado — ajuste se precisar)")
    fin_guess = auto_detect_financial(fin_df)
    led_guess = auto_detect_ledger(led_df)

    colA, colB = st.columns(2)

    with colA:
        st.markdown("**Extrato Financeiro**")
        fin_date = st.selectbox("Coluna de Data", fin_df.columns, index=fin_df.columns.get_loc(fin_guess["date"]) if fin_guess["date"] in fin_df.columns else 0)
        fin_operacao = st.selectbox("Coluna de Operação/Histórico", ["(nenhuma)"] + list(fin_df.columns),
                                   index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["operacao"]) if fin_guess["operacao"] in fin_df.columns else 0)
        fin_documento = st.selectbox("Coluna de Documento", ["(nenhuma)"] + list(fin_df.columns),
                                     index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["documento"]) if fin_guess["documento"] in fin_df.columns else 0)
        fin_prefixo = st.selectbox("Coluna de Prefixo/Título", ["(nenhuma)"] + list(fin_df.columns),
                                   index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["prefixo"]) if fin_guess["prefixo"] in fin_df.columns else 0)
        fin_entradas = st.selectbox("Coluna de Entradas", ["(nenhuma)"] + list(fin_df.columns),
                                    index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["entradas"]) if fin_guess["entradas"] in fin_df.columns else 0)
        fin_saidas = st.selectbox("Coluna de Saídas", ["(nenhuma)"] + list(fin_df.columns),
                                  index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["saidas"]) if fin_guess["saidas"] in fin_df.columns else 0)
        fin_amount = st.selectbox("OU coluna de Valor Único (opcional)", ["(usar Entradas - Saídas)"] + list(fin_df.columns), index=0)
        fin_saldo = st.selectbox("Coluna de Saldo (opcional)", ["(nenhuma)"] + list(fin_df.columns),
                                 index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["saldo"]) if fin_guess["saldo"] in fin_df.columns else 0)

    with colB:
        st.markdown("**Razão Contábil**")
        led_date = st.selectbox("Coluna de Data", led_df.columns, index=led_df.columns.get_loc(led_guess["date"]) if led_guess["date"] in led_df.columns else 0, key="led_date")
        led_historico = st.selectbox("Coluna de Histórico", ["(nenhuma)"] + list(led_df.columns),
                                     index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["historico"]) if led_guess["historico"] in led_df.columns else 0, key="led_hist")
        led_doc = st.selectbox("Coluna de Documento/Lote (opcional)", ["(nenhuma)"] + list(led_df.columns),
                               index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["doc"]) if led_guess["doc"] in led_df.columns else 0, key="led_doc")
        led_conta = st.selectbox("Coluna de Conta (opcional)", ["(nenhuma)"] + list(led_df.columns),
                                 index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["conta"]) if led_guess["conta"] in led_df.columns else 0, key="led_conta")
        led_debito = st.selectbox("Coluna de Débito", ["(nenhuma)"] + list(led_df.columns),
                                  index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["debito"]) if led_guess["debito"] in led_df.columns else 0, key="led_deb")
        led_credito = st.selectbox("Coluna de Crédito", ["(nenhuma)"] + list(led_df.columns),
                                   index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["credito"]) if led_guess["credito"] in led_df.columns else 0, key="led_cred")
        led_amount = st.selectbox("OU coluna de Valor Único (opcional)", ["(usar Débito - Crédito)"] + list(led_df.columns), index=0, key="led_amount")
        led_saldo = st.selectbox("Coluna de Saldo (opcional)", ["(nenhuma)"] + list(led_df.columns),
                                 index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["saldo"]) if led_guess["saldo"] in led_df.columns else 0, key="led_saldo")

    st.divider()
    st.subheader("Parâmetros de conciliação")
    p1, p2, p3 = st.columns(3)
    with p1:
        date_tol = st.number_input("Tolerância de dias para casar por data (0 = mesma data)", min_value=0, max_value=10, value=0, step=1)
    with p2:
        st.caption("Ordem do match: (1) Valor + DocKey, (2) Valor + Data (com tolerância)")
        st.write("")
    with p3:
        st.caption("DocKey: sequência numérica (>=6 dígitos) extraída de Operação/Histórico/Documento.")
        st.write("")

    cfg = {
        "fin_date": fin_date,
        "fin_operacao": None if fin_operacao == "(nenhuma)" else fin_operacao,
        "fin_documento": None if fin_documento == "(nenhuma)" else fin_documento,
        "fin_prefixo": None if fin_prefixo == "(nenhuma)" else fin_prefixo,
        "fin_entradas": None if fin_entradas == "(nenhuma)" else fin_entradas,
        "fin_saidas": None if fin_saidas == "(nenhuma)" else fin_saidas,
        "fin_amount": None if fin_amount == "(usar Entradas - Saídas)" else fin_amount,
        "fin_saldo": None if fin_saldo == "(nenhuma)" else fin_saldo,
        "led_date": led_date,
        "led_historico": None if led_historico == "(nenhuma)" else led_historico,
        "led_doc": None if led_doc == "(nenhuma)" else led_doc,
        "led_conta": None if led_conta == "(nenhuma)" else led_conta,
        "led_debito": None if led_debito == "(nenhuma)" else led_debito,
        "led_credito": None if led_credito == "(nenhuma)" else led_credito,
        "led_amount": None if led_amount == "(usar Débito - Crédito)" else led_amount,
        "led_saldo": None if led_saldo == "(nenhuma)" else led_saldo,
    }

    run = st.button("Rodar conciliação agora", type="primary")

    if run:
        with st.spinner("Processando..."):
            fin_out, led_out, div, summary = reconcile(fin_df, led_df, cfg, date_tol_days=int(date_tol))

        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Total Financeiro", f'{summary["Total Financeiro (soma movimentos)"]:.2f}')
        s2.metric("Total Contábil", f'{summary["Total Contábil (soma lançamentos)"]:.2f}')
        s3.metric("Diferença (FIN - CONT)", f'{summary["Diferença Financeiro - Contábil"]:.2f}')
        s4.metric("Fechamento (ideal 0,00)", f'{summary["Fechamento (deve bater com Diferença)"]:.2f}')

        st.subheader("Divergências (itens não pareados)")
        st.dataframe(div, use_container_width=True, height=320)

        excel_bytes = build_excel(fin_out, led_out, div, summary)
        st.download_button(
            "Baixar relatório Excel (ConciliaMais)",
            data=excel_bytes,
            file_name=f"ConciliaMais_Resultado_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.info("Dica: se a diferença final não fechar, ajuste o mapeamento de colunas e/ou a tolerância de datas.")
else:
    st.info("Faça o upload do Extrato Financeiro e do Razão Contábil para liberar o processamento.")
