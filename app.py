import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="ConciliaMais — Módulo 1 (MVP)", layout="wide")

# ----------------------------
# Helpers
# ----------------------------
def _to_date_series(s):
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s).dt.date
    out = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if out.notna().mean() < 0.6:
        out = pd.to_datetime(s, errors="coerce")
    return out.dt.date

def extract_doc_key(text):
    if pd.isna(text):
        return ""
    t = str(text)
    nums = re.findall(r"\d{6,}", t)
    if nums:
        nums = sorted(nums, key=lambda x: (len(x), t.rfind(x)))
        return nums[-1]
    return ""

def normalize_money(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip()
    if s == "":
        return np.nan
    s = s.replace("R$", "").strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return np.nan

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

def auto_detect_financial(df):
    cols = {c.lower(): c for c in df.columns}
    def pick(*cands):
        for c in cands:
            if c in cols:
                return cols[c]
        return None
    return dict(
        date=pick("data", "dt", "data mov", "data_mov"),
        entradas=pick("entradas", "entrada", "credito", "crédito"),
        saidas=pick("saidas", "saídas", "saida", "saída", "debito", "débito"),
        saldo=pick("saldo atual", "saldo", "saldo_atual"),
        operacao=pick("operacao", "operação", "historico", "histórico"),
        documento=pick("documento", "doc", "num documento", "número do documento"),
        prefixo=pick("prefixo/titulo", "prefixo/título", "prefixo", "titulo", "título"),
        valor=pick("valor", "vlr", "valor mov", "valor_mov"),
    )

def auto_detect_ledger(df):
    cols = {c.lower(): c for c in df.columns}
    def pick(*cands):
        for c in cands:
            if c in cols:
                return cols[c]
        return None
    return dict(
        date=pick("data", "dt", "data lanc", "data_lanc"),
        debito=pick("debito", "débito", "debit"),
        credito=pick("credito", "crédito", "credit"),
        saldo=pick("saldo atual", "saldo", "saldo_atual"),
        historico=pick("historico", "histórico", "operacao", "operação", "descricao", "descrição"),
        doc=pick("lote/sub/doc/linha", "documento", "doc", "num documento"),
        conta=pick("conta"),
        valor=pick("valor", "vlr"),
    )

def build_normalized(fin_df, led_df, cfg):
    # Finance
    f = fin_df.copy()
    f["__date"] = _to_date_series(f[cfg["fin_date"]])

    if cfg.get("fin_amount"):
        f["__amount"] = f[cfg["fin_amount"]].map(normalize_money)
    else:
        entradas = f[cfg["fin_entradas"]].map(normalize_money) if cfg.get("fin_entradas") else 0.0
        saidas = f[cfg["fin_saidas"]].map(normalize_money) if cfg.get("fin_saidas") else 0.0
        entradas = pd.Series(entradas).fillna(0.0)
        saidas = pd.Series(saidas).fillna(0.0)
        f["__amount"] = entradas - saidas

    f["__saldo"] = f[cfg["fin_saldo"]].map(normalize_money) if cfg.get("fin_saldo") else np.nan

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

    # Ledger
    l = led_df.copy()
    l["__date"] = _to_date_series(l[cfg["led_date"]])

    if cfg.get("led_amount"):
        l["__amount"] = l[cfg["led_amount"]].map(normalize_money)
    else:
        deb = l[cfg["led_debito"]].map(normalize_money) if cfg.get("led_debito") else 0.0
        cred = l[cfg["led_credito"]].map(normalize_money) if cfg.get("led_credito") else 0.0
        deb = pd.Series(deb).fillna(0.0)
        cred = pd.Series(cred).fillna(0.0)
        l["__amount"] = deb - cred

    l["__saldo"] = l[cfg["led_saldo"]].map(normalize_money) if cfg.get("led_saldo") else np.nan

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

    return f, l

def compute_saldo_anterior(df_norm):
    # usa SOMENTE linhas de movimento (DATA preenchida)
    dfv = df_norm.copy()
    dfv = dfv[dfv["__date"].notna()].copy()
    dfv = dfv[dfv["__amount"].notna()].copy()
    dfv = dfv[dfv["__saldo"].notna()].copy()
    if dfv.empty:
        return np.nan
    r = dfv.iloc[0]
    return round(float(r["__saldo"]) - float(r["__amount"]), 2)

def compute_saldo_final(df_norm):
    # usa SOMENTE linhas de movimento (DATA preenchida)
    dfv = df_norm.copy()
    dfv = dfv[dfv["__date"].notna()].copy()
    dfv = dfv[dfv["__amount"].notna()].copy()
    dfv = dfv[dfv["__saldo"].notna()].copy()
    if dfv.empty:
        return np.nan
    r = dfv.iloc[-1]
    return round(float(r["__saldo"]), 2)

def reconcile(fin_df, led_df, cfg, date_tol_days=0):
    f, l = build_normalized(fin_df, led_df, cfg)

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

    # Index ledger by (amount, doc_key)
    key_to_led = {}
    for _, r in l.iterrows():
        key = (round(float(r["__amount"]), 2), r["__doc_key"])
        key_to_led.setdefault(key, []).append(int(r["__idx"]))

    # Primary: amount + doc_key
    for _, r in f.iterrows():
        if r["__doc_key"]:
            key = (round(float(r["__amount"]), 2), r["__doc_key"])
            if key in key_to_led:
                try_match(int(r["__idx"]), key_to_led[key])

    # Secondary: amount + date (with tolerance)
    amt_to_led = {}
    for _, r in l.iterrows():
        amt_to_led.setdefault(round(float(r["__amount"]), 2), []).append(int(r["__idx"]))
    l_by_idx = l.set_index("__idx", drop=False)

    for _, r in f.iterrows():
        fi = int(r["__idx"])
        if fi in fin_match:
            continue
        amt = round(float(r["__amount"]), 2)
        if amt not in amt_to_led:
            continue
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

    # Outputs with status
    f_out = fin_df.copy()
    f_out["STATUS"] = f_out["CONCILIADO?"].map(lambda x: "✅ Conciliado" if x == "S" else "❌ Pendente")
    f_out["PAREADO_COM_IDX_CONTABIL"] = f["__idx"].map(lambda x: fin_match.get(int(x), ""))

    l_out = led_df.copy()
    l_out["STATUS"] = l_out["CONCILIADO?"].map(lambda x: "✅ Conciliado" if x == "S" else "❌ Pendente")
    l_out["PAREADO_COM_IDX_FINANCEIRO"] = l["__idx"].map(lambda x: led_match.get(int(x), ""))

    fin_only = f[~f["__idx"].astype(int).isin(fin_match.keys())].copy()
    led_only = l[~l["__idx"].astype(int).isin(led_match.keys())].copy()

def build_div_fin(fin_only_norm):
    # Divergências do Financeiro com campos "reconhecíveis"
    rows = []
    fin_by_idx = fin_df.reset_index(drop=True)

    for _, r in fin_only_norm.iterrows():
        i = int(r["__idx"])
        base = fin_by_idx.iloc[i] if 0 <= i < len(fin_by_idx) else pd.Series(dtype="object")

        rows.append({
            "ORIGEM": "Somente Financeiro",
            "DATA": r["__date"],
            "DOCUMENTO": (str(base.get(cfg.get("fin_documento"), "")) if cfg.get("fin_documento") else ""),
            "PREFIXO/TITULO": (str(base.get(cfg.get("fin_prefixo"), "")) if cfg.get("fin_prefixo") else ""),
            "OPERACAO/HISTORICO": (str(base.get(cfg.get("fin_operacao"), "")) if cfg.get("fin_operacao") else str(r["__text"]).strip()),
            "CHAVE_DOC": r["__doc_key"],
            "VALOR": round(float(r["__amount"]), 2),
            "SALDO_NA_LINHA": (round(float(r["__saldo"]), 2) if pd.notna(r["__saldo"]) else np.nan),
        })
    return pd.DataFrame(rows)

def build_div_led(led_only_norm):
    # Divergências do Contábil com campos "reconhecíveis"
    rows = []
    led_by_idx = led_df.reset_index(drop=True)

    for _, r in led_only_norm.iterrows():
        i = int(r["__idx"])
        base = led_by_idx.iloc[i] if 0 <= i < len(led_by_idx) else pd.Series(dtype="object")

        rows.append({
            "ORIGEM": "Somente Contábil",
            "DATA": r["__date"],
            "LOTE/SUB/DOC/LINHA": (str(base.get(cfg.get("led_doc"), "")) if cfg.get("led_doc") else ""),
            "CONTA": (str(base.get(cfg.get("led_conta"), "")) if cfg.get("led_conta") else ""),
            "HISTORICO": (str(base.get(cfg.get("led_historico"), "")) if cfg.get("led_historico") else str(r["__text"]).strip()),
            "CHAVE_DOC": r["__doc_key"],
            "VALOR": round(float(r["__amount"]), 2),
            "SALDO_NA_LINHA": (round(float(r["__saldo"]), 2) if pd.notna(r["__saldo"]) else np.nan),
        })
    return pd.DataFrame(rows)

div = pd.concat(
    [
        build_div_fin(fin_only),
        build_div_led(led_only),
    ],
    ignore_index=True
)

    # Summary numbers
    fin_total = round(float(f["__amount"].sum()), 2)
    led_total = round(float(l["__amount"].sum()), 2)
    fin_unmatched = round(float(fin_only["__amount"].sum()), 2)
    led_unmatched = round(float(led_only["__amount"].sum()), 2)

    saldo_ant_fin = compute_saldo_anterior(f)
    saldo_ant_led = compute_saldo_anterior(l)
    saldo_fin = compute_saldo_final(f)
    saldo_led = compute_saldo_final(l)

    diff_saldo_ant = (np.nan if (pd.isna(saldo_ant_fin) or pd.isna(saldo_ant_led)) else round(saldo_ant_fin - saldo_ant_led, 2))
    diff_final = (np.nan if (pd.isna(saldo_fin) or pd.isna(saldo_led)) else round(saldo_fin - saldo_led, 2))

    impacto = round(fin_unmatched - led_unmatched, 2)
    diff_esperada = (np.nan if pd.isna(diff_saldo_ant) else round(diff_saldo_ant + impacto, 2))
    conferencia = (np.nan if (pd.isna(diff_final) or pd.isna(diff_esperada)) else round(diff_final - diff_esperada, 2))

    resumo = {
        "Saldo anterior (antes do 1º movimento) - Financeiro": saldo_ant_fin,
        "Saldo anterior (antes do 1º movimento) - Contábil": saldo_ant_led,
        "Diferença de saldo anterior (Fin - Cont)": diff_saldo_ant,
        "Saldo final (último movimento) - Financeiro": saldo_fin,
        "Saldo final (último movimento) - Contábil": saldo_led,
        "Diferença final (Fin - Cont)": diff_final,
        "Soma pendente Somente Financeiro (Divergências)": fin_unmatched,
        "Soma pendente Somente Contábil (Divergências)": led_unmatched,
        "Impacto líquido pendentes (Fin - Cont)": impacto,
        "Diferença esperada (Dif. saldo anterior + Impacto)": diff_esperada,
        "Conferência (Dif. final - Dif. esperada) → precisa zerar": conferencia,

    }

    return f_out, l_out, div, resumo

def build_excel(fin_out, led_out, div, resumo):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        fin_out.to_excel(writer, index=False, sheet_name="Extrato_Financeiro")
        led_out.to_excel(writer, index=False, sheet_name="Razao_Contabil")
        div.to_excel(writer, index=False, sheet_name="Divergencias")
        pd.DataFrame(list(resumo.items()), columns=["Metrica", "Valor"]).to_excel(
            writer, index=False, sheet_name="Resumo_Fechamento"
        )

        wb = writer.book
        fmt_money = wb.add_format({"num_format": "#,##0.00"})
        fmt_hdr = wb.add_format({"bold": True})

        for sh in ["Extrato_Financeiro", "Razao_Contabil", "Divergencias", "Resumo_Fechamento"]:
            ws = writer.sheets[sh]
            ws.set_row(0, None, fmt_hdr)
            ws.freeze_panes(1, 0)

            if sh == "Resumo_Fechamento":
                ws.set_column(0, 0, 55)
                ws.set_column(1, 1, 18, fmt_money)
            else:
                ws.set_column(0, 0, 16)
                ws.set_column(1, 5, 28)
                ws.set_column(6, 30, 18, fmt_money)

    output.seek(0)
    return output

def fmt(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "-"
    try:
        return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

# ----------------------------
# UI
# ----------------------------
st.title("ConciliaMais — Módulo 1 (MVP)")
st.caption("Upload do Extrato Financeiro + Razão Contábil → Match automático → Divergências → Resumo de fechamento")

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
        st.markdown("### Extrato Financeiro")
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
        fin_amount = st.selectbox("OU coluna de Valor Único (opcional)", ["(usar Entradas - Saídas)"] + list(fin_df.columns),
                                  index=(["(usar Entradas - Saídas)"] + list(fin_df.columns)).index(fin_guess["valor"]) if fin_guess["valor"] in fin_df.columns else 0)
        fin_saldo = st.selectbox("Coluna de Saldo (opcional)", ["(nenhuma)"] + list(fin_df.columns),
                                 index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["saldo"]) if fin_guess["saldo"] in fin_df.columns else 0)

    with colB:
        st.markdown("### Razão Contábil")
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
        led_amount = st.selectbox("OU coluna de Valor Único (opcional)", ["(usar Débito - Crédito)"] + list(led_df.columns),
                                  index=(["(usar Débito - Crédito)"] + list(led_df.columns)).index(led_guess["valor"]) if led_guess["valor"] in led_df.columns else 0, key="led_amount")
        led_saldo = st.selectbox("Coluna de Saldo (opcional)", ["(nenhuma)"] + list(led_df.columns),
                                 index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["saldo"]) if led_guess["saldo"] in led_df.columns else 0, key="led_saldo")

    st.divider()
    st.subheader("Parâmetros de conciliação")
    p1, p2 = st.columns([1, 2])
    with p1:
        date_tol = st.number_input("Tolerância de dias (0 = mesma data)", min_value=0, max_value=10, value=0, step=1)
    with p2:
        st.caption("Matching: (1) Valor + DocKey, (2) Valor + Data (com tolerância). DocKey = sequência numérica (>=6 dígitos) do texto.")

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

    # Pre-check: saldo anterior mismatch warning
    f_norm, l_norm = build_normalized(fin_df, led_df, cfg)
    saldo_ant_fin = compute_saldo_anterior(f_norm)
    saldo_ant_led = compute_saldo_anterior(l_norm)
    diff_ant = (np.nan if (pd.isna(saldo_ant_fin) or pd.isna(saldo_ant_led)) else round(saldo_ant_fin - saldo_ant_led, 2))

    st.divider()
    st.subheader("Alertas iniciais")
    if pd.isna(diff_ant):
        st.info("Saldo anterior: não foi possível calcular automaticamente (verifique se selecionou a coluna de Saldo).")
        proceed_ok = True
    else:
        if abs(diff_ant) > 0.009:
            st.warning(f"Saldo anterior NÃO bate (Fin - Cont = {fmt(diff_ant)}). Isso indica diferença em períodos anteriores.")
            proceed_ok = st.checkbox("Quero prosseguir mesmo assim", value=False)
        else:
            st.success("Saldo anterior bate (OK).")
            proceed_ok = True

    run = st.button("Rodar conciliação agora", type="primary", disabled=not proceed_ok)

    if run:
        with st.spinner("Processando..."):
            fin_out, led_out, div, resumo = reconcile(fin_df, led_df, cfg, date_tol_days=int(date_tol))

        st.subheader("Resumo de fechamento (precisa zerar)")
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Dif. saldo anterior (Fin-Cont)", fmt(resumo["Diferença de saldo anterior (Fin - Cont)"]))
        r2.metric("Impacto pendentes (Fin-Cont)", fmt(resumo["Impacto líquido pendentes (Fin - Cont)"]))
        r3.metric("Dif. final (Fin-Cont)", fmt(resumo["Diferença final (Fin - Cont)"]))
        r4.metric("Conferência (ideal 0,00)", fmt(resumo["Conferência (Dif. final - Dif. esperada) → precisa zerar"]))

        ponte_tbl = pd.DataFrame([
            ["Saldo anterior (antes do 1º movimento) - Financeiro", resumo["Saldo anterior (antes do 1º movimento) - Financeiro"]],
            ["Saldo anterior (antes do 1º movimento) - Contábil", resumo["Saldo anterior (antes do 1º movimento) - Contábil"]],
            ["Diferença de saldo anterior (Fin - Cont)", resumo["Diferença de saldo anterior (Fin - Cont)"]],
            ["Saldo final (último movimento) - Financeiro", resumo["Saldo final (último movimento) - Financeiro"]],
            ["Saldo final (último movimento) - Contábil", resumo["Saldo final (último movimento) - Contábil"]],
            ["Diferença final (Fin - Cont)", resumo["Diferença final (Fin - Cont)"]],
            ["Soma pendente Somente Financeiro (Divergências)", resumo["Soma pendente Somente Financeiro (Divergências)"]],
            ["Soma pendente Somente Contábil (Divergências)", resumo["Soma pendente Somente Contábil (Divergências)"]],
            ["Impacto líquido pendentes (Fin - Cont)", resumo["Impacto líquido pendentes (Fin - Cont)"]],
            ["Diferença esperada (Dif. saldo anterior + Impacto)", resumo["Diferença esperada (Dif. saldo anterior + Impacto)"]],
            ["Conferência (Dif. final - Dif. esperada) → precisa zerar", resumo["Conferência (Dif. final - Dif. esperada) → precisa zerar"]],
        ], columns=["Métrica", "Valor"])

        st.dataframe(ponte_tbl, use_container_width=True, height=380)

        st.subheader("Divergências (itens não pareados)")
        st.dataframe(div, use_container_width=True, height=320)

        excel_bytes = build_excel(fin_out, led_out, div, resumo)
        st.download_button(
            "Baixar relatório Excel (ConciliaMais)",
            data=excel_bytes,
            file_name=f"ConciliaMais_Resultado_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

else:
    st.info("Faça o upload do Extrato Financeiro e do Razão Contábil para liberar o processamento.")
