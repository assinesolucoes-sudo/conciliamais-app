import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="ConciliaMais — Módulo 1", layout="wide")

# ----------------------------
# CSS (visual mais “produto”)
# ----------------------------
st.markdown(
    """
<style>
/* Layout geral */
.block-container { padding-top: 1.4rem; padding-bottom: 2rem; }
h1, h2, h3 { letter-spacing: -0.02em; }

/* Cards */
.cm-cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-top: 8px; }
.cm-card {
  border-radius: 16px;
  padding: 14px 14px 12px 14px;
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.08);
}
.cm-card .k { font-size: 12px; opacity: .80; margin-bottom: 6px; }
.cm-card .v { font-size: 22px; font-weight: 700; }
.cm-card .s { font-size: 12px; opacity: .75; margin-top: 6px; }
.cm-pill { display: inline-block; padding: 4px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; }
.cm-ok { background: rgba(34,197,94,.18); color: rgb(134,239,172); border: 1px solid rgba(34,197,94,.35); }
.cm-warn { background: rgba(245,158,11,.18); color: rgb(253,230,138); border: 1px solid rgba(245,158,11,.35); }
.cm-bad { background: rgba(239,68,68,.18); color: rgb(254,202,202); border: 1px solid rgba(239,68,68,.35); }

/* Barra de ações */
.cm-actions { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 10px; }

/* Separadores */
.cm-section { margin-top: 18px; }
</style>
""",
    unsafe_allow_html=True,
)

# ----------------------------
# Helpers
# ----------------------------
def _to_date_series(s):
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
    # pega a planilha mais “larga”
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
        doc=pick("lote/sub/doc/linha", "documento", "doc", "num documento", "lote"),
        conta=pick("conta"),
        valor=pick("valor", "vlr"),
    )

def build_normalized(fin_df, led_df, cfg):
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
        (f[op_col].astype(str) if op_col else "")
        + " "
        + (f[doc_col].astype(str) if doc_col else "")
        + " "
        + (f[pre_col].astype(str) if pre_col else "")
    ).astype(str)

    f["__doc_key"] = f["__text"].map(extract_doc_key)
    f["__idx"] = np.arange(len(f))

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
        (l[hist_col].astype(str) if hist_col else "")
        + " "
        + (l[doc_col2].astype(str) if doc_col2 else "")
        + " "
        + (l[conta_col].astype(str) if conta_col else "")
    ).astype(str)

    l["__doc_key"] = l["__text"].map(extract_doc_key)
    l["__idx"] = np.arange(len(l))
    return f, l

def compute_saldo_anterior(df_norm):
    dfv = df_norm.copy()
    dfv = dfv[dfv["__date"].notna()].copy()
    dfv = dfv[dfv["__amount"].notna()].copy()
    dfv = dfv[dfv["__saldo"].notna()].copy()
    if dfv.empty:
        return np.nan
    r = dfv.iloc[0]
    return round(float(r["__saldo"]) - float(r["__amount"]), 2)

def compute_saldo_final(df_norm):
    dfv = df_norm.copy()
    dfv = dfv[dfv["__date"].notna()].copy()
    dfv = dfv[dfv["__amount"].notna()].copy()
    dfv = dfv[dfv["__saldo"].notna()].copy()
    if dfv.empty:
        return np.nan
    r = dfv.iloc[-1]
    return round(float(r["__saldo"]), 2)

def fmt(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "-"
    try:
        return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

def status_pill(conferencia):
    if conferencia is None or (isinstance(conferencia, float) and np.isnan(conferencia)):
        return '<span class="cm-pill cm-warn">Conferência não calculada</span>'
    x = abs(float(conferencia))
    if x <= 0.01:
        return '<span class="cm-pill cm-ok">Fechou (0,00)</span>'
    if x <= 5:
        return '<span class="cm-pill cm-warn">Quase (verificar)</span>'
    return '<span class="cm-pill cm-bad">Não fechou</span>'

def reconcile(fin_df, led_df, cfg, date_tol_days=0):
    f, l = build_normalized(fin_df, led_df, cfg)

    ledger_used = set()
    fin_match = {}
    led_match = {}

    key_to_led = {}
    for _, r in l.iterrows():
        if pd.isna(r["__amount"]):
            continue
        amt = round(float(r["__amount"]), 2)
        key = (amt, r["__doc_key"])
        key_to_led.setdefault(key, []).append(int(r["__idx"]))

    def try_match(fin_idx, candidates):
        for li in candidates:
            if li in ledger_used:
                continue
            ledger_used.add(li)
            fin_match[fin_idx] = li
            led_match[li] = fin_idx
            return True
        return False

    # 1) amount + doc_key
    for _, r in f.iterrows():
        fi = int(r["__idx"])
        if r["__doc_key"] and pd.notna(r["__amount"]):
            key = (round(float(r["__amount"]), 2), r["__doc_key"])
            if key in key_to_led:
                try_match(fi, key_to_led[key])

    # 2) amount + date (tolerância)
    amt_to_led = {}
    for _, r in l.iterrows():
        if pd.isna(r["__amount"]):
            continue
        amt_to_led.setdefault(round(float(r["__amount"]), 2), []).append(int(r["__idx"]))
    l_by_idx = l.set_index("__idx", drop=False)

    for _, r in f.iterrows():
        fi = int(r["__idx"])
        if fi in fin_match or pd.isna(r["__amount"]) or pd.isna(r["__date"]):
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
            if pd.isna(ldate):
                continue
            if abs((pd.to_datetime(fdate) - pd.to_datetime(ldate)).days) <= int(date_tol_days):
                cands.append(li)
        if cands:
            try_match(fi, cands)

    fin_only = f[~f["__idx"].astype(int).isin(fin_match.keys())].copy()
    led_only = l[~l["__idx"].astype(int).isin(led_match.keys())].copy()

    fin_unmatched = round(float(fin_only["__amount"].sum()), 2) if not fin_only.empty else 0.0
    led_unmatched = round(float(led_only["__amount"].sum()), 2) if not led_only.empty else 0.0

    saldo_ant_fin = compute_saldo_anterior(f)
    saldo_ant_led = compute_saldo_anterior(l)
    saldo_fin = compute_saldo_final(f)
    saldo_led = compute_saldo_final(l)

    diff_saldo_ant = np.nan if (pd.isna(saldo_ant_fin) or pd.isna(saldo_ant_led)) else round(saldo_ant_fin - saldo_ant_led, 2)
    diff_final = np.nan if (pd.isna(saldo_fin) or pd.isna(saldo_led)) else round(saldo_fin - saldo_led, 2)

    impacto = round(fin_unmatched - led_unmatched, 2)
    diff_esperada = np.nan if pd.isna(diff_saldo_ant) else round(diff_saldo_ant + impacto, 2)
    conferencia = np.nan if (pd.isna(diff_final) or pd.isna(diff_esperada)) else round(diff_final - diff_esperada, 2)

    # Divergências humanizadas (duas “visões”)
    fin_rows = []
    fin_reset = fin_df.reset_index(drop=True)
    for _, r in fin_only.iterrows():
        i = int(r["__idx"])
        base = fin_reset.iloc[i] if 0 <= i < len(fin_reset) else pd.Series(dtype="object")
        fin_rows.append({
            "ORIGEM": "Somente Financeiro",
            "DATA": r["__date"],
            "DOCUMENTO": str(base.get(cfg.get("fin_documento"), "")) if cfg.get("fin_documento") else "",
            "PREFIXO_TITULO": str(base.get(cfg.get("fin_prefixo"), "")) if cfg.get("fin_prefixo") else "",
            "HISTORICO_OPERACAO": str(base.get(cfg.get("fin_operacao"), "")) if cfg.get("fin_operacao") else str(r["__text"]),
            "CHAVE_DOC": r["__doc_key"],
            "VALOR": round(float(r["__amount"]), 2) if pd.notna(r["__amount"]) else np.nan,
        })

    led_rows = []
    led_reset = led_df.reset_index(drop=True)
    for _, r in led_only.iterrows():
        i = int(r["__idx"])
        base = led_reset.iloc[i] if 0 <= i < len(led_reset) else pd.Series(dtype="object")
        led_rows.append({
            "ORIGEM": "Somente Contábil",
            "DATA": r["__date"],
            "DOCUMENTO": str(base.get(cfg.get("led_doc"), "")) if cfg.get("led_doc") else "",
            "CONTA": str(base.get(cfg.get("led_conta"), "")) if cfg.get("led_conta") else "",
            "HISTORICO_OPERACAO": str(base.get(cfg.get("led_historico"), "")) if cfg.get("led_historico") else str(r["__text"]),
            "CHAVE_DOC": r["__doc_key"],
            "VALOR": round(float(r["__amount"]), 2) if pd.notna(r["__amount"]) else np.nan,
        })

    div = pd.concat([pd.DataFrame(fin_rows), pd.DataFrame(led_rows)], ignore_index=True)

    stats = {
        "fin_total": int(len(f)),
        "fin_conc": int(len(fin_match)),
        "fin_pend": int(len(f) - len(fin_match)),
        "fin_pend_val": float(fin_unmatched),
        "led_pend_val": float(led_unmatched),
        "impacto": float(impacto),
        "diff_saldo_ant": float(diff_saldo_ant) if pd.notna(diff_saldo_ant) else np.nan,
        "diff_final": float(diff_final) if pd.notna(diff_final) else np.nan,
        "conferencia": float(conferencia) if pd.notna(conferencia) else np.nan,
    }

    return div, stats

def build_tratativa(div):
    if div is None or div.empty:
        return pd.DataFrame(columns=[
            "ORIGEM","DATA","IDENTIFICADOR","HISTORICO_OPERACAO","VALOR",
            "ACAO_SUGERIDA","RESPONSAVEL","PRAZO","STATUS","OBS"
        ])
    def ident(row):
        parts = []
        doc = str(row.get("DOCUMENTO","")).strip()
        pre = str(row.get("PREFIXO_TITULO","")).strip()
        cta = str(row.get("CONTA","")).strip()
        if doc: parts.append(f"DOC:{doc}")
        if pre: parts.append(f"PRE:{pre}")
        if cta: parts.append(f"CTA:{cta}")
        return " | ".join(parts)[:140]
    t = pd.DataFrame({
        "ORIGEM": div["ORIGEM"],
        "DATA": div["DATA"],
        "IDENTIFICADOR": div.apply(ident, axis=1),
        "HISTORICO_OPERACAO": div.get("HISTORICO_OPERACAO",""),
        "VALOR": div["VALOR"],
        "ACAO_SUGERIDA": "",
        "RESPONSAVEL": "",
        "PRAZO": "",
        "STATUS": "Pendente",
        "OBS": "",
    })
    return t

def to_excel_package(div, stats):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        # Resumo limpo (só o essencial)
        resumo = pd.DataFrame([
            ["Diferença saldo anterior (Fin - Cont)", stats["diff_saldo_ant"]],
            ["Impacto pendentes (Fin - Cont)", stats["impacto"]],
            ["Diferença final (Fin - Cont)", stats["diff_final"]],
            ["Conferência (ideal 0,00)", stats["conferencia"]],
        ], columns=["Metrica","Valor"])
        resumo.to_excel(w, index=False, sheet_name="Resumo")

        # Divergências
        div.to_excel(w, index=False, sheet_name="Divergencias")

        # Tratativa
        trat = build_tratativa(div)
        trat.to_excel(w, index=False, sheet_name="Tratativa")

        # Formatação mínima (não “enfeitar” demais)
        wb = w.book
        fmt_hdr = wb.add_format({"bold": True, "border": 1, "align": "center", "valign": "vcenter"})
        fmt_money = wb.add_format({"num_format": "#,##0.00", "border": 1})

        for sh in ["Resumo","Divergencias","Tratativa"]:
            ws = w.sheets[sh]
            ws.freeze_panes(1, 0)
            ws.set_row(0, 20, fmt_hdr)

        w.sheets["Resumo"].set_column(0, 0, 44)
        w.sheets["Resumo"].set_column(1, 1, 22, fmt_money)

        w.sheets["Divergencias"].set_column(0, 0, 18)
        w.sheets["Divergencias"].set_column(1, 1, 12)
        w.sheets["Divergencias"].set_column(2, 4, 38)
        w.sheets["Divergencias"].set_column(5, 5, 18)
        w.sheets["Divergencias"].set_column(6, 6, 16, fmt_money)

        w.sheets["Tratativa"].set_column(0, 0, 18)
        w.sheets["Tratativa"].set_column(1, 1, 12)
        w.sheets["Tratativa"].set_column(2, 3, 48)
        w.sheets["Tratativa"].set_column(4, 4, 16, fmt_money)
        w.sheets["Tratativa"].set_column(5, 9, 22)

    out.seek(0)
    return out

# ----------------------------
# Navegação simples por estado
# ----------------------------
if "page" not in st.session_state:
    st.session_state.page = "upload"
if "results" not in st.session_state:
    st.session_state.results = None

# ----------------------------
# Página: Upload
# ----------------------------
if st.session_state.page == "upload":
    st.title("ConciliaMais — Módulo 1")
    st.caption("Upload do Extrato Financeiro + Razão Contábil → Match automático → Divergências → Painel de fechamento")

    c1, c2 = st.columns(2)
    with c1:
        fin_file = st.file_uploader("Extrato Financeiro (.xlsx ou .csv)", type=["xlsx","csv"], key="fin")
    with c2:
        led_file = st.file_uploader("Razão Contábil (.xlsx ou .csv)", type=["xlsx","csv"], key="led")

    if not fin_file or not led_file:
        st.info("Faça o upload dos dois arquivos para liberar o processamento.")
        st.stop()

    fin_df = read_table(fin_file)
    led_df = read_table(led_file)

    fin_guess = auto_detect_financial(fin_df)
    led_guess = auto_detect_ledger(led_df)

    st.markdown("### Mapeamento de colunas (auto-detectado — ajuste se precisar)")
    a, b = st.columns(2)

    with a:
        st.markdown("#### Extrato Financeiro")
        fin_date = st.selectbox("Data", fin_df.columns,
            index=fin_df.columns.get_loc(fin_guess["date"]) if fin_guess["date"] in fin_df.columns else 0)

        fin_operacao = st.selectbox("Operação/Histórico", ["(nenhuma)"] + list(fin_df.columns),
            index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["operacao"]) if fin_guess["operacao"] in fin_df.columns else 0)

        fin_documento = st.selectbox("Documento", ["(nenhuma)"] + list(fin_df.columns),
            index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["documento"]) if fin_guess["documento"] in fin_df.columns else 0)

        fin_prefixo = st.selectbox("Prefixo/Título", ["(nenhuma)"] + list(fin_df.columns),
            index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["prefixo"]) if fin_guess["prefixo"] in fin_df.columns else 0)

        fin_entradas = st.selectbox("Entradas", ["(nenhuma)"] + list(fin_df.columns),
            index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["entradas"]) if fin_guess["entradas"] in fin_df.columns else 0)

        fin_saidas = st.selectbox("Saídas", ["(nenhuma)"] + list(fin_df.columns),
            index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["saidas"]) if fin_guess["saidas"] in fin_df.columns else 0)

        fin_amount = st.selectbox("OU Valor Único", ["(usar Entradas - Saídas)"] + list(fin_df.columns),
            index=(["(usar Entradas - Saídas)"] + list(fin_df.columns)).index(fin_guess["valor"]) if fin_guess["valor"] in fin_df.columns else 0)

        fin_saldo = st.selectbox("Saldo", ["(nenhuma)"] + list(fin_df.columns),
            index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["saldo"]) if fin_guess["saldo"] in fin_df.columns else 0)

    with b:
        st.markdown("#### Razão Contábil")
        led_date = st.selectbox("Data", led_df.columns,
            index=led_df.columns.get_loc(led_guess["date"]) if led_guess["date"] in led_df.columns else 0, key="ld")

        led_historico = st.selectbox("Histórico", ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["historico"]) if led_guess["historico"] in led_df.columns else 0, key="lh")

        led_doc = st.selectbox("Documento/Lote", ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["doc"]) if led_guess["doc"] in led_df.columns else 0, key="ldoc")

        led_conta = st.selectbox("Conta", ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["conta"]) if led_guess["conta"] in led_df.columns else 0, key="lcta")

        led_debito = st.selectbox("Débito", ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["debito"]) if led_guess["debito"] in led_df.columns else 0, key="ldb")

        led_credito = st.selectbox("Crédito", ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["credito"]) if led_guess["credito"] in led_df.columns else 0, key="lcr")

        led_amount = st.selectbox("OU Valor Único", ["(usar Débito - Crédito)"] + list(led_df.columns),
            index=(["(usar Débito - Crédito)"] + list(led_df.columns)).index(led_guess["valor"]) if led_guess["valor"] in led_df.columns else 0, key="lamt")

        led_saldo = st.selectbox("Saldo", ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["saldo"]) if led_guess["saldo"] in led_df.columns else 0, key="ls")

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

    st.markdown("### Alertas iniciais (Saldo anterior)")
    f_norm, l_norm = build_normalized(fin_df, led_df, cfg)
    saldo_ant_fin = compute_saldo_anterior(f_norm)
    saldo_ant_led = compute_saldo_anterior(l_norm)
    diff_ant = np.nan if (pd.isna(saldo_ant_fin) or pd.isna(saldo_ant_led)) else round(saldo_ant_fin - saldo_ant_led, 2)

    proceed_ok = True
    if pd.isna(diff_ant):
        st.info("Não foi possível calcular saldo anterior automaticamente. Selecione a coluna de saldo em ambos os arquivos, se existir.")
    else:
        if abs(diff_ant) > 0.01:
            st.warning(f"Saldo anterior não bate (Fin - Cont = {fmt(diff_ant)}). Diferença pode estar em períodos anteriores.")
            proceed_ok = st.checkbox("Prosseguir mesmo assim", value=False)
        else:
            st.success("Saldo anterior bate (OK).")

    date_tol = st.number_input("Tolerância de dias para match por data (0 = mesma data)", min_value=0, max_value=10, value=0, step=1)

    if st.button("Processar e ir para Resultados", type="primary", disabled=not proceed_ok):
        with st.spinner("Processando..."):
            div, stats = reconcile(fin_df, led_df, cfg, date_tol_days=int(date_tol))

        st.session_state.results = {
            "div": div,
            "stats": stats,
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        st.session_state.page = "resultados"
        st.rerun()

# ----------------------------
# Página: Resultados
# ----------------------------
else:
    res = st.session_state.results
    if not res:
        st.session_state.page = "upload"
        st.rerun()

    div = res["div"]
    stats = res["stats"]

    st.title("Resultados — ConciliaMais (Módulo 1)")
    st.caption(f"Processado em: {res['generated_at']}")

    # Cards
    st.markdown(
        f"""
<div class="cm-cards">
  <div class="cm-card">
    <div class="k">Conciliadas (Financeiro)</div>
    <div class="v">{stats["fin_conc"]} / {stats["fin_total"]}</div>
    <div class="s">{(stats["fin_conc"]/max(stats["fin_total"],1))*100:.1f}% conciliado</div>
  </div>
  <div class="cm-card">
    <div class="k">Pendentes (Financeiro)</div>
    <div class="v">{stats["fin_pend"]}</div>
    <div class="s">itens não pareados</div>
  </div>
  <div class="cm-card">
    <div class="k">Impacto pendentes (Fin - Cont)</div>
    <div class="v">{fmt(stats["impacto"])}</div>
    <div class="s">somente financeiro - somente contábil</div>
  </div>
  <div class="cm-card">
    <div class="k">Conferência (ideal 0,00)</div>
    <div class="v">{fmt(stats["conferencia"])}</div>
    <div class="s">{status_pill(stats["conferencia"])}</div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    st.markdown('<div class="cm-section"></div>', unsafe_allow_html=True)

    # Gráficos simples
    g1, g2 = st.columns(2)
    with g1:
        st.markdown("### Match vs Pendente (quantidade)")
        chart_df = pd.DataFrame(
            {"Quantidade": [stats["fin_conc"], stats["fin_pend"]]},
            index=["Conciliado", "Pendente"],
        )
        st.bar_chart(chart_df)

    with g2:
        st.markdown("### Pendências (valor)")
        chart2_df = pd.DataFrame(
            {"Valor": [stats["fin_pend_val"], stats["led_pend_val"]]},
            index=["Somente Financeiro", "Somente Contábil"],
        )
        st.bar_chart(chart2_df)

    # Filtros + tabela
    st.markdown("### Divergências (itens não pareados)")
    fcol1, fcol2, fcol3 = st.columns([1, 1, 2])

    with fcol1:
        origem = st.selectbox("Filtrar por origem", ["Todas", "Somente Financeiro", "Somente Contábil"])
    with fcol2:
        ordenar = st.selectbox("Ordenar por", ["DATA", "VALOR"])
    with fcol3:
        busca = st.text_input("Buscar (documento, histórico, chave)", value="")

    df = div.copy()
    if origem != "Todas":
        df = df[df["ORIGEM"] == origem].copy()

    if busca.strip():
        q = busca.strip().lower()
        cols_search = [c for c in df.columns if c in ["DOCUMENTO", "PREFIXO_TITULO", "CONTA", "HISTORICO_OPERACAO", "CHAVE_DOC"]]
        mask = False
        for c in cols_search:
            mask = mask | df[c].astype(str).str.lower().str.contains(q, na=False)
        df = df[mask].copy()

    if ordenar in df.columns:
        df = df.sort_values(by=ordenar, ascending=True)

    st.dataframe(df, use_container_width=True, height=420)

    # Exportar (separado)
    st.markdown("### Exportar")
    st.markdown('<div class="cm-actions">', unsafe_allow_html=True)

    fin_only = div[div["ORIGEM"] == "Somente Financeiro"].copy()
    led_only = div[div["ORIGEM"] == "Somente Contábil"].copy()
    trat = build_tratativa(div)

    st.download_button(
        "Baixar Somente Financeiro (CSV)",
        data=fin_only.to_csv(index=False).encode("utf-8"),
        file_name=f"ConciliaMais_SomenteFinanceiro_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv",
    )
    st.download_button(
        "Baixar Somente Contábil (CSV)",
        data=led_only.to_csv(index=False).encode("utf-8"),
        file_name=f"ConciliaMais_SomenteContabil_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv",
    )
    st.download_button(
        "Baixar Tratativa (CSV)",
        data=trat.to_csv(index=False).encode("utf-8"),
        file_name=f"ConciliaMais_Tratativa_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv",
    )

    excel_bytes = to_excel_package(div, stats)
    st.download_button(
        "Baixar Pacote Excel (Resumo + Divergências + Tratativa)",
        data=excel_bytes,
        file_name=f"ConciliaMais_Pacote_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")
    cback, creset = st.columns([1, 2])
    with cback:
        if st.button("Voltar para Upload"):
            st.session_state.page = "upload"
            st.rerun()
    with creset:
        if st.button("Limpar resultado e recomeçar"):
            st.session_state.results = None
            st.session_state.page = "upload"
            st.rerun()
