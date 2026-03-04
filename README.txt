import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime

# PDF (Relatório Resumo)
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

st.set_page_config(page_title="ConciliaMais — Conferência de Extrato Bancário", layout="wide")

# ----------------------------
# CSS
# ----------------------------
st.markdown(
    """
<style>
.block-container { padding-top: 1.1rem; padding-bottom: 2rem; }
h1, h2, h3 { letter-spacing: -0.02em; }

.cm-cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-top: 8px; }
.cm-card {
  border-radius: 16px;
  padding: 14px 14px 12px 14px;
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.08);
}
.cm-card .k { font-size: 12px; opacity: .80; margin-bottom: 6px; }
.cm-card .v { font-size: 22px; font-weight: 800; }
.cm-card .s { font-size: 12px; opacity: .75; margin-top: 6px; }

.cm-pill { display: inline-block; padding: 4px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; }
.cm-ok { background: rgba(34,197,94,.18); color: rgb(134,239,172); border: 1px solid rgba(34,197,94,.35); }
.cm-warn { background: rgba(245,158,11,.18); color: rgb(253,230,138); border: 1px solid rgba(245,158,11,.35); }
.cm-bad { background: rgba(239,68,68,.18); color: rgb(254,202,202); border: 1px solid rgba(239,68,68,.35); }

.cm-section { margin-top: 18px; }
.cm-help { opacity: .78; font-size: 13px; margin-top: -6px; }

.cm-mini {
  border-radius: 14px;
  padding: 10px 12px;
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.10);
  text-align: right;
}
.cm-mini .k { font-size: 12px; opacity: .80; margin-bottom: 4px; }
.cm-mini .v { font-size: 20px; font-weight: 900; letter-spacing: -0.01em; }

.cm-detail {
  border-radius: 16px;
  padding: 14px;
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.10);
}
.cm-detail .title { font-size: 14px; font-weight: 900; margin-bottom: 10px; }
.cm-detail .row { font-size: 13px; opacity: .92; margin: 4px 0; }
.cm-detail .label { opacity: .70; }
.cm-detail .val { font-weight: 650; }
</style>
""",
    unsafe_allow_html=True,
)

# ----------------------------
# Helpers (motor)
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

def pill_calculo(conferencia):
    if conferencia is None or (isinstance(conferencia, float) and np.isnan(conferencia)):
        return '<span class="cm-pill cm-warn">Não calculada</span>'
    x = abs(float(conferencia))
    if x <= 0.01:
        return '<span class="cm-pill cm-ok">Consistente (0,00)</span>'
    if x <= 5:
        return '<span class="cm-pill cm-warn">Quase (verificar)</span>'
    return '<span class="cm-pill cm-bad">Inconsistente</span>'

def extract_doc_from_ledger_history(x):
    if pd.isna(x):
        return ""
    t = str(x)
    m = re.search(r"/\s*(\d{6,})", t)
    if m:
        return m.group(1)
    m = re.search(r"\bNF[:\s-]*\s*(\d{6,})\b", t, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    nums = re.findall(r"\d{6,}", t)
    if nums:
        return nums[-1]
    return ""

def suggest_nucleo_motivo(row):
    origem = str(row.get("ORIGEM", "")).lower()
    hist = str(row.get("HISTORICO_OPERACAO", "")).lower()
    doc = str(row.get("DOCUMENTO","")).strip()

    if any(k in hist for k in ["cancelamento de baixa", "canc baixa", "estorno de baixa", "estorno baixa", "canc. baixa"]):
        return ("Processo interno", "Cancelamento/estorno de baixa — possível estorno sem confirmar contabilização")
    if any(k in hist for k in ["baixa", "liquidação", "liquidacao", "pagamento", "pagto", "estorno"]):
        return ("Processo interno", "Movimento de baixa/estorno — revisar execução completa do processo")

    if "somente financeiro" in origem and (doc != "" or "mov" in hist):
        return ("Cadastro", "Financeiro sem contabilização — revisar natureza/parametrização contábil")

    if any(k in hist for k in ["rp", "reprocess", "rotina", "processamento", "integracao", "integração"]):
        return ("Configuração RP", "Possível falha/ausência de rotina (RP) — revisar parametrização e execução")

    return ("Não identificado", "Não identificado — preencher motivo confirmado")

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
        hist_val = str(base.get(cfg.get("led_historico"), "")) if cfg.get("led_historico") else str(r["__text"])
        doc_from_hist = extract_doc_from_ledger_history(hist_val)
        led_rows.append({
            "ORIGEM": "Somente Contábil",
            "DATA": r["__date"],
            "DOCUMENTO": doc_from_hist,
            "PREFIXO_TITULO": "",
            "HISTORICO_OPERACAO": hist_val,
            "CHAVE_DOC": r["__doc_key"],
            "VALOR": round(float(r["__amount"]), 2) if pd.notna(r["__amount"]) else np.nan,
        })

    div = pd.concat([pd.DataFrame(fin_rows), pd.DataFrame(led_rows)], ignore_index=True)

    stats = {
        "fin_total": int(len(f)),
        "led_total": int(len(l)),
        "fin_conc": int(len(fin_match)),
        "fin_pend": int(len(f) - len(fin_match)),
        "led_pend": int(len(l) - len(led_match)),
        "fin_pend_val": float(fin_unmatched),
        "led_pend_val": float(led_unmatched),
        "impacto": float(impacto),
        "diff_saldo_ant": float(diff_saldo_ant) if pd.notna(diff_saldo_ant) else np.nan,
        "diff_final": float(diff_final) if pd.notna(diff_final) else np.nan,
        "diff_esperada": float(diff_esperada) if pd.notna(diff_esperada) else np.nan,
        "saldo_ant_fin": float(saldo_ant_fin) if pd.notna(saldo_ant_fin) else np.nan,
        "saldo_ant_led": float(saldo_ant_led) if pd.notna(saldo_ant_led) else np.nan,
        "saldo_fin": float(saldo_fin) if pd.notna(saldo_fin) else np.nan,
        "saldo_led": float(saldo_led) if pd.notna(saldo_led) else np.nan,
        "conferencia": float(conferencia) if pd.notna(conferencia) else np.nan,
    }

    return div, stats

# ----------------------------
# Excel: igual ao filtro (formatado)
# ----------------------------
def _autofit_worksheet(ws, df, start_col, max_width=75, min_width=10):
    for j, col in enumerate(df.columns):
        ser = df[col].astype(str).fillna("")
        sample = ser.head(250).tolist()
        max_len = max([len(str(col))] + [len(s) for s in sample]) if sample else len(str(col))
        width = max(min_width, min(max_width, max_len + 2))
        ws.set_column(start_col + j, start_col + j, width)

def to_excel_divergencias_filtradas(df_filtrado, total_filtrado, total_aberto, filtros, stats, generated_at):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        wb = w.book
        fmt_title = wb.add_format({"bold": True, "font_size": 14})
        fmt_k = wb.add_format({"bold": True, "font_size": 10, "font_color": "#666666"})
        fmt_txt = wb.add_format({"border": 1})
        fmt_hdr = wb.add_format({"bold": True, "border": 1, "align": "center", "valign": "vcenter"})
        fmt_date = wb.add_format({"num_format": "dd/mm/yyyy", "border": 1})
        fmt_money = wb.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00', "border": 1})
        fmt_money_big = wb.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00', "bold": True})

        sh = "Divergencias"
        df = df_filtrado.copy()

        if "DATA" in df.columns:
            df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
        if "VALOR" in df.columns:
            df["VALOR"] = df["VALOR"].map(normalize_money)

        start_row_table = 8
        df.to_excel(w, index=True, sheet_name=sh, startrow=start_row_table)
        ws = w.sheets[sh]

        ws.write(0, 0, "ConciliaMais — Divergências (Excel igual à tela)", fmt_title)
        ws.write(1, 0, "Processado em:", fmt_k)
        ws.write(1, 1, generated_at)

        ws.write(2, 0, "Origem:", fmt_k)
        ws.write(2, 1, filtros.get("origem", "Todas"))
        ws.write(3, 0, "Visualização:", fmt_k)
        ws.write(3, 1, filtros.get("ver", "Todas"))
        ws.write(4, 0, "Busca:", fmt_k)
        ws.write(4, 1, filtros.get("busca", ""))

        ws.write(2, 4, "Total do filtro:", fmt_k)
        ws.write_number(2, 5, float(total_filtrado or 0.0), fmt_money_big)
        ws.write(3, 4, "Total em aberto:", fmt_k)
        ws.write_number(3, 5, float(total_aberto or 0.0), fmt_money_big)

        ws.set_row(start_row_table, 20, fmt_hdr)
        ws.freeze_panes(start_row_table + 1, 0)

        nrows = len(df)
        ncols = len(df.columns) + 1

        if nrows > 0 and ncols > 0:
            table_last_row = start_row_table + nrows
            table_last_col = ncols - 1
            columns = [{"header": "ID"}] + [{"header": c} for c in df.columns]
            ws.add_table(
                start_row_table, 0, table_last_row, table_last_col,
                {"style": "Table Style Medium 9", "columns": columns, "autofilter": True}
            )

        col_idx = {c: i+1 for i, c in enumerate(df.columns)}
        for r in range(nrows):
            excel_r = start_row_table + 1 + r
            ws.write_number(excel_r, 0, int(df.index[r]), fmt_txt)
            for c, j in col_idx.items():
                val = df.iloc[r, j-1]
                if c == "VALOR":
                    if pd.notna(val):
                        ws.write_number(excel_r, j, float(val), fmt_money)
                    else:
                        ws.write_blank(excel_r, j, None, fmt_money)
                elif c == "DATA":
                    if pd.notna(val):
                        ws.write_datetime(excel_r, j, val.to_pydatetime(), fmt_date)
                    else:
                        ws.write_blank(excel_r, j, None, fmt_date)
                else:
                    ws.write(excel_r, j, "" if pd.isna(val) else str(val), fmt_txt)

        _autofit_worksheet(ws, pd.concat([pd.Series(df.index, name="ID"), df.reset_index(drop=True)], axis=1), start_col=0)

        resumo = pd.DataFrame([
            ["Saldo anterior – Financeiro", stats.get("saldo_ant_fin", np.nan)],
            ["Saldo anterior – Contábil", stats.get("saldo_ant_led", np.nan)],
            ["Diferença saldo anterior (Fin - Cont)", stats.get("diff_saldo_ant", np.nan)],
            ["", ""],
            ["Saldo final – Financeiro", stats.get("saldo_fin", np.nan)],
            ["Saldo final – Contábil", stats.get("saldo_led", np.nan)],
            ["Diferença final (Fin - Cont)", stats.get("diff_final", np.nan)],
            ["", ""],
            ["Soma só Financeiro (divergências)", stats.get("fin_pend_val", 0.0)],
            ["Soma só Contábil (divergências)", stats.get("led_pend_val", 0.0)],
            ["Impacto líquido (Fin - Cont)", stats.get("impacto", 0.0)],
            ["Diferença esperada", stats.get("diff_esperada", np.nan)],
            ["Conferência do cálculo", stats.get("conferencia", np.nan)],
        ], columns=["Métrica", "Valor"])
        resumo.to_excel(w, index=False, sheet_name="Resumo")
        ws3 = w.sheets["Resumo"]
        ws3.freeze_panes(1, 0)
        ws3.set_row(0, 20, fmt_hdr)
        ws3.set_column(0, 0, 52)
        ws3.set_column(1, 1, 26, wb.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00'}))

    out.seek(0)
    return out

# ----------------------------
# PDF Resumo
# ----------------------------
def to_pdf_resumo(stats, generated_at, div_master):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("Relatório de Conciliação Bancária — ConciliaMais (Módulo 1)", styles["Title"]))
    story.append(Spacer(1, 6))
    story.append(Paragraph(f"Processado em: {generated_at}", styles["Normal"]))
    story.append(Spacer(1, 14))

    df = div_master.copy()
    df["VALOR"] = df["VALOR"].map(normalize_money)
    df = df[df["VALOR"].notna()].copy()

    resolved = df["RESOLVIDO"].fillna(False) | (df["STATUS"].astype(str).str.lower().eq("resolvido"))
    df["__RES"] = resolved

    total_itens = len(df)
    itens_res = int(df["__RES"].sum())
    itens_ab = int(total_itens - itens_res)
    val_ab = float(df.loc[~df["__RES"], "VALOR"].sum()) if total_itens else 0.0
    pct_res = (itens_res / total_itens * 100.0) if total_itens else 0.0

    story.append(Paragraph("Resumo executivo", styles["Heading2"]))
    story.append(Spacer(1, 6))

    kpi_data = [
        ["Indicador", "Valor"],
        ["Diferenças encontradas (itens)", f"{total_itens}"],
        ["Diferenças resolvidas (itens)", f"{itens_res} ({pct_res:.1f}%)"],
        ["Pendências em aberto (itens)", f"{itens_ab}"],
        ["Pendências em aberto (valor)", fmt(val_ab)],
        ["Conferência do cálculo (ideal 0,00)", fmt(stats.get("conferencia", np.nan))],
    ]
    t1 = Table(kpi_data, colWidths=[240, 260])
    t1.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#111827")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#D1D5DB")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F9FAFB")]),
    ]))
    story.append(t1)
    story.append(Spacer(1, 14))

    story.append(Paragraph("Composição de saldos", styles["Heading2"]))
    story.append(Spacer(1, 6))
    comp = [
        ["Composição de saldos", "Valor"],
        ["Saldo anterior – Financeiro", fmt(stats.get("saldo_ant_fin", np.nan))],
        ["Saldo anterior – Contábil", fmt(stats.get("saldo_ant_led", np.nan))],
        ["Diferença saldo anterior (Fin - Cont)", fmt(stats.get("diff_saldo_ant", np.nan))],
        ["", ""],
        ["Saldo final – Financeiro", fmt(stats.get("saldo_fin", np.nan))],
        ["Saldo final – Contábil", fmt(stats.get("saldo_led", np.nan))],
        ["Diferença final (Fin - Cont)", fmt(stats.get("diff_final", np.nan))],
        ["", ""],
        ["Impacto líquido (Fin - Cont)", fmt(stats.get("impacto", 0.0))],
        ["Diferença esperada", fmt(stats.get("diff_esperada", np.nan))],
        ["Conferência do cálculo", fmt(stats.get("conferencia", np.nan))],
    ]
    t2 = Table(comp, colWidths=[340, 160])
    t2.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#111827")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#D1D5DB")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F9FAFB")]),
    ]))
    story.append(t2)
    story.append(Spacer(1, 14))

    story.append(Paragraph("Distribuição por núcleo (confirmado)", styles["Heading2"]))
    story.append(Spacer(1, 6))

    nuc = df.copy()
    nuc["NUCLEO_CONFIRMADO"] = nuc["NUCLEO_CONFIRMADO"].fillna("Não identificado").replace("", "Não identificado")
    dist = nuc.groupby("NUCLEO_CONFIRMADO", dropna=False).agg(
        Itens=("VALOR", "size"),
        Valor=("VALOR", "sum"),
        Abertos=("__RES", lambda x: int((~x).sum()))
    ).reset_index().sort_values("Valor", ascending=False)

    dist_rows = [["Núcleo", "Itens", "Abertos", "Valor"]]
    for _, r in dist.iterrows():
        dist_rows.append([str(r["NUCLEO_CONFIRMADO"]), str(int(r["Itens"])), str(int(r["Abertos"])), fmt(float(r["Valor"]))])

    t4 = Table(dist_rows, colWidths=[220, 60, 60, 160])
    t4.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#111827")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#D1D5DB")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F9FAFB")]),
    ]))
    story.append(t4)
    story.append(Spacer(1, 14))

    concl = "Conclusão: Pendências em aberto — revisar itens e concluir tratativa."
    try:
        if abs(val_ab) <= 0.01:
            concl = "Conclusão: Tratativa concluída — pendências em aberto zeradas (0,00)."
            conf = stats.get("conferencia", np.nan)
            if pd.notna(conf) and abs(float(conf)) <= 0.01:
                concl += " Fechamento do cálculo consistente (0,00)."
            else:
                concl += " Atenção: fechamento do cálculo ainda requer validação."
    except Exception:
        pass

    story.append(Paragraph("Conclusão", styles["Heading2"]))
    story.append(Spacer(1, 6))
    story.append(Paragraph(concl, styles["Normal"]))

    doc.build(story)
    buf.seek(0)
    return buf

# ----------------------------
# State
# ----------------------------
if "page" not in st.session_state:
    st.session_state.page = "upload"
if "results" not in st.session_state:
    st.session_state.results = None
if "div_master" not in st.session_state:
    st.session_state.div_master = None

NUCLEOS = ["Processo interno", "Cadastro", "Configuração RP", "Não identificado"]
STATUS_OPTS = ["Pendente", "Em análise", "Resolvido"]

# ----------------------------
# Página: Upload
# ----------------------------
if st.session_state.page == "upload":
    st.title("ConciliaMais — Conferência de Extrato Bancário")
    st.caption("Extrato Financeiro + Razão Contábil → Match automático → Divergências → Tratativa (check Resolvido)")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Extrato Financeiro")
        st.markdown('<div class="cm-help">Faça o upload da planilha do Extrato Financeiro.</div>', unsafe_allow_html=True)
        fin_file = st.file_uploader("Upload do Extrato Financeiro (.xlsx ou .csv)", type=["xlsx","csv"], key="fin")
    with c2:
        st.subheader("Razão Contábil")
        st.markdown('<div class="cm-help">Faça o upload da planilha do Razão Contábil.</div>', unsafe_allow_html=True)
        led_file = st.file_uploader("Upload do Razão Contábil (.xlsx ou .csv)", type=["xlsx","csv"], key="led")

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

        fin_prefixo = st.selectbox("Prefixo/Título (usaremos como DOCUMENTO na divergência)", ["(nenhuma)"] + list(fin_df.columns),
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

        led_doc = st.selectbox("Documento/Lote (não será usado como documento final)", ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["doc"]) if led_guess["doc"] in led_df.columns else 0, key="ldoc")

        led_conta = st.selectbox("Conta (opcional)", ["(nenhuma)"] + list(led_df.columns),
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

        generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        div["VALOR"] = div["VALOR"].map(normalize_money)
        div = div[div["VALOR"].notna()].copy()
        div = div[div["VALOR"].abs() > 1e-12].copy()

        for c in ["DOCUMENTO", "PREFIXO_TITULO", "HISTORICO_OPERACAO", "CHAVE_DOC"]:
            if c in div.columns:
                div[c] = div[c].replace({np.nan: "", "nan": "", "None": ""}).astype(str).str.strip()

        mask_fin = div["ORIGEM"].eq("Somente Financeiro")
        if "PREFIXO_TITULO" in div.columns:
            div.loc[mask_fin, "DOCUMENTO"] = div.loc[mask_fin, "PREFIXO_TITULO"].where(
                div.loc[mask_fin, "PREFIXO_TITULO"].astype(str).str.len() > 0,
                div.loc[mask_fin, "DOCUMENTO"]
            )

        mask_led = div["ORIGEM"].eq("Somente Contábil")
        if "HISTORICO_OPERACAO" in div.columns:
            missing = mask_led & (div["DOCUMENTO"].astype(str).str.len() == 0)
            div.loc[missing, "DOCUMENTO"] = div.loc[missing, "HISTORICO_OPERACAO"].map(extract_doc_from_ledger_history)

        # Remover colunas não desejadas
        for dropc in ["PREFIXO_TITULO", "CONTA"]:
            if dropc in div.columns:
                div = div.drop(columns=[dropc])

        # Inserir colunas de tratativa
        nuc_sug, mot_sug = [], []
        for _, r in div.iterrows():
            n, m = suggest_nucleo_motivo(r)
            nuc_sug.append(n)
            mot_sug.append(m)

        div["NUCLEO_SUGERIDO"] = nuc_sug
        div["MOTIVO_SUGERIDO"] = mot_sug

        div["NUCLEO_CONFIRMADO"] = div["NUCLEO_SUGERIDO"]
        div["MOTIVO_CONFIRMADO_SN"] = "Não"   # Sim / Não
        div["MOTIVO_CONFIRMADO"] = ""
        div["OBS_USUARIO"] = ""
        div["STATUS"] = "Pendente"
        div["RESOLVIDO"] = False

        div = div.reset_index(drop=True)
        div.index = np.arange(1, len(div) + 1)  # ID 1..N

        st.session_state.results = {"stats": stats, "generated_at": generated_at}
        st.session_state.div_master = div
        st.session_state.page = "resultados"
        st.rerun()

# ----------------------------
# Página: Resultados
# ----------------------------
else:
    if not st.session_state.results or st.session_state.div_master is None:
        st.session_state.page = "upload"
        st.rerun()

    stats = st.session_state.results["stats"]
    generated_at = st.session_state.results["generated_at"]

    NUCLEOS = ["Processo interno", "Cadastro", "Configuração RP", "Não identificado"]
    STATUS_OPTS = ["Pendente", "Em análise", "Resolvido"]

    # ----------------
    # Filtros
    # ----------------
    st.title("Resultados — ConciliaMais (Módulo 1)")
    st.caption(f"Processado em: {generated_at}")

    div_master = st.session_state.div_master.copy()
    div_master["VALOR"] = div_master["VALOR"].map(normalize_money)
    div_master["RESOLVIDO"] = div_master["RESOLVIDO"].fillna(False)
    div_master["STATUS"] = div_master["STATUS"].fillna("Pendente").astype(str)
    div_master["MOTIVO_CONFIRMADO_SN"] = div_master.get("MOTIVO_CONFIRMADO_SN", "Não").fillna("Não").astype(str)

    # Coerência: RESOLVIDO -> STATUS Resolvido
    div_master.loc[div_master["RESOLVIDO"], "STATUS"] = "Resolvido"
    st.session_state.div_master = div_master

    # Filtros (para construir df_view antes do editor)
    fcol1, fcol2, fcol3, fcol4 = st.columns([1.25, 1.0, 2.25, 1.1])
    with fcol1:
        origem = st.selectbox("Filtrar por origem", ["Todas", "Somente Financeiro", "Somente Contábil"])
    with fcol2:
        ver = st.selectbox("Visualizar", ["Todas", "Somente em aberto", "Somente resolvidas"])
    with fcol3:
        busca = st.text_input("Buscar (documento, histórico, chave, motivo)", value="")
    with fcol4:
        st.markdown("<div style='height:1px'></div>", unsafe_allow_html=True)

    df = div_master.copy()

    if origem != "Todas":
        df = df[df["ORIGEM"] == origem].copy()

    res_mask = df["RESOLVIDO"] | (df["STATUS"].astype(str).str.lower().eq("resolvido"))
    if ver == "Somente em aberto":
        df = df[~res_mask].copy()
    elif ver == "Somente resolvidas":
        df = df[res_mask].copy()

    if busca.strip():
        q = busca.strip().lower()
        cols_search = ["DOCUMENTO", "HISTORICO_OPERACAO", "CHAVE_DOC", "NUCLEO_CONFIRMADO", "MOTIVO_CONFIRMADO", "OBS_USUARIO"]
        mask = False
        for c in cols_search:
            if c in df.columns:
                mask = mask | df[c].astype(str).str.lower().str.contains(q, na=False)
        df = df[mask].copy()

    total_filtrado = float(df["VALOR"].sum()) if not df.empty else 0.0

    with fcol4:
        st.markdown(
            f"""
<div class="cm-mini">
  <div class="k">Total do filtro</div>
  <div class="v">{fmt(total_filtrado)}</div>
</div>
""",
            unsafe_allow_html=True,
        )

    # Ordenação
    if "DATA" in df.columns:
        df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df.sort_values(by=["DATA", "VALOR"], ascending=[True, True])

    # Colunas de tela
    view_cols = [
        "ORIGEM", "DATA", "DOCUMENTO", "HISTORICO_OPERACAO", "CHAVE_DOC", "VALOR",
        "NUCLEO_CONFIRMADO", "MOTIVO_CONFIRMADO_SN", "MOTIVO_CONFIRMADO",
        "STATUS", "RESOLVIDO", "OBS_USUARIO"
    ]
    df_view = df[view_cols].copy()

    df_view_display = df_view.copy()
    df_view_display["DATA"] = df_view_display["DATA"].dt.strftime("%d/%m/%Y").fillna("")

    # ----------------
    # Editor
    # ----------------
    st.markdown("#### Tratativa (marque como resolvido e confirme motivo)")

    column_config = {
        "ORIGEM": st.column_config.TextColumn(disabled=True),
        "DATA": st.column_config.TextColumn(disabled=True),
        "DOCUMENTO": st.column_config.TextColumn(disabled=True),
        "HISTORICO_OPERACAO": st.column_config.TextColumn(disabled=True),
        "CHAVE_DOC": st.column_config.TextColumn(disabled=True),
        "VALOR": st.column_config.NumberColumn(format="R$ %.2f", disabled=True),
        "NUCLEO_CONFIRMADO": st.column_config.SelectboxColumn(options=NUCLEOS),
        "MOTIVO_CONFIRMADO_SN": st.column_config.SelectboxColumn(options=["Sim", "Não"]),
        "MOTIVO_CONFIRMADO": st.column_config.TextColumn(),
        "STATUS": st.column_config.SelectboxColumn(options=STATUS_OPTS),
        "RESOLVIDO": st.column_config.CheckboxColumn(),
        "OBS_USUARIO": st.column_config.TextColumn(),
    }

    edited = st.data_editor(
        df_view_display,
        use_container_width=True,
        height=420,
        column_config=column_config,
        key="editor_tratativa",
        hide_index=False,
    )

    # Aplicar mudanças (por ID = index)
    if edited is not None and len(edited) == len(df_view_display):
        to_update = edited.copy()

        # Coerência: RESOLVIDO -> STATUS Resolvido
        res_col = to_update["RESOLVIDO"].fillna(False)
        to_update.loc[res_col, "STATUS"] = "Resolvido"

        # Regra: se RESOLVIDO=True, Núcleo confirmado obrigatório
        bad = res_col & (to_update["NUCLEO_CONFIRMADO"].isna() | (to_update["NUCLEO_CONFIRMADO"].astype(str).str.strip() == ""))
        if bad.any():
            st.error("Para marcar como Resolvido, é obrigatório informar o Núcleo confirmado. Os itens inválidos foram desmarcados.")
            to_update.loc[bad, "RESOLVIDO"] = False
            to_update.loc[bad, "STATUS"] = "Pendente"

        upd_cols = ["NUCLEO_CONFIRMADO", "MOTIVO_CONFIRMADO_SN", "MOTIVO_CONFIRMADO", "STATUS", "RESOLVIDO", "OBS_USUARIO"]
        dm = st.session_state.div_master.copy()
        for c in upd_cols:
            dm.loc[to_update.index, c] = to_update[c].values

        st.session_state.div_master = dm
        div_master = dm.copy()

    # ----------------
    # KPIs (calculados DEPOIS do update, para refletir na hora)
    # ----------------
    div_master["VALOR"] = div_master["VALOR"].map(normalize_money)
    div_master["RESOLVIDO"] = div_master["RESOLVIDO"].fillna(False)
    div_master["STATUS"] = div_master["STATUS"].fillna("Pendente").astype(str)

    resolved_mask = div_master["RESOLVIDO"] | (div_master["STATUS"].str.lower().eq("resolvido"))
    total_itens = len(div_master)
    itens_res = int(resolved_mask.sum())
    itens_ab = int(total_itens - itens_res)
    valor_aberto = float(div_master.loc[~resolved_mask, "VALOR"].sum()) if total_itens else 0.0
    pct_res = (itens_res / total_itens * 100.0) if total_itens else 0.0

    # Cards
    st.markdown(
        f"""
<div class="cm-cards">
  <div class="cm-card">
    <div class="k">Diferenças encontradas</div>
    <div class="v">{total_itens}</div>
    <div class="s">itens de divergência identificados</div>
  </div>
  <div class="cm-card">
    <div class="k">Pendências em aberto</div>
    <div class="v">{fmt(valor_aberto)}</div>
    <div class="s">{itens_ab} itens em aberto</div>
  </div>
  <div class="cm-card">
    <div class="k">Progresso resolvido</div>
    <div class="v">{itens_res} / {total_itens}</div>
    <div class="s">{pct_res:.1f}% resolvido</div>
  </div>
  <div class="cm-card">
    <div class="k">Conferência do cálculo</div>
    <div class="v">{fmt(stats.get("conferencia", np.nan))}</div>
    <div class="s">{pill_calculo(stats.get("conferencia", np.nan))}</div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    st.markdown('<div class="cm-section"></div>', unsafe_allow_html=True)

    # Gráfico: divergências
    st.markdown("### Divergências (Financeiro x Contábil) — visão do motor")
    chart_df = pd.DataFrame(
        {"Valor": [float(stats.get("fin_pend_val", 0.0)), float(stats.get("led_pend_val", 0.0))]},
        index=["Somente Financeiro", "Somente Contábil"],
    )
    st.bar_chart(chart_df)

    # ----------------
    # Ações em massa
    # ----------------
    st.markdown("### Ações em massa")

    scope = st.radio("Aplicar em:", ["Itens filtrados (tela atual)", "IDs informados"], horizontal=True)

    target_ids = []
    if scope == "Itens filtrados (tela atual)":
        target_ids = list(df_view.index)
    else:
        ids_txt = st.text_input("Informe os IDs separados por vírgula (ex: 1,2,15,18)")
        target_ids = [int(x.strip()) for x in ids_txt.split(",") if x.strip().isdigit()]

    cA, cB, cC, cD = st.columns([1.2, 1.2, 1.4, 1.2])

    with cA:
        bulk_motivo_sn = st.selectbox("Motivo confirmado (Sim/Não)", ["(não alterar)", "Sim", "Não"])
    with cB:
        bulk_resolvido = st.selectbox("Marcar como Resolvido", ["(não alterar)", "Sim", "Não"])
    with cC:
        bulk_nucleo = st.selectbox("Núcleo confirmado", ["(não alterar)"] + NUCLEOS)
    with cD:
        bulk_status = st.selectbox("Status", ["(não alterar)"] + STATUS_OPTS)

    if st.button("Aplicar nos selecionados", type="primary", disabled=(len(target_ids) == 0)):
        dm = st.session_state.div_master.copy()

        if bulk_motivo_sn != "(não alterar)":
            dm.loc[target_ids, "MOTIVO_CONFIRMADO_SN"] = bulk_motivo_sn

        if bulk_nucleo != "(não alterar)":
            dm.loc[target_ids, "NUCLEO_CONFIRMADO"] = bulk_nucleo

        if bulk_status != "(não alterar)":
            dm.loc[target_ids, "STATUS"] = bulk_status

        if bulk_resolvido != "(não alterar)":
            if bulk_resolvido == "Sim":
                nuc = dm.loc[target_ids, "NUCLEO_CONFIRMADO"].astype(str).str.strip()
                bad = nuc.eq("") | nuc.isna()
                if bad.any():
                    st.error("Não foi possível marcar como Resolvido: há itens sem Núcleo confirmado. Defina o Núcleo e tente novamente.")
                else:
                    dm.loc[target_ids, "RESOLVIDO"] = True
                    dm.loc[target_ids, "STATUS"] = "Resolvido"
            else:
                dm.loc[target_ids, "RESOLVIDO"] = False
                dm.loc[target_ids, "STATUS"] = dm.loc[target_ids, "STATUS"].replace({"Resolvido": "Pendente"})

        st.session_state.div_master = dm
        st.success(f"Ação aplicada em {len(target_ids)} itens.")
        st.rerun()

    # ----------------
    # Detalhe do item
    # ----------------
    st.markdown("### Detalhe do item")
    pick_id = st.number_input("Digite o ID do item para ver detalhes", min_value=1, max_value=max(1, int(div_master.index.max())), value=1, step=1)
    if pick_id in div_master.index:
        r = div_master.loc[pick_id]
        dt_txt = ""
        try:
            if pd.notna(r.get("DATA")):
                dt_txt = pd.to_datetime(r.get("DATA")).strftime("%d/%m/%Y")
        except Exception:
            dt_txt = str(r.get("DATA") or "")

        resumo = (
            f"ID: {pick_id}\n"
            f"ORIGEM: {r.get('ORIGEM','')}\n"
            f"DATA: {dt_txt}\n"
            f"DOCUMENTO: {r.get('DOCUMENTO','')}\n"
            f"CHAVE: {r.get('CHAVE_DOC','')}\n"
            f"VALOR: {fmt(r.get('VALOR', np.nan))}\n"
            f"NUCLEO_CONFIRMADO: {r.get('NUCLEO_CONFIRMADO','')}\n"
            f"MOTIVO_CONFIRMADO_SN: {r.get('MOTIVO_CONFIRMADO_SN','')}\n"
            f"MOTIVO_CONFIRMADO: {r.get('MOTIVO_CONFIRMADO','')}\n"
            f"STATUS: {r.get('STATUS','')}\n"
            f"RESOLVIDO: {bool(r.get('RESOLVIDO', False))}\n"
            f"OBS: {r.get('OBS_USUARIO','')}\n"
            f"HISTÓRICO: {r.get('HISTORICO_OPERACAO','')}"
        )

        st.markdown(
            f"""
<div class="cm-detail">
  <div class="title">Item #{pick_id}</div>
  <div class="row"><span class="label">Origem:</span> <span class="val">{r.get('ORIGEM','')}</span></div>
  <div class="row"><span class="label">Data:</span> <span class="val">{dt_txt}</span></div>
  <div class="row"><span class="label">Documento:</span> <span class="val">{r.get('DOCUMENTO','')}</span></div>
  <div class="row"><span class="label">Valor:</span> <span class="val">{fmt(r.get('VALOR', np.nan))}</span></div>
  <div class="row"><span class="label">Núcleo:</span> <span class="val">{r.get('NUCLEO_CONFIRMADO','')}</span></div>
  <div class="row"><span class="label">Status:</span> <span class="val">{r.get('STATUS','')}</span></div>
</div>
""",
            unsafe_allow_html=True,
        )
        st.text_area("Copiar resumo (e-mail/ticket)", value=resumo, height=160)

    # ----------------
    # Export
    # ----------------
    st.markdown("### Exportar")
    filtros = {"origem": origem, "ver": ver, "busca": busca.strip()}

    excel_bytes = to_excel_divergencias_filtradas(
        df_filtrado=df_view,
        total_filtrado=float(df_view["VALOR"].sum()) if len(df_view) else 0.0,
        total_aberto=valor_aberto,
        filtros=filtros,
        stats=stats,
        generated_at=generated_at
    )

    st.download_button(
        "Baixar Divergências (Excel) — exatamente como filtrado",
        data=excel_bytes,
        file_name=f"ConciliaMais_DivergenciasFiltradas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    pdf_bytes = to_pdf_resumo(stats, generated_at, st.session_state.div_master)
    st.download_button(
        "Baixar Relatório Resumo (PDF) — executivo",
        data=pdf_bytes,
        file_name=f"ConciliaMais_Resumo_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf",
    )

    st.markdown("---")
    cback, creset = st.columns([1, 2])
    with cback:
        if st.button("Voltar para Upload"):
            st.session_state.page = "upload"
            st.rerun()
    with creset:
        if st.button("Limpar resultado e recomeçar"):
            st.session_state.results = None
            st.session_state.div_master = None
            st.session_state.page = "upload"
            st.rerun()
