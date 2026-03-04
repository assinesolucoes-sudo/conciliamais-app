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
# CSS (Light Theme + UX)
# ----------------------------
st.markdown(
    """
<style>
/* Base */
:root{
  --bg: #F7F9FC;
  --card: #FFFFFF;
  --border: #E6EAF2;
  --text: #0F172A;
  --muted: #64748B;
  --primary: #2563EB;
  --ok: #16A34A;
  --warn: #F59E0B;
  --bad: #DC2626;
  --shadow: 0 10px 24px rgba(15, 23, 42, 0.06);
}

html, body, [class*="css"]  { color: var(--text) !important; }
body { background: var(--bg) !important; }
.block-container { padding-top: 1.0rem; padding-bottom: 3.5rem; max-width: 1400px; }

h1, h2, h3 { letter-spacing: -0.02em; }
small, .stCaption, .stMarkdown p { color: var(--muted); }

/* Cards / blocks */
.cm-shell {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 18px;
  padding: 14px;
  box-shadow: var(--shadow);
}

.cm-help { color: var(--muted); font-size: 13px; margin-top: -6px; }
.cm-section { margin-top: 16px; }

.cm-cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-top: 10px; }
.cm-card {
  border-radius: 16px;
  padding: 14px 14px 12px 14px;
  background: var(--card);
  border: 1px solid var(--border);
  box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
}
.cm-card .k { font-size: 12px; color: var(--muted); margin-bottom: 6px; }
.cm-card .v { font-size: 22px; font-weight: 900; color: var(--text); }
.cm-card .s { font-size: 12px; color: var(--muted); margin-top: 6px; }

.cm-mini {
  border-radius: 14px;
  padding: 10px 12px;
  background: var(--card);
  border: 1px solid var(--border);
  text-align: right;
  box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
}
.cm-mini .k { font-size: 12px; color: var(--muted); margin-bottom: 4px; }
.cm-mini .v { font-size: 20px; font-weight: 900; letter-spacing: -0.01em; color: var(--text); }

/* Pills */
.cm-pill { display: inline-block; padding: 4px 10px; border-radius: 999px; font-size: 12px; font-weight: 800; border: 1px solid transparent; }
.cm-ok { background: rgba(22,163,74,.10); color: var(--ok); border-color: rgba(22,163,74,.25); }
.cm-warn { background: rgba(245,158,11,.12); color: #8A5A00; border-color: rgba(245,158,11,.25); }
.cm-bad { background: rgba(220,38,38,.10); color: var(--bad); border-color: rgba(220,38,38,.22); }

.cm-tag { display:inline-flex; align-items:center; gap:6px; padding:4px 10px; border-radius:999px; font-size:12px; font-weight:800; border: 1px solid var(--border); background: #F8FAFC; color: var(--text); }
.cm-dot { width:8px; height:8px; border-radius:99px; display:inline-block; }
.cm-dot-fin { background: #2563EB; }
.cm-dot-led { background: #7C3AED; }
.cm-dot-ok { background: var(--ok); }
.cm-dot-warn { background: var(--warn); }
.cm-dot-bad { background: var(--bad); }

.cm-detail {
  border-radius: 16px;
  padding: 14px;
  background: var(--card);
  border: 1px solid var(--border);
  box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
}
.cm-detail .title { font-size: 14px; font-weight: 900; margin-bottom: 10px; }
.cm-detail .row { font-size: 13px; margin: 4px 0; }
.cm-detail .label { color: var(--muted); }
.cm-detail .val { font-weight: 700; color: var(--text); }

/* Sticky footer actions */
.cm-footer {
  position: fixed;
  left: 0;
  right: 0;
  bottom: 0;
  background: rgba(247, 249, 252, 0.92);
  backdrop-filter: blur(8px);
  border-top: 1px solid var(--border);
  padding: 10px 16px;
  z-index: 999;
}
.cm-footer-inner{
  max-width: 1400px;
  margin: 0 auto;
  display:flex;
  gap:10px;
  align-items:center;
  justify-content: space-between;
}
.cm-footer-left{ display:flex; gap:10px; align-items:center; }
.cm-footer-right{ display:flex; gap:10px; align-items:center; }
.cm-footer-note{ font-size:12px; color: var(--muted); }
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
# UX helpers
# ----------------------------
def origem_tag(origem: str) -> str:
    if str(origem) == "Somente Financeiro":
        return '<span class="cm-tag"><span class="cm-dot cm-dot-fin"></span>Somente Financeiro</span>'
    if str(origem) == "Somente Contábil":
        return '<span class="cm-tag"><span class="cm-dot cm-dot-led"></span>Somente Contábil</span>'
    return '<span class="cm-tag"><span class="cm-dot"></span>Todas</span>'

def severidade(valor) -> str:
    try:
        v = abs(float(valor))
    except Exception:
        return "Normal"
    if v <= 100:
        return "Normal"
    if v <= 1000:
        return "Atenção"
    return "Crítica"

def severidade_tag(label: str) -> str:
    lab = str(label)
    if lab == "Normal":
        return '<span class="cm-tag"><span class="cm-dot cm-dot-ok"></span>Normal</span>'
    if lab == "Atenção":
        return '<span class="cm-tag"><span class="cm-dot cm-dot-warn"></span>Atenção</span>'
    return '<span class="cm-tag"><span class="cm-dot cm-dot-bad"></span>Crítica</span>'

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
        fmt_k = wb.add_format({"bold": True, "font_size": 10, "font_color": "#475569"})
        fmt_txt = wb.add_format({"border": 1})
        fmt_hdr = wb.add_format({"bold": True, "border": 1, "align": "center", "valign": "vcenter", "bg_color": "#EEF2FF"})
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
# PDF Resumo (Executivo + Top 10)
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

    resolved = df.get("RESOLVIDO", False)
    if isinstance(resolved, (pd.Series, pd.Index)):
        resolved = resolved.fillna(False)
    else:
        resolved = pd.Series([False] * len(df), index=df.index)
    status = df.get("STATUS", "Pendente").astype(str).fillna("Pendente")
    resolved = resolved | (status.str.lower().eq("resolvido"))
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
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1E293B")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
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
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1E293B")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
    ]))
    story.append(t2)
    story.append(Spacer(1, 14))

    # Top 10 pendências por impacto (abertas)
    story.append(Paragraph("Top 10 pendências mais impactantes (em aberto)", styles["Heading2"]))
    story.append(Spacer(1, 6))

    top = df.loc[~df["__RES"]].copy()
    top["ABS"] = top["VALOR"].abs()
    top = top.sort_values(["ABS"], ascending=False).head(10)

    top_rows = [["#", "Origem", "Data", "Documento", "Valor", "Núcleo"]]
    for i, (_, r) in enumerate(top.iterrows(), start=1):
        dt = ""
        try:
            dt = pd.to_datetime(r.get("DATA")).strftime("%d/%m/%Y") if pd.notna(r.get("DATA")) else ""
        except Exception:
            dt = str(r.get("DATA") or "")
        top_rows.append([
            str(i),
            str(r.get("ORIGEM","")),
            dt,
            str(r.get("DOCUMENTO","")),
            fmt(float(r.get("VALOR", 0.0))),
            str(r.get("NUCLEO_CONFIRMADO","Não identificado") or "Não identificado"),
        ])

    t_top = Table(top_rows, colWidths=[22, 85, 58, 175, 75, 80])
    t_top.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1E293B")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("ALIGN", (0,0), (0,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ]))
    story.append(t_top)
    story.append(Spacer(1, 14))

    story.append(Paragraph("Distribuição por núcleo (confirmado)", styles["Heading2"]))
    story.append(Spacer(1, 6))

    nuc = df.copy()
    nuc["NUCLEO_CONFIRMADO"] = nuc.get("NUCLEO_CONFIRMADO", "Não identificado").fillna("Não identificado").replace("", "Não identificado")
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
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1E293B")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
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
if "upload_step" not in st.session_state:
    st.session_state.upload_step = 1
if "toast_once" not in st.session_state:
    st.session_state.toast_once = False

NUCLEOS = ["Processo interno", "Cadastro", "Configuração RP", "Não identificado"]
STATUS_OPTS = ["Pendente", "Em análise", "Resolvido"]

# ----------------------------
# Página: Upload (Wizard)
# ----------------------------
if st.session_state.page == "upload":
    st.title("ConciliaMais — Conferência de Extrato Bancário")
    st.caption("Extrato Financeiro + Razão Contábil -> Match automático -> Divergências -> Tratativa")

    with st.container():
        st.markdown('<div class="cm-shell">', unsafe_allow_html=True)
        st.markdown("### Etapas")
        st.progress(min(max((st.session_state.upload_step - 1) / 3, 0.0), 1.0))
        st.markdown(
            f"<div class='cm-help'>1) Upload  |  2) Mapeamento  |  3) Validação de saldos  |  4) Processar</div>",
            unsafe_allow_html=True
        )
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="cm-section"></div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1) Extrato Financeiro")
        st.markdown('<div class="cm-help">Faça o upload da planilha do Extrato Financeiro.</div>', unsafe_allow_html=True)
        fin_file = st.file_uploader("Upload do Extrato Financeiro (.xlsx ou .csv)", type=["xlsx","csv"], key="fin")
    with c2:
        st.subheader("1) Razão Contábil")
        st.markdown('<div class="cm-help">Faça o upload da planilha do Razão Contábil.</div>', unsafe_allow_html=True)
        led_file = st.file_uploader("Upload do Razão Contábil (.xlsx ou .csv)", type=["xlsx","csv"], key="led")

    if not fin_file or not led_file:
        st.session_state.upload_step = 1
        st.info("Faça o upload dos dois arquivos para liberar o restante.")
        st.stop()

    st.session_state.upload_step = 2

    fin_df = read_table(fin_file)
    led_df = read_table(led_file)

    fin_guess = auto_detect_financial(fin_df)
    led_guess = auto_detect_ledger(led_df)

    st.markdown("### 2) Mapeamento de colunas (auto-detectado — ajuste se precisar)")
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

    # Preview rápido para evitar erro de mapeamento
    st.markdown("### Pré-visualização rápida (5 linhas)")
    p1, p2 = st.columns(2)
    with p1:
        cols_prev_fin = [cfg["fin_date"]]
        for cc in [cfg.get("fin_operacao"), cfg.get("fin_documento"), cfg.get("fin_prefixo"), cfg.get("fin_amount"), cfg.get("fin_entradas"), cfg.get("fin_saidas"), cfg.get("fin_saldo")]:
            if cc and cc not in cols_prev_fin:
                cols_prev_fin.append(cc)
        st.dataframe(fin_df[cols_prev_fin].head(5), use_container_width=True, height=210)
    with p2:
        cols_prev_led = [cfg["led_date"]]
        for cc in [cfg.get("led_historico"), cfg.get("led_doc"), cfg.get("led_conta"), cfg.get("led_amount"), cfg.get("led_debito"), cfg.get("led_credito"), cfg.get("led_saldo")]:
            if cc and cc not in cols_prev_led:
                cols_prev_led.append(cc)
        st.dataframe(led_df[cols_prev_led].head(5), use_container_width=True, height=210)

    st.session_state.upload_step = 3

    st.markdown("### 3) Alertas iniciais (Saldo anterior)")
    f_norm, l_norm = build_normalized(fin_df, led_df, cfg)
    saldo_ant_fin = compute_saldo_anterior(f_norm)
    saldo_ant_led = compute_saldo_anterior(l_norm)
    diff_ant = np.nan if (pd.isna(saldo_ant_fin) or pd.isna(saldo_ant_led)) else round(saldo_ant_fin - saldo_ant_led, 2)

    proceed_ok = True
    if pd.isna(diff_ant):
        st.info("Não foi possível calcular saldo anterior automaticamente. Se existir saldo nos dois arquivos, selecione a coluna de saldo.")
    else:
        if abs(diff_ant) > 0.01:
            st.warning(f"Saldo anterior não bate (Fin - Cont = {fmt(diff_ant)}). Diferença pode estar em períodos anteriores.")
            proceed_ok = st.checkbox("Prosseguir mesmo assim", value=False)
        else:
            st.success("Saldo anterior bate (OK).")

    date_tol = st.number_input("Tolerância de dias para match por data (0 = mesma data)", min_value=0, max_value=10, value=0, step=1)

    st.session_state.upload_step = 4

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

        # Severidade (UX)
        div["SEVERIDADE"] = div["VALOR"].map(severidade)

        div = div.reset_index(drop=True)
        div.index = np.arange(1, len(div) + 1)  # ID 1..N

        st.session_state.results = {"stats": stats, "generated_at": generated_at}
        st.session_state.div_master = div
        st.session_state.page = "resultados"
        st.session_state.toast_once = False
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

    st.title("Resultados — ConciliaMais (Módulo 1)")
    st.caption(f"Processado em: {generated_at}")

    div_master = st.session_state.div_master.copy()
    div_master["VALOR"] = div_master["VALOR"].map(normalize_money)
    div_master["RESOLVIDO"] = div_master["RESOLVIDO"].fillna(False)
    div_master["STATUS"] = div_master["STATUS"].fillna("Pendente").astype(str)
    div_master["MOTIVO_CONFIRMADO_SN"] = div_master.get("MOTIVO_CONFIRMADO_SN", "Não").fillna("Não").astype(str)

    # Coerência: RESOLVIDO -> STATUS Resolvido
    div_master.loc[div_master["RESOLVIDO"], "STATUS"] = "Resolvido"

    # Garantir severidade
    if "SEVERIDADE" not in div_master.columns:
        div_master["SEVERIDADE"] = div_master["VALOR"].map(severidade)

    st.session_state.div_master = div_master

    # KPIs globais
    resolved_mask = div_master["RESOLVIDO"] | (div_master["STATUS"].str.lower().eq("resolvido"))
    total_itens = len(div_master)
    itens_res = int(resolved_mask.sum())
    itens_ab = int(total_itens - itens_res)
    valor_aberto = float(div_master.loc[~resolved_mask, "VALOR"].sum()) if total_itens else 0.0
    pct_res = (itens_res / total_itens * 100.0) if total_itens else 0.0

    # Header cards
    st.markdown(
        f"""
<div class="cm-cards">
  <div class="cm-card">
    <div class="k">Diferenças encontradas</div>
    <div class="v">{total_itens}</div>
    <div class="s">itens de divergência identificados</div>
  </div>
  <div class="cm-card">
    <div class="k">Pendências em aberto (valor)</div>
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

    # Resumo analítico (priorização)
    with st.expander("Resumo para priorização (abertos, top impacto, pendências de preenchimento)", expanded=True):
        df_open = div_master.loc[~resolved_mask].copy()
        df_open["ABS"] = df_open["VALOR"].abs()
        top_open = df_open.sort_values("ABS", ascending=False).head(10)

        miss_nuc = df_open["NUCLEO_CONFIRMADO"].isna() | (df_open["NUCLEO_CONFIRMADO"].astype(str).str.strip() == "")
        miss_mot = (df_open["MOTIVO_CONFIRMADO_SN"].astype(str).str.strip() != "Sim")

        csa, csb, csc, csd = st.columns(4)
        with csa:
            st.markdown("**Top 10 em aberto por impacto**")
            st.dataframe(
                top_open[["ORIGEM", "DATA", "DOCUMENTO", "VALOR", "SEVERIDADE", "NUCLEO_CONFIRMADO"]].copy(),
                use_container_width=True,
                height=260
            )
        with csb:
            st.markdown("**Pendências de preenchimento**")
            st.write(f"Itens em aberto sem Núcleo confirmado: {int(miss_nuc.sum())}")
            st.write(f"Itens em aberto sem Motivo confirmado (Sim): {int(miss_mot.sum())}")
            st.write("")
            st.markdown("**Distribuição por Núcleo (abertos)**")
            if len(df_open):
                dist = df_open.copy()
                dist["NUCLEO_CONFIRMADO"] = dist["NUCLEO_CONFIRMADO"].fillna("Não identificado").replace("", "Não identificado")
                dist = dist.groupby("NUCLEO_CONFIRMADO", dropna=False).agg(Itens=("VALOR","size"), Valor=("VALOR","sum")).reset_index()
                dist = dist.sort_values("Valor", ascending=False)
                st.dataframe(dist, use_container_width=True, height=260)
            else:
                st.info("Sem pendências em aberto.")
        with csc:
            st.markdown("**Sinais rápidos**")
            st.markdown(origem_tag("Somente Financeiro"), unsafe_allow_html=True)
            st.markdown(origem_tag("Somente Contábil"), unsafe_allow_html=True)
            st.markdown(severidade_tag("Normal"), unsafe_allow_html=True)
            st.markdown(severidade_tag("Atenção"), unsafe_allow_html=True)
            st.markdown(severidade_tag("Crítica"), unsafe_allow_html=True)
        with csd:
            st.markdown("**Visão do motor (valores)**")
            chart_df = pd.DataFrame(
                {"Valor": [float(stats.get("fin_pend_val", 0.0)), float(stats.get("led_pend_val", 0.0))]},
                index=["Somente Financeiro", "Somente Contábil"],
            )
            st.bar_chart(chart_df)

    # ----------------
    # Filtros
    # ----------------
    st.markdown("### Filtros")
    fcol1, fcol2, fcol3, fcol4, fcol5 = st.columns([1.15, 1.0, 1.0, 2.15, 1.1])
    with fcol1:
        origem = st.selectbox("Origem", ["Todas", "Somente Financeiro", "Somente Contábil"])
    with fcol2:
        ver = st.selectbox("Visualizar", ["Todas", "Somente em aberto", "Somente resolvidas"])
    with fcol3:
        sev = st.selectbox("Severidade", ["Todas", "Normal", "Atenção", "Crítica"])
    with fcol4:
        busca = st.text_input("Buscar (documento, histórico, chave, núcleo, motivo)", value="")
    with fcol5:
        st.markdown("<div style='height:1px'></div>", unsafe_allow_html=True)

    # Filtros rápidos
    q1, q2, q3, q4 = st.columns([1.2, 1.2, 1.2, 1.2])
    with q1:
        if st.button("Top 20 por impacto (abertos)"):
            origem = "Todas"
            ver = "Somente em aberto"
            sev = "Todas"
            busca = ""
            st.session_state["_quick_top"] = True
        else:
            st.session_state["_quick_top"] = st.session_state.get("_quick_top", False)
    with q2:
        if st.button("Sem Núcleo (abertos)"):
            st.session_state["_quick_sem_nuc"] = True
        else:
            st.session_state["_quick_sem_nuc"] = st.session_state.get("_quick_sem_nuc", False)
    with q3:
        if st.button("Sem Motivo confirmado (abertos)"):
            st.session_state["_quick_sem_mot"] = True
        else:
            st.session_state["_quick_sem_mot"] = st.session_state.get("_quick_sem_mot", False)
    with q4:
        if st.button("Limpar filtros rápidos"):
            st.session_state["_quick_top"] = False
            st.session_state["_quick_sem_nuc"] = False
            st.session_state["_quick_sem_mot"] = False

    df = div_master.copy()

    if origem != "Todas":
        df = df[df["ORIGEM"] == origem].copy()

    res_mask_df = df["RESOLVIDO"] | (df["STATUS"].astype(str).str.lower().eq("resolvido"))
    if ver == "Somente em aberto":
        df = df[~res_mask_df].copy()
    elif ver == "Somente resolvidas":
        df = df[res_mask_df].copy()

    if sev != "Todas":
        df = df[df["SEVERIDADE"] == sev].copy()

    if busca.strip():
        q = busca.strip().lower()
        cols_search = ["DOCUMENTO", "HISTORICO_OPERACAO", "CHAVE_DOC", "NUCLEO_CONFIRMADO", "MOTIVO_CONFIRMADO", "OBS_USUARIO", "SEVERIDADE"]
        mask = False
        for c in cols_search:
            if c in df.columns:
                mask = mask | df[c].astype(str).str.lower().str.contains(q, na=False)
        df = df[mask].copy()

    # Aplicar filtros rápidos sem quebrar os filtros manuais
    if st.session_state.get("_quick_top", False):
        df2 = df.copy()
        df2 = df2[~(df2["RESOLVIDO"] | (df2["STATUS"].astype(str).str.lower().eq("resolvido")))].copy()
        df2["ABS"] = df2["VALOR"].abs()
        df = df2.sort_values("ABS", ascending=False).head(20)

    if st.session_state.get("_quick_sem_nuc", False):
        df2 = df.copy()
        df2 = df2[~(df2["RESOLVIDO"] | (df2["STATUS"].astype(str).str.lower().eq("resolvido")))].copy()
        df = df2[df2["NUCLEO_CONFIRMADO"].isna() | (df2["NUCLEO_CONFIRMADO"].astype(str).str.strip() == "")].copy()

    if st.session_state.get("_quick_sem_mot", False):
        df2 = df.copy()
        df2 = df2[~(df2["RESOLVIDO"] | (df2["STATUS"].astype(str).str.lower().eq("resolvido")))].copy()
        df = df2[df2["MOTIVO_CONFIRMADO_SN"].astype(str).str.strip() != "Sim"].copy()

    total_filtrado = float(df["VALOR"].sum()) if not df.empty else 0.0
    with fcol5:
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
        "ORIGEM", "SEVERIDADE", "DATA", "DOCUMENTO", "HISTORICO_OPERACAO", "CHAVE_DOC", "VALOR",
        "NUCLEO_CONFIRMADO", "MOTIVO_CONFIRMADO_SN", "MOTIVO_CONFIRMADO",
        "STATUS", "RESOLVIDO", "OBS_USUARIO"
    ]
    df_view = df[view_cols].copy()

    df_view_display = df_view.copy()
    df_view_display["DATA"] = pd.to_datetime(df_view_display["DATA"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")

    # ----------------
    # Editor
    # ----------------
    st.markdown("### Tratativa (edite e marque como resolvido)")
    st.markdown('<div class="cm-help">Regras: para marcar como Resolvido, é obrigatório Núcleo confirmado e Motivo confirmado = Sim.</div>', unsafe_allow_html=True)

    column_config = {
        "ORIGEM": st.column_config.TextColumn(disabled=True),
        "SEVERIDADE": st.column_config.SelectboxColumn(options=["Normal", "Atenção", "Crítica"], disabled=True),
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

        # Regras de validação para marcar como resolvido
        nuc_ok = ~(to_update["NUCLEO_CONFIRMADO"].isna() | (to_update["NUCLEO_CONFIRMADO"].astype(str).str.strip() == ""))
        mot_ok = (to_update["MOTIVO_CONFIRMADO_SN"].astype(str).str.strip() == "Sim")

        bad = res_col & (~nuc_ok | ~mot_ok)
        if bad.any():
            st.error("Para marcar como Resolvido, é obrigatório informar Núcleo confirmado e Motivo confirmado = Sim. Itens inválidos foram desmarcados.")
            to_update.loc[bad, "RESOLVIDO"] = False
            to_update.loc[bad, "STATUS"] = "Pendente"

        upd_cols = ["NUCLEO_CONFIRMADO", "MOTIVO_CONFIRMADO_SN", "MOTIVO_CONFIRMADO", "STATUS", "RESOLVIDO", "OBS_USUARIO"]
        dm = st.session_state.div_master.copy()
        for c in upd_cols:
            dm.loc[to_update.index, c] = to_update[c].values

        # Atualizar severidade (não deveria mudar, mas mantém consistente)
        dm["SEVERIDADE"] = dm["VALOR"].map(severidade)

        st.session_state.div_master = dm
        div_master = dm.copy()

        if not st.session_state.toast_once:
            try:
                st.toast("Alterações aplicadas.", icon=None)
            except Exception:
                st.success("Alterações aplicadas.")
            st.session_state.toast_once = True

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
            f"SEVERIDADE: {r.get('SEVERIDADE','')}\n"
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
  <div class="row"><span class="label">Severidade:</span> <span class="val">{r.get('SEVERIDADE','')}</span></div>
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
    # Export (prepara bytes)
    # ----------------
    st.markdown("### Exportar")
    filtros = {"origem": origem, "ver": ver, "severidade": sev, "busca": busca.strip()}

    excel_bytes = to_excel_divergencias_filtradas(
        df_filtrado=df_view,
        total_filtrado=float(df_view["VALOR"].sum()) if len(df_view) else 0.0,
        total_aberto=valor_aberto,
        filtros=filtros,
        stats=stats,
        generated_at=generated_at
    )
    pdf_bytes = to_pdf_resumo(stats, generated_at, st.session_state.div_master)

    # Sticky footer actions (downloads + navegação)
    st.markdown(
        f"""
<div class="cm-footer">
  <div class="cm-footer-inner">
    <div class="cm-footer-left">
      <div class="cm-footer-note">Ações rápidas</div>
    </div>
    <div class="cm-footer-right">
      <!-- Espaço reservado: os botões do Streamlit vão logo abaixo -->
    </div>
  </div>
</div>
""",
        unsafe_allow_html=True
    )

    # Colocar os botões logo abaixo (Streamlit controla o HTML do botão, mas a área fica próxima do footer)
    colA, colB, colC, colD = st.columns([1.9, 1.9, 1.2, 1.2])
    with colA:
        st.download_button(
            "Baixar Divergências (Excel) — como filtrado",
            data=excel_bytes,
            file_name=f"ConciliaMais_DivergenciasFiltradas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with colB:
        st.download_button(
            "Baixar Relatório Resumo (PDF) — executivo",
            data=pdf_bytes,
            file_name=f"ConciliaMais_Resumo_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
    with colC:
        if st.button("Voltar para Upload", use_container_width=True):
            st.session_state.page = "upload"
            st.session_state.upload_step = 1
            st.rerun()
    with colD:
        if st.button("Limpar e recomeçar", use_container_width=True):
            st.session_state.results = None
            st.session_state.div_master = None
            st.session_state.page = "upload"
            st.session_state.upload_step = 1
            st.rerun()
