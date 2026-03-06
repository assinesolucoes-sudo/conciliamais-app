# =========================================================
# ConciliaMais — V6 (UX first, sem perder a lógica)
# - Legenda interativa (aplica filtros)
# - Severidade explicada + thresholds configuráveis
# - Filtros próximos da tabela
# - Busca inclui VALOR
# - Ações em massa após tabela (fluxo natural)
# - Export Excel/PDF com Sim/Não (sem True/False)
# =========================================================

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

# =========================================================
# Config
# =========================================================
st.set_page_config(page_title="ConciliaMais — Conferência de Extrato Bancário", layout="wide")

# =========================================================
# CSS (Dark + Marca Azul, sem vermelho como CTA)
# =========================================================
st.markdown(
    """
<style>
:root{
  --bg: #0B1220;
  --card: #0F172A;
  --card2:#111C33;
  --border: rgba(148,163,184,.18);
  --text: #E5E7EB;
  --muted: rgba(226,232,240,.72);
  --primary: #2563EB; /* azul marca */
  --primary2:#1D4ED8;
  --ok: #16A34A;
  --warn: #F59E0B;
  --bad: #EF4444;

  --shadow: 0 10px 24px rgba(0,0,0,.35);
}

html, body, [class*="css"]  { color: var(--text) !important; }
body { background: var(--bg) !important; }

.block-container {
  padding-top: 1.0rem;
  padding-bottom: 2.2rem;
  max-width: 1450px;
}

/* cards */
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
  box-shadow: 0 6px 18px rgba(0,0,0,.28);
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
  box-shadow: 0 6px 18px rgba(0,0,0,.25);
}
.cm-mini .k { font-size: 12px; color: var(--muted); margin-bottom: 4px; }
.cm-mini .v { font-size: 20px; font-weight: 900; letter-spacing: -0.01em; color: var(--text); }

.cm-pill { display: inline-block; padding: 4px 10px; border-radius: 999px; font-size: 12px; font-weight: 800; border: 1px solid transparent; }
.cm-ok { background: rgba(22,163,74,.14); color: #86EFAC; border-color: rgba(22,163,74,.35); }
.cm-warn { background: rgba(245,158,11,.14); color: #FCD34D; border-color: rgba(245,158,11,.35); }
.cm-bad { background: rgba(239,68,68,.14); color: #FCA5A5; border-color: rgba(239,68,68,.35); }

.cm-tag { display:inline-flex; align-items:center; gap:6px; padding:4px 10px; border-radius:999px; font-size:12px; font-weight:800; border: 1px solid var(--border); background: rgba(255,255,255,.04); color: var(--text); }
.cm-dot { width:8px; height:8px; border-radius:99px; display:inline-block; }
.cm-dot-fin { background: #60A5FA; }
.cm-dot-led { background: #A78BFA; }
.cm-dot-ok { background: #22C55E; }
.cm-dot-warn { background: #F59E0B; }
.cm-dot-bad { background: #EF4444; }

.cm-detail {
  border-radius: 16px;
  padding: 14px;
  background: var(--card);
  border: 1px solid var(--border);
  box-shadow: 0 6px 18px rgba(0,0,0,.25);
}
.cm-detail .title { font-size: 14px; font-weight: 900; margin-bottom: 10px; }
.cm-detail .row { font-size: 13px; margin: 4px 0; }
.cm-detail .label { color: var(--muted); }
.cm-detail .val { font-weight: 700; color: var(--text); }

.cm-banner {
  border-radius: 16px;
  padding: 12px 14px;
  background: rgba(245,158,11,.10);
  border: 1px solid rgba(245,158,11,.22);
  box-shadow: 0 6px 18px rgba(0,0,0,.22);
}
.cm-banner strong{ color:#FCD34D; }
.cm-banner .muted{ color: var(--muted); font-size: 13px; margin-top: 4px; }

.cm-breadcrumb{
  color: rgba(226,232,240,.78);
  font-size: 13px;
  margin-top: -8px;
}

.cm-actionbar{
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 16px;
  padding: 12px;
  box-shadow: 0 6px 18px rgba(0,0,0,.22);
}

.cm-badge{
  display:inline-flex;
  align-items:center;
  gap:6px;
  padding:6px 10px;
  border-radius:999px;
  font-size:12px;
  font-weight:900;
  border:1px solid var(--border);
  background: rgba(37,99,235,.14);
  color:#BFDBFE;
}

.cm-divider{
  height: 1px;
  background: rgba(148,163,184,.18);
  margin: 12px 0;
}

div.stButton > button[kind="primary"]{
  background: var(--primary) !important;
  border: 1px solid rgba(147,197,253,.35) !important;
  color: white !important;
  border-radius: 12px !important;
  font-weight: 900 !important;
  height: 42px !important;
}
div.stButton > button[kind="primary"]:hover{
  background: var(--primary2) !important;
}

div.stButton > button{
  border-radius: 12px !important;
  height: 42px !important;
}
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# Helpers (motor)
# =========================================================
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

# Núcleo
NUCLEOS = ["Processo interno", "Cadastro", "Configuração RP", "Não identificado"]
STATUS_OPTS = ["Pendente", "Em análise", "Resolvido"]

def suggest_nucleo(row):
    origem = str(row.get("ORIGEM", "")).lower()
    hist = str(row.get("HISTORICO_OPERACAO", "")).lower()
    doc = str(row.get("DOCUMENTO","")).strip()

    if any(k in hist for k in ["cancelamento de baixa", "canc baixa", "estorno de baixa", "estorno baixa", "canc. baixa"]):
        return "Processo interno"
    if any(k in hist for k in ["baixa", "liquidação", "liquidacao", "pagamento", "pagto", "estorno"]):
        return "Processo interno"

    if "somente financeiro" in origem and (doc != "" or "mov" in hist or "titulo" in hist or "título" in hist):
        return "Cadastro"

    if any(k in hist for k in ["rp", "reprocess", "rotina", "processamento", "integracao", "integração"]):
        return "Configuração RP"

    return "Não identificado"

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

# =========================================================
# UX helpers (V6)
# =========================================================
def _money_search_tokens(v) -> str:
    """Gera tokens de busca para valor: aceita 1600, 1.600,00, -1600.00 etc."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    try:
        x = float(v)
    except Exception:
        return str(v)
    br = fmt(x)                      # "1.600,00"
    raw1 = str(x)                    # "1600.0"
    raw2 = f"{x:.2f}"                # "1600.00"
    raw3 = br.replace(".", "")       # "1600,00"
    raw4 = raw3.replace(",", ".")    # "1600.00"
    return f"{br} {raw1} {raw2} {raw3} {raw4}"

def origem_tag(origem: str) -> str:
    if str(origem) == "Somente Financeiro":
        return '<span class="cm-tag"><span class="cm-dot cm-dot-fin"></span>Somente Financeiro</span>'
    if str(origem) == "Somente Contábil":
        return '<span class="cm-tag"><span class="cm-dot cm-dot-led"></span>Somente Contábil</span>'
    return '<span class="cm-tag"><span class="cm-dot"></span>Todas</span>'

def severidade_tag(label: str) -> str:
    lab = str(label)
    if lab == "Normal":
        return '<span class="cm-tag"><span class="cm-dot cm-dot-ok"></span>Normal</span>'
    if lab == "Atenção":
        return '<span class="cm-tag"><span class="cm-dot cm-dot-warn"></span>Atenção</span>'
    return '<span class="cm-tag"><span class="cm-dot cm-dot-bad"></span>Crítica</span>'

def bool_to_sim_nao(x):
    return "Sim" if bool(x) else "Não"

# =========================================================
# Excel Export (Tabela + Header azul + TotalRow SUBTOTAL)
# =========================================================
def _autofit_worksheet(ws, df, start_col, max_width=70, min_width=10):
    for j, col in enumerate(df.columns):
        ser = df[col].astype(str).fillna("")
        sample = ser.head(250).tolist()
        max_len = max([len(str(col))] + [len(s) for s in sample]) if sample else len(str(col))
        width = max(min_width, min(max_width, max_len + 2))
        ws.set_column(start_col + j, start_col + j, width)

def to_excel_divergencias_filtradas(df_filtrado, filtros, stats, generated_at):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        wb = w.book

        fmt_title = wb.add_format({"bold": True, "font_size": 14, "font_color": "#0F172A"})
        fmt_k = wb.add_format({"bold": True, "font_size": 10, "font_color": "#334155"})
        fmt_info = wb.add_format({"font_size": 10, "font_color": "#334155"})
        fmt_hdr = wb.add_format({"bold": True, "border": 1, "align": "center", "valign": "vcenter", "bg_color": "#DBEAFE", "font_color": "#0F172A"})
        fmt_txt = wb.add_format({"border": 1})
        fmt_date = wb.add_format({"num_format": "dd/mm/yyyy", "border": 1})
        fmt_money = wb.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00', "border": 1})
        fmt_money_big = wb.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00', "bold": True})

        sh = "Divergencias"
        df = df_filtrado.copy()

        # tipos
        if "DATA" in df.columns:
            df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
        if "VALOR" in df.columns:
            df["VALOR"] = df["VALOR"].map(normalize_money)

        # booleans -> Sim/Não (V6)
        for c in ["CONFIRMADO", "RESOLVIDO"]:
            if c in df.columns:
                df[c] = df[c].fillna(False).map(bool_to_sim_nao)

        # topo
        ws = wb.add_worksheet(sh)
        w.sheets[sh] = ws

        ws.write(0, 0, "ConciliaMais — Divergências (Excel igual à tela)", fmt_title)
        ws.write(1, 0, "Processado em:", fmt_k)
        ws.write(1, 1, generated_at, fmt_info)

        ws.write(2, 0, "Origem:", fmt_k);       ws.write(2, 1, filtros.get("origem", "Todas"), fmt_info)
        ws.write(3, 0, "Visualização:", fmt_k); ws.write(3, 1, filtros.get("ver", "Todas"), fmt_info)
        ws.write(4, 0, "Severidade:", fmt_k);   ws.write(4, 1, filtros.get("severidade", "Todas"), fmt_info)
        ws.write(5, 0, "Busca:", fmt_k);        ws.write(5, 1, filtros.get("busca", ""), fmt_info)

        total_aberto = float(filtros.get("_total_aberto", 0.0) or 0.0)
        total_filtrado = float(df["VALOR"].sum()) if ("VALOR" in df.columns and len(df)) else 0.0

        ws.write(2, 6, "Total do filtro:", fmt_k)
        ws.write_number(2, 7, total_filtrado, fmt_money_big)
        ws.write(3, 6, "Total em aberto:", fmt_k)
        ws.write_number(3, 7, total_aberto, fmt_money_big)

        start_row_table = 8
        start_col_table = 0

        # escreve dataframe (sem index) e cria coluna ID manual para tabela
        df2 = df.copy().reset_index(drop=True)
        df2.insert(0, "ID", np.arange(1, len(df2) + 1))

        # header
        for j, col in enumerate(df2.columns):
            ws.write(start_row_table, start_col_table + j, col, fmt_hdr)
        ws.set_row(start_row_table, 22)

        # dados
        for r in range(len(df2)):
            excel_r = start_row_table + 1 + r
            for j, col in enumerate(df2.columns):
                val = df2.iloc[r, j]
                if col == "DATA":
                    if pd.notna(val):
                        ws.write_datetime(excel_r, start_col_table + j, val.to_pydatetime(), fmt_date)
                    else:
                        ws.write_blank(excel_r, start_col_table + j, None, fmt_date)
                elif col == "VALOR":
                    if pd.notna(val):
                        ws.write_number(excel_r, start_col_table + j, float(val), fmt_money)
                    else:
                        ws.write_blank(excel_r, start_col_table + j, None, fmt_money)
                elif col == "ID":
                    ws.write_number(excel_r, start_col_table + j, int(val), fmt_txt)
                else:
                    ws.write(excel_r, start_col_table + j, "" if pd.isna(val) else str(val), fmt_txt)

        # tabela excel com total row (SUBTOTAL)
        nrows = len(df2)
        ncols = len(df2.columns)
        if nrows > 0:
            last_row = start_row_table + nrows
            last_col = start_col_table + ncols - 1

            columns = []
            for col in df2.columns:
                if col == "VALOR":
                    columns.append({"header": col, "total_function": "sum"})
                else:
                    columns.append({"header": col})
            ws.add_table(
                start_row_table, start_col_table, last_row + 1, last_col,
                {
                    "style": "Table Style Medium 9",
                    "columns": columns,
                    "autofilter": True,
                    "total_row": True,
                }
            )
            ws.write(last_row + 1, start_col_table + 0, "TOTAL (dinâmico por filtro)", wb.add_format({"bold": True}))

        ws.freeze_panes(start_row_table + 1, 0)
        _autofit_worksheet(ws, df2, start_col_table)

        # aba resumo
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

# =========================================================
# PDF Resumo
# =========================================================
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

    # Top 10 em aberto
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
            str(r.get("NUCLEO","Não identificado") or "Não identificado"),
        ])

    t_top = Table(top_rows, colWidths=[22, 95, 60, 170, 75, 80])
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

    doc.build(story)
    buf.seek(0)
    return buf

# =========================================================
# State (V6: filtros em session_state para legenda interativa)
# =========================================================
if "page" not in st.session_state:
    st.session_state.page = "upload"
if "results" not in st.session_state:
    st.session_state.results = None
if "div_master" not in st.session_state:
    st.session_state.div_master = None
if "upload_step" not in st.session_state:
    st.session_state.upload_step = 1

# filtros persistentes (usados por legenda e pelos selects)
if "f_origem" not in st.session_state:
    st.session_state.f_origem = "Todas"
if "f_ver" not in st.session_state:
    st.session_state.f_ver = "Somente em aberto"
if "f_sev" not in st.session_state:
    st.session_state.f_sev = "Todas"
if "f_busca" not in st.session_state:
    st.session_state.f_busca = ""

# severidade thresholds (configurável)
if "sev_normal_max" not in st.session_state:
    st.session_state.sev_normal_max = 100.0
if "sev_atencao_max" not in st.session_state:
    st.session_state.sev_atencao_max = 1000.0

def severidade(valor) -> str:
    try:
        v = abs(float(valor))
    except Exception:
        return "Normal"
    if v <= float(st.session_state.sev_normal_max):
        return "Normal"
    if v <= float(st.session_state.sev_atencao_max):
        return "Atenção"
    return "Crítica"

# =========================================================
# Sidebar Navegação (estrutura futura)
# =========================================================
with st.sidebar:
    st.markdown("## Navegação")
    mod = st.radio("Módulo", ["Financeiro", "ProCV"], index=0)
    if mod == "Financeiro":
        area = st.radio("Área", ["Extrato Bancário", "Posição a Pagar", "Posição a Receber"], index=0)
    else:
        area = st.radio("Área", ["Em construção"], index=0)
    st.markdown("---")
    st.caption(f"Você está em: {mod} > {area}")

    st.markdown("### Severidade (regra)")
    st.caption("Baseada no valor absoluto da divergência.")
    cA, cB = st.columns(2)
    with cA:
        st.number_input("Normal até", min_value=0.0, value=float(st.session_state.sev_normal_max), step=10.0, key="sev_normal_max")
    with cB:
        st.number_input("Atenção até", min_value=0.0, value=float(st.session_state.sev_atencao_max), step=50.0, key="sev_atencao_max")
    st.caption("Crítica = acima do limite de Atenção.")

# Apenas Financeiro > Extrato Bancário ativo
if mod != "Financeiro" or area != "Extrato Bancário":
    st.title("ConciliaMais")
    st.info("Esta área ainda está em construção. Por enquanto, use Financeiro > Extrato Bancário.")
    st.stop()

# =========================================================
# Página: Upload (Wizard)
# =========================================================
if st.session_state.page == "upload":
    st.title("ConciliaMais — Conferência de Extrato Bancário")
    st.markdown('<div class="cm-breadcrumb">Financeiro  ›  Extrato Bancário</div>', unsafe_allow_html=True)
    st.caption("Extrato Financeiro + Razão Contábil → Match automático → Divergências → Tratativa")

    with st.container():
        st.markdown('<div class="cm-shell">', unsafe_allow_html=True)
        st.markdown("### Etapas")
        st.progress(min(max((st.session_state.upload_step - 1) / 3, 0.0), 1.0))
        st.markdown("<div class='cm-help'>1) Upload  |  2) Mapeamento  |  3) Validação de saldos  |  4) Processar</div>", unsafe_allow_html=True)
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
    st.caption("Dica: se o sistema não achar VALOR único, use Entradas/Saídas no Financeiro e Débito/Crédito no Contábil.")
    a, b = st.columns(2)

    with a:
        st.markdown("#### Extrato Financeiro")
        fin_date = st.selectbox("Data", fin_df.columns, index=fin_df.columns.get_loc(fin_guess["date"]) if fin_guess["date"] in fin_df.columns else 0)
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
        led_date = st.selectbox("Data", led_df.columns, index=led_df.columns.get_loc(led_guess["date"]) if led_guess["date"] in led_df.columns else 0, key="ld")
        led_historico = st.selectbox("Histórico", ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["historico"]) if led_guess["historico"] in led_df.columns else 0, key="lh")
        led_doc = st.selectbox("Documento/Lote", ["(nenhuma)"] + list(led_df.columns),
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

    st.session_state.upload_step = 3
    st.markdown("### 3) Validação de saldos (Saldo anterior)")

    f_norm, l_norm = build_normalized(fin_df, led_df, cfg)
    saldo_ant_fin = compute_saldo_anterior(f_norm)
    saldo_ant_led = compute_saldo_anterior(l_norm)
    diff_ant = np.nan if (pd.isna(saldo_ant_fin) or pd.isna(saldo_ant_led)) else round(saldo_ant_fin - saldo_ant_led, 2)

    proceed_ok = True
    proceed_checkbox = False

    if pd.isna(diff_ant):
        st.info("Não foi possível calcular saldo anterior automaticamente. Se existir saldo nos dois arquivos, selecione a coluna de saldo.")
    else:
        if abs(diff_ant) > 0.01:
            st.markdown(
                f"""
<div class="cm-banner">
  <strong>Atenção: saldo anterior não bate</strong>
  <div class="muted">Diferença (Financeiro - Contábil): <b>R$ {fmt(diff_ant)}</b>. Isso pode indicar divergências de períodos anteriores.</div>
</div>
""",
                unsafe_allow_html=True,
            )
            proceed_checkbox = st.checkbox("Prosseguir mesmo assim", value=False)
            proceed_ok = bool(proceed_checkbox)
        else:
            st.success("Saldo anterior bate (OK).")

    date_tol = st.number_input("Tolerância de dias para match por data (0 = mesma data)", min_value=0, max_value=10, value=0, step=1)

    st.session_state.upload_step = 4

    st.markdown('<div class="cm-help">Ao processar, o sistema gera divergências e habilita tratativa (Confirmado, Status, Observação).</div>', unsafe_allow_html=True)
    with st.form("form_processar", clear_on_submit=False):
        colb1, colb2 = st.columns([1.2, 2.0])
        with colb1:
            submit = st.form_submit_button("Processar e ir para Resultados", type="primary", disabled=not proceed_ok)
        with colb2:
            if not proceed_ok and not pd.isna(diff_ant):
                st.markdown('<div class="cm-help">Para liberar o botão, marque <b>Prosseguir mesmo assim</b>.</div>', unsafe_allow_html=True)

    if submit:
        with st.spinner("Processando..."):
            div, stats = reconcile(fin_df, led_df, cfg, date_tol_days=int(date_tol))

        generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # limpa e normaliza
        div["VALOR"] = div["VALOR"].map(normalize_money)
        div = div[div["VALOR"].notna()].copy()
        div = div[div["VALOR"].abs() > 1e-12].copy()

        for c in ["DOCUMENTO", "PREFIXO_TITULO", "HISTORICO_OPERACAO", "CHAVE_DOC"]:
            if c in div.columns:
                div[c] = div[c].replace({np.nan: "", "nan": "", "None": ""}).astype(str).str.strip()

        # Documento preferencial do financeiro: prefixo/título
        mask_fin = div["ORIGEM"].eq("Somente Financeiro")
        if "PREFIXO_TITULO" in div.columns and "DOCUMENTO" in div.columns:
            div.loc[mask_fin, "DOCUMENTO"] = div.loc[mask_fin, "PREFIXO_TITULO"].where(
                div.loc[mask_fin, "PREFIXO_TITULO"].astype(str).str.len() > 0,
                div.loc[mask_fin, "DOCUMENTO"],
            )

        # Documento do contábil extraído do histórico quando faltar
        mask_led = div["ORIGEM"].eq("Somente Contábil")
        if "HISTORICO_OPERACAO" in div.columns and "DOCUMENTO" in div.columns:
            missing = mask_led & (div["DOCUMENTO"].astype(str).str.strip().eq(""))
            div.loc[missing, "DOCUMENTO"] = div.loc[missing, "HISTORICO_OPERACAO"].map(extract_doc_from_ledger_history)

        # drop colunas que não precisam ir pra tela
        for dropc in ["PREFIXO_TITULO", "CONTA"]:
            if dropc in div.columns:
                div = div.drop(columns=[dropc])

        # Núcleo sugerido + flags
        div["NUCLEO_SUGERIDO"] = [suggest_nucleo(r) for _, r in div.iterrows()]
        div["CONFIRMADO"] = False
        div["NUCLEO"] = "Não identificado"  # só assume sugerido quando confirmar
        div["STATUS"] = "Pendente"
        div["RESOLVIDO"] = False
        div["OBS_USUARIO"] = ""
        div["SEVERIDADE"] = div["VALOR"].map(severidade)
        div["SELECIONADO"] = False

        div = div.reset_index(drop=True)
        div.index = np.arange(1, len(div) + 1)

        st.session_state.results = {"stats": stats, "generated_at": generated_at}
        st.session_state.div_master = div
        st.session_state.page = "resultados"
        st.rerun()

# =========================================================
# Página: Resultados
# =========================================================
else:
    if not st.session_state.results or st.session_state.div_master is None:
        st.session_state.page = "upload"
        st.rerun()

    stats = st.session_state.results["stats"]
    generated_at = st.session_state.results["generated_at"]

    topbar = st.columns([1.4, 1.0, 1.0])
    with topbar[0]:
        st.title("Resultados — ConciliaMais (Módulo 1)")
        st.markdown('<div class="cm-breadcrumb">Financeiro  ›  Extrato Bancário</div>', unsafe_allow_html=True)
        st.caption(f"Processado em: {generated_at}")
    with topbar[2]:
        if st.button("← Voltar para Upload", use_container_width=True):
            st.session_state.page = "upload"
            st.session_state.upload_step = 1
            st.rerun()

    div_master = st.session_state.div_master.copy()
    div_master["VALOR"] = div_master["VALOR"].map(normalize_money)
    div_master["RESOLVIDO"] = div_master["RESOLVIDO"].fillna(False)
    div_master["STATUS"] = div_master["STATUS"].fillna("Pendente").astype(str)
    div_master["CONFIRMADO"] = div_master.get("CONFIRMADO", False)
    div_master["NUCLEO"] = div_master.get("NUCLEO", "Não identificado").fillna("Não identificado")

    # coerências
    if "NUCLEO_SUGERIDO" in div_master.columns:
        need = div_master["CONFIRMADO"] & (div_master["NUCLEO"].astype(str).str.strip().eq("") | div_master["NUCLEO"].eq("Não identificado"))
        div_master.loc[need, "NUCLEO"] = div_master.loc[need, "NUCLEO_SUGERIDO"].fillna("Não identificado")

    div_master.loc[div_master["RESOLVIDO"], "STATUS"] = "Resolvido"

    if "SEVERIDADE" not in div_master.columns:
        div_master["SEVERIDADE"] = div_master["VALOR"].map(severidade)
    else:
        # recalcula se thresholds mudaram
        div_master["SEVERIDADE"] = div_master["VALOR"].map(severidade)

    if "SELECIONADO" not in div_master.columns:
        div_master["SELECIONADO"] = False

    st.session_state.div_master = div_master

    resolved_mask = div_master["RESOLVIDO"] | (div_master["STATUS"].str.lower().eq("resolvido"))
    total_itens = len(div_master)
    itens_res = int(resolved_mask.sum())
    itens_ab = int(total_itens - itens_res)
    valor_aberto = float(div_master.loc[~resolved_mask, "VALOR"].sum()) if total_itens else 0.0
    pct_res = (itens_res / total_itens * 100.0) if total_itens else 0.0

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

    # =========================
    # Resumo (Top + Distribuições + Legenda interativa)
    # =========================
    with st.expander("Resumo para priorização (abertos, top impacto, distribuições)", expanded=True):
        df_open = div_master.loc[~resolved_mask].copy()
        df_open["ABS"] = df_open["VALOR"].abs()
        top_open = df_open.sort_values("ABS", ascending=False).head(10)

        left, right = st.columns([2.2, 1.0], gap="large")
        with left:
            st.markdown("**Top 10 em aberto por impacto**")
            show_cols = ["ORIGEM", "DATA", "DOCUMENTO", "VALOR", "SEVERIDADE", "NUCLEO"]
            st.dataframe(top_open[show_cols].copy(), use_container_width=True, height=320)

        with right:
            st.markdown("**Legenda (clique para filtrar)**")
            cL1, cL2 = st.columns(2)
            with cL1:
                if st.button("Somente Financeiro", use_container_width=True):
                    st.session_state.f_origem = "Somente Financeiro"
                    st.session_state.f_ver = "Somente em aberto"
                    st.rerun()
                if st.button("Normal", use_container_width=True):
                    st.session_state.f_sev = "Normal"
                    st.session_state.f_ver = "Somente em aberto"
                    st.rerun()
                st.markdown(origem_tag("Somente Financeiro"), unsafe_allow_html=True)
                st.markdown(severidade_tag("Normal"), unsafe_allow_html=True)
            with cL2:
                if st.button("Somente Contábil", use_container_width=True):
                    st.session_state.f_origem = "Somente Contábil"
                    st.session_state.f_ver = "Somente em aberto"
                    st.rerun()
                if st.button("Atenção", use_container_width=True):
                    st.session_state.f_sev = "Atenção"
                    st.session_state.f_ver = "Somente em aberto"
                    st.rerun()
                st.markdown(origem_tag("Somente Contábil"), unsafe_allow_html=True)
                st.markdown(severidade_tag("Atenção"), unsafe_allow_html=True)

            if st.button("Crítica", use_container_width=True):
                st.session_state.f_sev = "Crítica"
                st.session_state.f_ver = "Somente em aberto"
                st.rerun()
            st.markdown(severidade_tag("Crítica"), unsafe_allow_html=True)

            if st.button("Limpar filtros", use_container_width=True):
                st.session_state.f_origem = "Todas"
                st.session_state.f_ver = "Somente em aberto"
                st.session_state.f_sev = "Todas"
                st.session_state.f_busca = ""
                st.rerun()

        st.markdown('<div class="cm-divider"></div>', unsafe_allow_html=True)

        # Distribuições
        st.markdown("**Distribuição por Origem (abertos)**")
        if len(df_open):
            dist_origem = df_open.groupby("ORIGEM", dropna=False).agg(Itens=("VALOR","size"), Valor=("VALOR","sum")).reset_index().sort_values("Valor", ascending=False)
            st.dataframe(dist_origem, use_container_width=True, height=160)

            st.markdown("**Distribuição por Origem × Núcleo (abertos)**")
            dist_on = df_open.groupby(["ORIGEM","NUCLEO"], dropna=False).agg(Itens=("VALOR","size"), Valor=("VALOR","sum")).reset_index().sort_values(["ORIGEM","Valor"], ascending=[True, False])
            st.dataframe(dist_on, use_container_width=True, height=220)

            st.markdown("**Distribuição por Severidade (abertos)**")
            dist_sev = df_open.groupby("SEVERIDADE", dropna=False).agg(Itens=("VALOR","size"), Valor=("VALOR","sum")).reset_index().sort_values("Valor", ascending=False)
            st.dataframe(dist_sev, use_container_width=True, height=160)

            st.markdown("**Distribuição por Origem × Severidade (abertos)**")
            dist_os = df_open.groupby(["ORIGEM","SEVERIDADE"], dropna=False).agg(Itens=("VALOR","size"), Valor=("VALOR","sum")).reset_index().sort_values(["ORIGEM","Valor"], ascending=[True, False])
            st.dataframe(dist_os, use_container_width=True, height=200)
        else:
            st.info("Sem pendências em aberto.")

    # =========================
    # Filtros (V6: grudados na tabela)
    # =========================
    st.markdown("### Tratativa (tabela)")
    st.markdown('<div class="cm-help">Fluxo: 1) Ajuste filtros  2) Marque itens (Selecionado)  3) Faça ações em massa (abaixo) ou edite linha a linha.</div>', unsafe_allow_html=True)

    f1, f2, f3, f4, f5 = st.columns([1.1, 1.1, 1.1, 2.3, 1.0], gap="large")
    with f1:
        origem = st.selectbox("Origem", ["Todas", "Somente Financeiro", "Somente Contábil"], key="ui_origem", index=["Todas","Somente Financeiro","Somente Contábil"].index(st.session_state.f_origem))
    with f2:
        ver = st.selectbox("Visualizar", ["Todas", "Somente em aberto", "Somente resolvidas"], key="ui_ver", index=["Todas","Somente em aberto","Somente resolvidas"].index(st.session_state.f_ver))
    with f3:
        sev = st.selectbox("Severidade", ["Todas", "Normal", "Atenção", "Crítica"], key="ui_sev", index=["Todas","Normal","Atenção","Crítica"].index(st.session_state.f_sev))
    with f4:
        busca = st.text_input("Buscar (documento, histórico, chave, núcleo e valor)", value=st.session_state.f_busca, key="ui_busca")
    with f5:
        st.markdown("<div style='height:1px'></div>", unsafe_allow_html=True)

    # persiste filtros
    st.session_state.f_origem = origem
    st.session_state.f_ver = ver
    st.session_state.f_sev = sev
    st.session_state.f_busca = busca

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

        cols_search = ["DOCUMENTO", "HISTORICO_OPERACAO", "CHAVE_DOC", "NUCLEO", "SEVERIDADE", "ORIGEM", "STATUS", "OBS_USUARIO"]
        mask = pd.Series([False] * len(df), index=df.index)

        for c in cols_search:
            if c in df.columns:
                mask = mask | df[c].astype(str).str.lower().str.contains(q, na=False)

        # inclui VALOR na busca (V6)
        if "VALOR" in df.columns:
            tokens = df["VALOR"].map(_money_search_tokens).astype(str).str.lower()
            mask = mask | tokens.str.contains(q, na=False)

        df = df[mask].copy()

    total_filtrado = float(df["VALOR"].sum()) if not df.empty else 0.0
    with f5:
        st.markdown(
            f"""
<div class="cm-mini">
  <div class="k">Total do filtro</div>
  <div class="v">{fmt(total_filtrado)}</div>
</div>
""",
            unsafe_allow_html=True,
        )

    if "DATA" in df.columns:
        df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df.sort_values(by=["DATA", "VALOR"], ascending=[True, True])

    # =========================
    # Tabela (data_editor)
    # =========================
    view_cols = [
        "SELECIONADO",
        "ORIGEM", "SEVERIDADE", "DATA", "DOCUMENTO", "HISTORICO_OPERACAO", "CHAVE_DOC", "VALOR",
        "NUCLEO_SUGERIDO",
        "CONFIRMADO", "NUCLEO",
        "STATUS", "RESOLVIDO", "OBS_USUARIO"
    ]
    df_view = df[view_cols].copy()
    df_view_display = df_view.copy()
    df_view_display["DATA"] = pd.to_datetime(df_view_display["DATA"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")

    column_config = {
        "SELECIONADO": st.column_config.CheckboxColumn(help="Marque para ações em massa."),
        "ORIGEM": st.column_config.TextColumn(disabled=True),
        "SEVERIDADE": st.column_config.TextColumn(disabled=True, help=f"Normal <= {st.session_state.sev_normal_max} | Atenção <= {st.session_state.sev_atencao_max} | Crítica > {st.session_state.sev_atencao_max}"),
        "DATA": st.column_config.TextColumn(disabled=True),
        "DOCUMENTO": st.column_config.TextColumn(disabled=True),
        "HISTORICO_OPERACAO": st.column_config.TextColumn(disabled=True),
        "CHAVE_DOC": st.column_config.TextColumn(disabled=True),
        "VALOR": st.column_config.NumberColumn(format="R$ %.2f", disabled=True),
        "NUCLEO_SUGERIDO": st.column_config.TextColumn(disabled=True),
        "CONFIRMADO": st.column_config.CheckboxColumn(),
        "NUCLEO": st.column_config.SelectboxColumn(options=NUCLEOS),
        "STATUS": st.column_config.SelectboxColumn(options=STATUS_OPTS),
        "RESOLVIDO": st.column_config.CheckboxColumn(),
        "OBS_USUARIO": st.column_config.TextColumn(),
    }

    edited = st.data_editor(
        df_view_display,
        use_container_width=True,
        height=520,
        column_config=column_config,
        key="editor_tratativa_v6",
        hide_index=False,
    )

    # Aplicar mudanças
    if edited is not None and len(edited) == len(df_view_display):
        to_update = edited.copy()

        # confirmando -> se núcleo vazio, assume sugerido
        if "NUCLEO_SUGERIDO" in to_update.columns:
            to_update["NUCLEO"] = to_update["NUCLEO"].fillna("Não identificado").replace("", "Não identificado")
            need = to_update["CONFIRMADO"].fillna(False) & (to_update["NUCLEO"].astype(str).str.strip().eq("") | to_update["NUCLEO"].eq("Não identificado"))
            to_update.loc[need, "NUCLEO"] = to_update.loc[need, "NUCLEO_SUGERIDO"].fillna("Não identificado")

        # resolvido -> status
        res_col = to_update["RESOLVIDO"].fillna(False)
        to_update.loc[res_col, "STATUS"] = "Resolvido"

        upd_cols = ["SELECIONADO", "CONFIRMADO", "NUCLEO", "STATUS", "RESOLVIDO", "OBS_USUARIO"]
        dm = st.session_state.div_master.copy()
        for c in upd_cols:
            dm.loc[to_update.index, c] = to_update[c].values

        dm["SEVERIDADE"] = dm["VALOR"].map(severidade)
        st.session_state.div_master = dm
        div_master = dm.copy()

    # =========================
    # Ações em massa (V6: depois da tabela, como você pediu)
    # =========================
    st.markdown("### Ações em massa")
    st.markdown('<div class="cm-help">Use quando marcar vários itens em “Selecionado”.</div>', unsafe_allow_html=True)

    ids_filtrados = list(df.index)
    dm0 = st.session_state.div_master.copy()
    selecionados_count = int(dm0.loc[ids_filtrados, "SELECIONADO"].fillna(False).sum()) if len(ids_filtrados) else 0

    with st.container():
        st.markdown('<div class="cm-actionbar">', unsafe_allow_html=True)
        a1, a2, a3 = st.columns([1.25, 1.25, 2.0], gap="large")
        with a1:
            if st.button("Selecionar todos do filtro", use_container_width=True):
                dm = st.session_state.div_master.copy()
                dm.loc[ids_filtrados, "SELECIONADO"] = True
                st.session_state.div_master = dm
                st.rerun()
        with a2:
            if st.button("Limpar seleção do filtro", use_container_width=True):
                dm = st.session_state.div_master.copy()
                dm.loc[ids_filtrados, "SELECIONADO"] = False
                st.session_state.div_master = dm
                st.rerun()
        with a3:
            st.markdown(f'<span class="cm-badge">Selecionados no filtro: {selecionados_count}</span>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    scope = st.radio("Aplicar em:", ["Selecionados", "Todos do filtro"], horizontal=True)
    target_ids = list(dm0.index[dm0["SELECIONADO"].fillna(False)]) if scope == "Selecionados" else ids_filtrados

    bA, bB, bC, bD, bE = st.columns([1.0, 1.0, 1.2, 1.8, 1.0], gap="large")
    with bA:
        bulk_confirm = st.selectbox("Confirmado", ["(não alterar)", "Sim", "Não"])
    with bB:
        bulk_resolvido = st.selectbox("Resolvido", ["(não alterar)", "Sim", "Não"])
    with bC:
        bulk_status = st.selectbox("Status", ["(não alterar)"] + STATUS_OPTS)
    with bD:
        bulk_obs = st.text_input("OBS (opcional)", value="")
    with bE:
        do_apply = st.button("Aplicar", type="primary", disabled=(len(target_ids) == 0), use_container_width=True)

    if do_apply:
        dm = st.session_state.div_master.copy()

        if bulk_confirm != "(não alterar)":
            if bulk_confirm == "Sim":
                dm.loc[target_ids, "CONFIRMADO"] = True
                if "NUCLEO_SUGERIDO" in dm.columns:
                    dm.loc[target_ids, "NUCLEO"] = dm.loc[target_ids, "NUCLEO_SUGERIDO"].fillna("Não identificado")
            else:
                dm.loc[target_ids, "CONFIRMADO"] = False
                dm.loc[target_ids, "NUCLEO"] = "Não identificado"

        if bulk_obs.strip():
            dm.loc[target_ids, "OBS_USUARIO"] = bulk_obs.strip()

        if bulk_status != "(não alterar)":
            dm.loc[target_ids, "STATUS"] = bulk_status

        if bulk_resolvido != "(não alterar)":
            if bulk_resolvido == "Sim":
                dm.loc[target_ids, "RESOLVIDO"] = True
                dm.loc[target_ids, "STATUS"] = "Resolvido"
            else:
                dm.loc[target_ids, "RESOLVIDO"] = False
                dm.loc[target_ids, "STATUS"] = dm.loc[target_ids, "STATUS"].replace({"Resolvido": "Pendente"})

        dm.loc[target_ids, "SELECIONADO"] = False
        dm["SEVERIDADE"] = dm["VALOR"].map(severidade)

        st.session_state.div_master = dm
        st.success(f"Ação aplicada em {len(target_ids)} itens.")
        st.rerun()

    # =========================
    # Detalhe do item
    # =========================
    st.markdown("### Detalhe do item")
    pick_id = st.number_input(
        "Digite o ID do item para ver detalhes",
        min_value=1,
        max_value=max(1, int(st.session_state.div_master.index.max())),
        value=1,
        step=1
    )
    dm_now = st.session_state.div_master
    if pick_id in dm_now.index:
        r = dm_now.loc[pick_id]
        dt_txt = ""
        try:
            if pd.notna(r.get("DATA")):
                dt_txt = pd.to_datetime(r.get("DATA")).strftime("%d/%m/%Y")
        except Exception:
            dt_txt = str(r.get("DATA") or "")

        confirmado_txt = "Sim" if bool(r.get("CONFIRMADO", False)) else "Não"

        st.markdown(
            f"""
<div class="cm-detail">
  <div class="title">Item #{pick_id}</div>
  <div class="row"><span class="label">Origem:</span> <span class="val">{r.get('ORIGEM','')}</span></div>
  <div class="row"><span class="label">Severidade:</span> <span class="val">{r.get('SEVERIDADE','')}</span></div>
  <div class="row"><span class="label">Data:</span> <span class="val">{dt_txt}</span></div>
  <div class="row"><span class="label">Documento:</span> <span class="val">{r.get('DOCUMENTO','')}</span></div>
  <div class="row"><span class="label">Chave:</span> <span class="val">{r.get('CHAVE_DOC','')}</span></div>
  <div class="row"><span class="label">Valor:</span> <span class="val">{fmt(r.get('VALOR', np.nan))}</span></div>
  <div class="row"><span class="label">Núcleo sugerido:</span> <span class="val">{r.get('NUCLEO_SUGERIDO','')}</span></div>
  <div class="row"><span class="label">Confirmado:</span> <span class="val">{confirmado_txt}</span></div>
  <div class="row"><span class="label">Núcleo:</span> <span class="val">{r.get('NUCLEO','')}</span></div>
  <div class="row"><span class="label">Status:</span> <span class="val">{r.get('STATUS','')}</span></div>
  <div class="row"><span class="label">Resolvido:</span> <span class="val">{'Sim' if bool(r.get('RESOLVIDO', False)) else 'Não'}</span></div>
</div>
""",
            unsafe_allow_html=True,
        )

        resumo = (
            f"ID: {pick_id}\n"
            f"ORIGEM: {r.get('ORIGEM','')}\n"
            f"SEVERIDADE: {r.get('SEVERIDADE','')}\n"
            f"DATA: {dt_txt}\n"
            f"DOCUMENTO: {r.get('DOCUMENTO','')}\n"
            f"CHAVE: {r.get('CHAVE_DOC','')}\n"
            f"VALOR: {fmt(r.get('VALOR', np.nan))}\n"
            f"NUCLEO_SUGERIDO: {r.get('NUCLEO_SUGERIDO','')}\n"
            f"CONFIRMADO: {confirmado_txt}\n"
            f"NUCLEO: {r.get('NUCLEO','')}\n"
            f"STATUS: {r.get('STATUS','')}\n"
            f"RESOLVIDO: {'Sim' if bool(r.get('RESOLVIDO', False)) else 'Não'}\n"
            f"OBS: {r.get('OBS_USUARIO','')}\n"
            f"HISTÓRICO: {r.get('HISTORICO_OPERACAO','')}"
        )
        st.text_area("Copiar resumo (e-mail/ticket)", value=resumo, height=210)

    # =========================
    # Export
    # =========================
    st.markdown("### Exportar")
    filtros = {"origem": origem, "ver": ver, "severidade": sev, "busca": busca.strip(), "_total_aberto": valor_aberto}

    # exporta exatamente como filtrado (sem SELECIONADO)
    df_export = df_view.drop(columns=["SELECIONADO"]).copy()

    excel_bytes = to_excel_divergencias_filtradas(df_filtrado=df_export, filtros=filtros, stats=stats, generated_at=generated_at)
    pdf_bytes = to_pdf_resumo(stats, generated_at, st.session_state.div_master)

    exA, exB, exC = st.columns([1.9, 1.9, 1.2], gap="large")
    with exA:
        st.download_button(
            "Baixar Divergências (Excel) — como filtrado",
            data=excel_bytes,
            file_name=f"ConciliaMais_DivergenciasFiltradas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with exB:
        st.download_button(
            "Baixar Relatório Resumo (PDF) — executivo",
            data=pdf_bytes,
            file_name=f"ConciliaMais_Resumo_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
    with exC:
        if st.button("Limpar e recomeçar", use_container_width=True):
            st.session_state.results = None
            st.session_state.div_master = None
            st.session_state.page = "upload"
            st.session_state.upload_step = 1
            st.rerun()
