import streamlit as st
import pandas as pd
import numpy as np
import re
import json
import os
import unicodedata
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

RULES_FILE = "regras.json"

# =========================================================
# CSS (Dark + Marca Azul)
# =========================================================
st.markdown(
    """
<style>
:root{
  --bg: #0B1220;
  --card: #0F172A;
  --border: rgba(148,163,184,.18);
  --text: #E5E7EB;
  --muted: rgba(226,232,240,.72);
  --primary: #2563EB;
  --primary2:#1D4ED8;
  --shadow: 0 10px 24px rgba(0,0,0,.35);
}

html, body, [class*="css"]  { color: var(--text) !important; }
body { background: var(--bg) !important; }

.block-container {
  padding-top: 1.0rem;
  padding-bottom: 2.2rem;
  max-width: 1450px;
}

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
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# Regras / Persistência
# =========================================================
def default_rules_payload():
    return {"nucleo_rules": [], "criticidade_rules": []}

def load_rules():
    if not os.path.exists(RULES_FILE):
        payload = default_rules_payload()
        save_rules(payload)
        return payload
    try:
        with open(RULES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            data = default_rules_payload()
        data.setdefault("nucleo_rules", [])
        data.setdefault("criticidade_rules", [])
        return data
    except Exception:
        payload = default_rules_payload()
        save_rules(payload)
        return payload

def save_rules(payload):
    with open(RULES_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

def next_rule_id(rules):
    used = []
    for r in rules:
        try:
            used.append(int(r.get("id", 0)))
        except Exception:
            pass
    return (max(used) + 1) if used else 1

def add_rule(rule_type, rule_dict):
    payload = load_rules()
    bucket = "nucleo_rules" if rule_type == "nucleo" else "criticidade_rules"
    rule_dict = dict(rule_dict)
    rule_dict["id"] = next_rule_id(payload[bucket])
    payload[bucket].append(rule_dict)
    payload[bucket] = sorted(payload[bucket], key=lambda x: (int(x.get("prioridade", 9999)), int(x.get("id", 0))))
    save_rules(payload)

def update_rule_status(rule_type, rule_id, active):
    payload = load_rules()
    bucket = "nucleo_rules" if rule_type == "nucleo" else "criticidade_rules"
    for r in payload[bucket]:
        if int(r.get("id", 0)) == int(rule_id):
            r["ativa"] = bool(active)
            break
    save_rules(payload)

def delete_rule(rule_type, rule_id):
    payload = load_rules()
    bucket = "nucleo_rules" if rule_type == "nucleo" else "criticidade_rules"
    payload[bucket] = [r for r in payload[bucket] if int(r.get("id", 0)) != int(rule_id)]
    save_rules(payload)

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

# =========================================================
# Núcleo / Severidade base
# =========================================================
NUCLEOS = ["Processo interno", "Cadastro", "Configuração RP", "Não identificado"]
STATUS_OPTS = ["Pendente", "Em análise", "Resolvido"]
SEVERIDADES = ["Normal", "Atenção", "Crítica"]

def suggest_nucleo_base(row):
    origem = str(row.get("ORIGEM", "")).lower()
    hist = str(row.get("HISTORICO_OPERACAO", "")).lower()
    doc = str(row.get("DOCUMENTO", "")).strip()

    if any(k in hist for k in ["cancelamento de baixa", "canc baixa", "estorno de baixa", "estorno baixa", "canc. baixa"]):
        return "Processo interno"

    if any(k in hist for k in ["baixa", "liquidação", "liquidacao", "pagamento", "pagto", "estorno"]):
        return "Processo interno"

    if "somente financeiro" in origem and (doc != "" or "mov" in hist or "titulo" in hist or "título" in hist):
        return "Cadastro"

    if any(k in hist for k in ["rp", "reprocess", "rotina", "processamento", "integracao", "integração"]):
        return "Configuração RP"

    return "Não identificado"

def severidade_base(valor) -> str:
    try:
        v = abs(float(valor))
    except Exception:
        return "Normal"
    if v <= 100:
        return "Normal"
    if v <= 1000:
        return "Atenção"
    return "Crítica"

# =========================================================
# Motivo base / Regras
# =========================================================
def strip_accents(text):
    if text is None:
        return ""
    text = str(text)
    return "".join(ch for ch in unicodedata.normalize("NFD", text) if unicodedata.category(ch) != "Mn")

def normalize_text_rule(text):
    text = strip_accents(str(text).lower().strip())
    text = re.sub(r"\b\d{6,}\b", " ", text)
    text = re.sub(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", " ", text)
    text = re.sub(r"\b\d+\b", " ", text)
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

STOPWORDS_MOTIVO = {
    "de", "da", "do", "das", "dos", "para", "por", "com", "sem", "na", "no",
    "em", "a", "o", "e", "ou", "um", "uma", "ao", "aos", "as", "os"
}

def build_motivo_base(text):
    txt = normalize_text_rule(text)
    toks = [t for t in txt.split() if t not in STOPWORDS_MOTIVO]
    toks = toks[:8]
    return " ".join(toks).strip()

def safe_float(x, default=None):
    try:
        if pd.isna(x):
            return default
        return float(x)
    except Exception:
        return default

def match_rule_value(valor, vmin, vmax):
    v = safe_float(valor, None)
    if v is None:
        return False
    if vmin not in [None, "", "nan"]:
        try:
            if abs(v) < float(vmin):
                return False
        except Exception:
            pass
    if vmax not in [None, "", "nan"]:
        try:
            if abs(v) > float(vmax):
                return False
        except Exception:
            pass
    return True

def rule_matches(row, rule):
    if not bool(rule.get("ativa", True)):
        return False

    origem = str(row.get("ORIGEM", "")).strip()
    documento = str(row.get("DOCUMENTO", "")).strip()
    hist = str(row.get("HISTORICO_OPERACAO", ""))
    hist_norm = normalize_text_rule(hist)
    doc_norm = normalize_text_rule(documento)

    rule_origem = str(rule.get("origem", "")).strip()
    if rule_origem and rule_origem != origem:
        return False

    texto_contem = str(rule.get("texto_contem", "")).strip()
    if texto_contem:
        if normalize_text_rule(texto_contem) not in hist_norm and normalize_text_rule(texto_contem) not in doc_norm:
            return False

    regex = str(rule.get("regex", "")).strip()
    if regex:
        try:
            if not re.search(regex, hist, flags=re.IGNORECASE):
                return False
        except Exception:
            return False

    documento_prefixo = str(rule.get("documento_prefixo", "")).strip()
    if documento_prefixo:
        if not documento.upper().startswith(documento_prefixo.upper()):
            return False

    valor_min = rule.get("valor_min", None)
    valor_max = rule.get("valor_max", None)
    if (valor_min not in [None, "", "nan"]) or (valor_max not in [None, "", "nan"]):
        if not match_rule_value(row.get("VALOR", np.nan), valor_min, valor_max):
            return False

    return True

def apply_rules_to_row(row, rules, default_value):
    rules_sorted = sorted(rules, key=lambda x: (int(x.get("prioridade", 9999)), int(x.get("id", 0))))
    for rule in rules_sorted:
        if rule_matches(row, rule):
            return rule.get("resultado", default_value), f"Regra #{rule.get('id')} - {rule.get('nome', '')}"
    return default_value, "Base"

def apply_classification_rules(df):
    payload = load_rules()
    nuc_rules = payload.get("nucleo_rules", [])
    crit_rules = payload.get("criticidade_rules", [])

    nuc_result = []
    nuc_trace = []
    sev_result = []
    sev_trace = []

    for _, row in df.iterrows():
        base_nuc = suggest_nucleo_base(row)
        nuc, nuc_applied = apply_rules_to_row(row, nuc_rules, base_nuc)

        base_sev = severidade_base(row.get("VALOR", np.nan))
        sev, sev_applied = apply_rules_to_row(row, crit_rules, base_sev)

        nuc_result.append(nuc)
        nuc_trace.append(nuc_applied)
        sev_result.append(sev)
        sev_trace.append(sev_applied)

    df["NUCLEO_SUGERIDO"] = nuc_result
    df["REGRA_NUCLEO_APLICADA"] = nuc_trace
    df["SEVERIDADE"] = sev_result
    df["REGRA_SEVERIDADE_APLICADA"] = sev_trace
    return df

# =========================================================
# Reconcile
# =========================================================
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

    for _, r in f.iterrows():
        fi = int(r["__idx"])
        if r["__doc_key"] and pd.notna(r["__amount"]):
            key = (round(float(r["__amount"]), 2), r["__doc_key"])
            if key in key_to_led:
                try_match(fi, key_to_led[key])

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
# Excel Export
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

        for bcol in ["CONFIRMADO", "RESOLVIDO"]:
            if bcol in df.columns:
                df[bcol] = df[bcol].fillna(False).map(lambda x: "Sim" if bool(x) else "Não")

        if "DATA" in df.columns:
            df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
        if "VALOR" in df.columns:
            df["VALOR"] = df["VALOR"].map(normalize_money)

        ws = wb.add_worksheet(sh)
        w.sheets[sh] = ws

        ws.write(0, 0, "ConciliaMais — Divergências (Excel igual à tela)", fmt_title)
        ws.write(1, 0, "Processado em:", fmt_k)
        ws.write(1, 1, generated_at, fmt_info)

        ws.write(2, 0, "Origem:", fmt_k);      ws.write(2, 1, filtros.get("origem", "Todas"), fmt_info)
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

        df2 = df.copy().reset_index(drop=True)
        df2.insert(0, "ID", np.arange(1, len(df2) + 1))

        for j, col in enumerate(df2.columns):
            ws.write(start_row_table, start_col_table + j, col, fmt_hdr)
        ws.set_row(start_row_table, 22)

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
# State
# =========================================================
if "page" not in st.session_state:
    st.session_state.page = "upload"
if "results" not in st.session_state:
    st.session_state.results = None
if "div_master" not in st.session_state:
    st.session_state.div_master = None
if "upload_step" not in st.session_state:
    st.session_state.upload_step = 1

# =========================================================
# Sidebar Navegação
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

if mod != "Financeiro" or area != "Extrato Bancário":
    st.title("ConciliaMais")
    st.info("Esta área ainda está em construção. Por enquanto, use Financeiro > Extrato Bancário.")
    st.stop()

# =========================================================
# Página: Upload
# =========================================================
if st.session_state.page == "upload":
    st.title("ConciliaMais — Conferência de Extrato Bancário")
    st.markdown('<div class="cm-breadcrumb">Financeiro  ›  Extrato Bancário</div>', unsafe_allow_html=True)
    st.caption("Extrato Financeiro + Razão Contábil → Match automático → Divergências → Tratativa")

    st.markdown("### Etapas")
    steps = ["1) Upload", "2) Mapeamento", "3) Validação", "4) Processar"]
    try:
        st.segmented_control("Fluxo", options=steps, default=steps[min(max(st.session_state.upload_step-1,0),3)], disabled=True)
    except Exception:
        st.caption(" > ".join(steps))

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
            proceed_ok = bool(st.checkbox("Prosseguir mesmo assim", value=False))
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

        div["VALOR"] = div["VALOR"].map(normalize_money)
        div = div[div["VALOR"].notna()].copy()
        div = div[div["VALOR"].abs() > 1e-12].copy()

        for c in ["DOCUMENTO", "PREFIXO_TITULO", "HISTORICO_OPERACAO", "CHAVE_DOC"]:
            if c in div.columns:
                div[c] = div[c].replace({np.nan: "", "nan": "", "None": ""}).astype(str).str.strip()

        mask_fin = div["ORIGEM"].eq("Somente Financeiro")
        if "PREFIXO_TITULO" in div.columns and "DOCUMENTO" in div.columns:
            div.loc[mask_fin, "DOCUMENTO"] = div.loc[mask_fin, "PREFIXO_TITULO"].where(
                div.loc[mask_fin, "PREFIXO_TITULO"].astype(str).str.len() > 0,
                div.loc[mask_fin, "DOCUMENTO"],
            )

        mask_led = div["ORIGEM"].eq("Somente Contábil")
        if "HISTORICO_OPERACAO" in div.columns and "DOCUMENTO" in div.columns:
            missing = mask_led & (div["DOCUMENTO"].astype(str).str.strip().eq(""))
            div.loc[missing, "DOCUMENTO"] = div.loc[missing, "HISTORICO_OPERACAO"].map(extract_doc_from_ledger_history)

        for dropc in ["PREFIXO_TITULO", "CONTA"]:
            if dropc in div.columns:
                div = div.drop(columns=[dropc])

        div["MOTIVO_BASE"] = div["HISTORICO_OPERACAO"].map(build_motivo_base)
        div = apply_classification_rules(div)

        div["CONFIRMADO"] = False
        div["NUCLEO"] = "Não identificado"
        div["STATUS"] = "Pendente"
        div["RESOLVIDO"] = False
        div["OBS_USUARIO"] = ""
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
    div_master["MOTIVO_BASE"] = div_master.get("MOTIVO_BASE", div_master["HISTORICO_OPERACAO"].map(build_motivo_base))

    if "NUCLEO_SUGERIDO" in div_master.columns:
        need = div_master["CONFIRMADO"] & (div_master["NUCLEO"].astype(str).str.strip().eq("") | div_master["NUCLEO"].eq("Não identificado"))
        div_master.loc[need, "NUCLEO"] = div_master.loc[need, "NUCLEO_SUGERIDO"].fillna("Não identificado")

    div_master.loc[div_master["RESOLVIDO"], "STATUS"] = "Resolvido"

    if "SEVERIDADE" not in div_master.columns:
        div_master = apply_classification_rules(div_master)
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

    # =====================================================
    # Biblioteca de regras
    # =====================================================
    with st.expander("Biblioteca de regras (persistente)", expanded=False):
        payload = load_rules()

        st.markdown("#### Criar regra de Núcleo")
        c1, c2, c3 = st.columns([1.4, 1.0, 1.0])
        with c1:
            nr_nome = st.text_input("Nome da regra (núcleo)", key="nr_nome")
            nr_texto = st.text_input("Texto contém", key="nr_texto")
            nr_regex = st.text_input("Regex (opcional)", key="nr_regex")
            nr_doc_pref = st.text_input("Prefixo do documento (opcional)", key="nr_doc_pref")
        with c2:
            nr_origem = st.selectbox("Origem", ["", "Somente Financeiro", "Somente Contábil"], key="nr_origem")
            nr_valor_min = st.text_input("Valor mínimo abs", key="nr_valor_min")
            nr_valor_max = st.text_input("Valor máximo abs", key="nr_valor_max")
        with c3:
            nr_resultado = st.selectbox("Resultado", NUCLEOS, key="nr_resultado")
            nr_prioridade = st.number_input("Prioridade", min_value=1, value=100, step=1, key="nr_prioridade")
            nr_ativa = st.checkbox("Ativa", value=True, key="nr_ativa")
            if st.button("Salvar regra de Núcleo", type="primary"):
                add_rule("nucleo", {
                    "nome": nr_nome.strip() or f"Núcleo {nr_resultado}",
                    "prioridade": int(nr_prioridade),
                    "ativa": bool(nr_ativa),
                    "origem": nr_origem.strip(),
                    "texto_contem": nr_texto.strip(),
                    "regex": nr_regex.strip(),
                    "documento_prefixo": nr_doc_pref.strip(),
                    "valor_min": nr_valor_min.strip(),
                    "valor_max": nr_valor_max.strip(),
                    "resultado": nr_resultado,
                })
                dm = st.session_state.div_master.copy()
                dm = apply_classification_rules(dm)
                st.session_state.div_master = dm
                st.success("Regra de núcleo salva.")
                st.rerun()

        st.markdown("---")
        st.markdown("#### Criar regra de Criticidade")
        d1, d2, d3 = st.columns([1.4, 1.0, 1.0])
        with d1:
            cr_nome = st.text_input("Nome da regra (criticidade)", key="cr_nome")
            cr_texto = st.text_input("Texto contém", key="cr_texto")
            cr_regex = st.text_input("Regex (opcional)", key="cr_regex")
            cr_doc_pref = st.text_input("Prefixo do documento (opcional)", key="cr_doc_pref")
        with d2:
            cr_origem = st.selectbox("Origem ", ["", "Somente Financeiro", "Somente Contábil"], key="cr_origem")
            cr_valor_min = st.text_input("Valor mínimo abs ", key="cr_valor_min")
            cr_valor_max = st.text_input("Valor máximo abs ", key="cr_valor_max")
        with d3:
            cr_resultado = st.selectbox("Resultado ", SEVERIDADES, key="cr_resultado")
            cr_prioridade = st.number_input("Prioridade ", min_value=1, value=100, step=1, key="cr_prioridade")
            cr_ativa = st.checkbox("Ativa ", value=True, key="cr_ativa")
            if st.button("Salvar regra de Criticidade", type="primary"):
                add_rule("criticidade", {
                    "nome": cr_nome.strip() or f"Criticidade {cr_resultado}",
                    "prioridade": int(cr_prioridade),
                    "ativa": bool(cr_ativa),
                    "origem": cr_origem.strip(),
                    "texto_contem": cr_texto.strip(),
                    "regex": cr_regex.strip(),
                    "documento_prefixo": cr_doc_pref.strip(),
                    "valor_min": cr_valor_min.strip(),
                    "valor_max": cr_valor_max.strip(),
                    "resultado": cr_resultado,
                })
                dm = st.session_state.div_master.copy()
                dm = apply_classification_rules(dm)
                st.session_state.div_master = dm
                st.success("Regra de criticidade salva.")
                st.rerun()

        st.markdown("---")
        st.markdown("#### Regras cadastradas")

        nuc_df = pd.DataFrame(payload.get("nucleo_rules", []))
        crit_df = pd.DataFrame(payload.get("criticidade_rules", []))

        st.markdown("**Regras de Núcleo**")
        if nuc_df.empty:
            st.info("Nenhuma regra de núcleo cadastrada.")
        else:
            st.dataframe(nuc_df, use_container_width=True, height=220)
            rid = st.number_input("ID da regra de núcleo", min_value=1, step=1, key="rid_nuc")
            colx1, colx2, colx3 = st.columns(3)
            with colx1:
                if st.button("Ativar regra núcleo"):
                    update_rule_status("nucleo", rid, True)
                    st.rerun()
            with colx2:
                if st.button("Inativar regra núcleo"):
                    update_rule_status("nucleo", rid, False)
                    st.rerun()
            with colx3:
                if st.button("Excluir regra núcleo"):
                    delete_rule("nucleo", rid)
                    st.rerun()

        st.markdown("**Regras de Criticidade**")
        if crit_df.empty:
            st.info("Nenhuma regra de criticidade cadastrada.")
        else:
            st.dataframe(crit_df, use_container_width=True, height=220)
            rid2 = st.number_input("ID da regra de criticidade", min_value=1, step=1, key="rid_cri")
            coly1, coly2, coly3 = st.columns(3)
            with coly1:
                if st.button("Ativar regra criticidade"):
                    update_rule_status("criticidade", rid2, True)
                    st.rerun()
            with coly2:
                if st.button("Inativar regra criticidade"):
                    update_rule_status("criticidade", rid2, False)
                    st.rerun()
            with coly3:
                if st.button("Excluir regra criticidade"):
                    delete_rule("criticidade", rid2)
                    st.rerun()

    # =====================================================
    # Resumo / priorização / motivos detectados
    # =====================================================
    with st.expander("Resumo para priorização (abertos, top impacto, distribuições, motivos detectados)", expanded=True):
        df_open = div_master.loc[~resolved_mask].copy()
        df_open["ABS"] = df_open["VALOR"].abs()

        st.markdown("**Top 10 em aberto por impacto**")
        t1, t2, t3 = st.columns([1.1, 1.8, 2.1], gap="large")
        with t1:
            top_origem = st.selectbox("Origem (Top 10)", ["Todas", "Somente Financeiro", "Somente Contábil"], key="top10_origem")
        with t2:
            nuc_opts = ["Todos"] + sorted([x for x in df_open["NUCLEO"].fillna("Não identificado").unique().tolist() if str(x).strip() != ""])
            top_nucleo = st.selectbox("Núcleo (Top 10)", nuc_opts, key="top10_nucleo")
        with t3:
            st.caption("Estes filtros atuam apenas no Top 10 (não mexem na tratativa).")

        top_src = df_open.copy()
        if top_origem != "Todas":
            top_src = top_src[top_src["ORIGEM"] == top_origem].copy()
        if top_nucleo != "Todos":
            top_src = top_src[top_src["NUCLEO"] == top_nucleo].copy()

        top_open = top_src.sort_values("ABS", ascending=False).head(10)
        show_cols = ["ORIGEM", "DATA", "DOCUMENTO", "VALOR", "NUCLEO"]
        st.dataframe(top_open[show_cols].copy(), use_container_width=True, height=320)

        st.markdown("**Distribuição por Origem (abertos)**")
        if len(df_open):
            dist_origem = df_open.groupby("ORIGEM", dropna=False).agg(Itens=("VALOR","size"), Valor=("VALOR","sum")).reset_index().sort_values("Valor", ascending=False)
            st.dataframe(dist_origem, use_container_width=True, height=160)

            st.markdown("**Distribuição por Origem × Núcleo (abertos)**")
            dist_on = df_open.groupby(["ORIGEM","NUCLEO"], dropna=False).agg(Itens=("VALOR","size"), Valor=("VALOR","sum")).reset_index().sort_values(["ORIGEM","Valor"], ascending=[True, False])
            st.dataframe(dist_on, use_container_width=True, height=220)

            st.markdown("**Comparativo (abertos): Financeiro × Contábil**")
            comp = df_open.groupby("ORIGEM", dropna=False)["VALOR"].sum().reset_index()
            comp = comp[comp["ORIGEM"].isin(["Somente Financeiro", "Somente Contábil"])].copy()
            comp = comp.set_index("ORIGEM")
            st.bar_chart(comp["VALOR"])

            st.markdown("**Motivos detectados (agrupamento base)**")
            motivos = (
                df_open.groupby(["MOTIVO_BASE", "ORIGEM"], dropna=False)
                .agg(
                    Itens=("VALOR", "size"),
                    Impacto=("VALOR", "sum"),
                    Maior_Valor=("VALOR", lambda s: float(np.max(np.abs(s))) if len(s) else 0.0)
                )
                .reset_index()
            )
            motivos["ABS_IMPACTO"] = motivos["Impacto"].abs()
            motivos = motivos.sort_values(["Itens", "ABS_IMPACTO"], ascending=[False, False])
            st.dataframe(motivos[["MOTIVO_BASE", "ORIGEM", "Itens", "Impacto", "Maior_Valor"]].head(25), use_container_width=True, height=320)

            st.markdown("**Criar regra rápida a partir do motivo detectado**")
            r1, r2, r3 = st.columns([2.2, 1.2, 1.0])
            motivos_opts = [""] + motivos["MOTIVO_BASE"].fillna("").astype(str).unique().tolist()
            with r1:
                motivo_sel = st.selectbox("Motivo base", motivos_opts, key="motivo_sel_rapido")
            with r2:
                origem_sel = st.selectbox("Origem do motivo", ["", "Somente Financeiro", "Somente Contábil"], key="origem_sel_rapido")
            with r3:
                tipo_rapido = st.selectbox("Tipo", ["Núcleo", "Criticidade"], key="tipo_rapido")

            r4, r5, r6 = st.columns([1.2, 1.2, 1.0])
            with r4:
                res_nuc = st.selectbox("Resultado Núcleo", NUCLEOS, key="res_nuc_rapido")
            with r5:
                res_crit = st.selectbox("Resultado Criticidade", SEVERIDADES, key="res_crit_rapido")
            with r6:
                prio_rapido = st.number_input("Prioridade rápida", min_value=1, value=80, step=1, key="prio_rapido")

            if st.button("Salvar regra rápida", type="primary"):
                if not str(motivo_sel).strip():
                    st.warning("Selecione um motivo base.")
                else:
                    if tipo_rapido == "Núcleo":
                        add_rule("nucleo", {
                            "nome": f"Motivo: {motivo_sel[:60]}",
                            "prioridade": int(prio_rapido),
                            "ativa": True,
                            "origem": origem_sel.strip(),
                            "texto_contem": motivo_sel.strip(),
                            "regex": "",
                            "documento_prefixo": "",
                            "valor_min": "",
                            "valor_max": "",
                            "resultado": res_nuc,
                        })
                    else:
                        add_rule("criticidade", {
                            "nome": f"Motivo: {motivo_sel[:60]}",
                            "prioridade": int(prio_rapido),
                            "ativa": True,
                            "origem": origem_sel.strip(),
                            "texto_contem": motivo_sel.strip(),
                            "regex": "",
                            "documento_prefixo": "",
                            "valor_min": "",
                            "valor_max": "",
                            "resultado": res_crit,
                        })

                    dm = st.session_state.div_master.copy()
                    dm = apply_classification_rules(dm)
                    st.session_state.div_master = dm
                    st.success("Regra rápida salva e reaplicada.")
                    st.rerun()
        else:
            st.info("Sem pendências em aberto.")

    # =====================================================
    # Filtros
    # =====================================================
    st.markdown("### Filtros")
    f1, f2, f3, f4, f5 = st.columns([1.1, 1.05, 1.05, 2.3, 1.0], gap="large")
    with f1:
        origem = st.selectbox("Origem", ["Todas", "Somente Financeiro", "Somente Contábil"])
    with f2:
        ver = st.selectbox("Visualizar", ["Todas", "Somente em aberto", "Somente resolvidas"])
    with f3:
        sev = st.selectbox("Severidade", ["Todas", "Normal", "Atenção", "Crítica"])
    with f4:
        busca = st.text_input("Buscar (inclui valor)", value="")
    with f5:
        st.markdown("<div style='height:1px'></div>", unsafe_allow_html=True)

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
        cols_search = [
            "DOCUMENTO", "HISTORICO_OPERACAO", "CHAVE_DOC", "NUCLEO", "ORIGEM",
            "SEVERIDADE", "MOTIVO_BASE", "REGRA_NUCLEO_APLICADA", "REGRA_SEVERIDADE_APLICADA"
        ]
        mask = False
        for c in cols_search:
            if c in df.columns:
                mask = mask | df[c].astype(str).str.lower().str.contains(q, na=False)

        if "VALOR" in df.columns:
            mask = mask | df["VALOR"].map(lambda x: fmt(x)).astype(str).str.lower().str.contains(q, na=False)

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

    # =====================================================
    # Ações em massa
    # =====================================================
    st.markdown("### Ações em massa")
    st.markdown('<div class="cm-help">Fluxo: 1) Filtre  2) Selecione  3) Defina ação  4) Clique em Aplicar.</div>', unsafe_allow_html=True)

    ids_filtrados = list(df.index)
    dm0 = st.session_state.div_master.copy()
    selecionados_count = int(dm0["SELECIONADO"].fillna(False).sum())

    bar = st.container()
    with bar:
        st.markdown('<div class="cm-actionbar">', unsafe_allow_html=True)
        a1, a2, a3 = st.columns([1.2, 1.2, 2.2], gap="large")
        with a1:
            if st.button("Selecionar todos do filtro"):
                dm = st.session_state.div_master.copy()
                dm.loc[ids_filtrados, "SELECIONADO"] = True
                st.session_state.div_master = dm
                st.rerun()
        with a2:
            if st.button("Limpar seleção do filtro"):
                dm = st.session_state.div_master.copy()
                dm.loc[ids_filtrados, "SELECIONADO"] = False
                st.session_state.div_master = dm
                st.rerun()
        with a3:
            st.markdown(f'<span class="cm-badge">Selecionados: {selecionados_count}</span>', unsafe_allow_html=True)
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
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
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
        st.session_state.div_master = dm
        st.success(f"Ação aplicada em {len(target_ids)} itens.")
        st.rerun()

    # =====================================================
    # Tratativa
    # =====================================================
    st.markdown("### Tratativa (tabela)")
    st.markdown('<div class="cm-help">Sugestão: confirme quando fizer sentido; status e obs ajudam na rastreabilidade. Resolver marca Status=Resolvido.</div>', unsafe_allow_html=True)

    view_cols = [
        "SELECIONADO",
        "ORIGEM", "SEVERIDADE", "DATA", "DOCUMENTO", "HISTORICO_OPERACAO", "CHAVE_DOC", "VALOR",
        "MOTIVO_BASE",
        "NUCLEO_SUGERIDO", "REGRA_NUCLEO_APLICADA", "REGRA_SEVERIDADE_APLICADA",
        "CONFIRMADO", "NUCLEO",
        "STATUS", "RESOLVIDO", "OBS_USUARIO"
    ]
    df_view = df[view_cols].copy()
    df_view_display = df_view.copy()
    df_view_display["DATA"] = pd.to_datetime(df_view_display["DATA"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")

    column_config = {
        "SELECIONADO": st.column_config.CheckboxColumn(),
        "ORIGEM": st.column_config.TextColumn(disabled=True),
        "SEVERIDADE": st.column_config.TextColumn(disabled=True),
        "DATA": st.column_config.TextColumn(disabled=True),
        "DOCUMENTO": st.column_config.TextColumn(disabled=True),
        "HISTORICO_OPERACAO": st.column_config.TextColumn(disabled=True),
        "CHAVE_DOC": st.column_config.TextColumn(disabled=True),
        "VALOR": st.column_config.NumberColumn(format="R$ %.2f", disabled=True),
        "MOTIVO_BASE": st.column_config.TextColumn(disabled=True),
        "NUCLEO_SUGERIDO": st.column_config.TextColumn(disabled=True),
        "REGRA_NUCLEO_APLICADA": st.column_config.TextColumn(disabled=True),
        "REGRA_SEVERIDADE_APLICADA": st.column_config.TextColumn(disabled=True),
        "CONFIRMADO": st.column_config.CheckboxColumn(),
        "NUCLEO": st.column_config.SelectboxColumn(options=NUCLEOS),
        "STATUS": st.column_config.SelectboxColumn(options=STATUS_OPTS),
        "RESOLVIDO": st.column_config.CheckboxColumn(),
        "OBS_USUARIO": st.column_config.TextColumn(),
    }

    edited = st.data_editor(
        df_view_display,
        use_container_width=True,
        height=560,
        column_config=column_config,
        key="editor_tratativa",
        hide_index=False,
    )

    if edited is not None and len(edited) == len(df_view_display):
        to_update = edited.copy()

        if "NUCLEO_SUGERIDO" in to_update.columns:
            to_update["NUCLEO"] = to_update["NUCLEO"].fillna("Não identificado").replace("", "Não identificado")
            need = to_update["CONFIRMADO"].fillna(False) & (to_update["NUCLEO"].astype(str).str.strip().eq("") | to_update["NUCLEO"].eq("Não identificado"))
            to_update.loc[need, "NUCLEO"] = to_update.loc[need, "NUCLEO_SUGERIDO"].fillna("Não identificado")

        res_col = to_update["RESOLVIDO"].fillna(False)
        to_update.loc[res_col, "STATUS"] = "Resolvido"

        upd_cols = ["SELECIONADO", "CONFIRMADO", "NUCLEO", "STATUS", "RESOLVIDO", "OBS_USUARIO"]
        dm = st.session_state.div_master.copy()
        for c in upd_cols:
            dm.loc[to_update.index, c] = to_update[c].values

        st.session_state.div_master = dm
        div_master = dm.copy()

    # =====================================================
    # Detalhe do item
    # =====================================================
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
        resolvido_txt = "Sim" if bool(r.get("RESOLVIDO", False)) else "Não"

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
  <div class="row"><span class="label">Motivo base:</span> <span class="val">{r.get('MOTIVO_BASE','')}</span></div>
  <div class="row"><span class="label">Núcleo sugerido:</span> <span class="val">{r.get('NUCLEO_SUGERIDO','')}</span></div>
  <div class="row"><span class="label">Regra núcleo:</span> <span class="val">{r.get('REGRA_NUCLEO_APLICADA','')}</span></div>
  <div class="row"><span class="label">Regra criticidade:</span> <span class="val">{r.get('REGRA_SEVERIDADE_APLICADA','')}</span></div>
  <div class="row"><span class="label">Confirmado:</span> <span class="val">{confirmado_txt}</span></div>
  <div class="row"><span class="label">Núcleo:</span> <span class="val">{r.get('NUCLEO','')}</span></div>
  <div class="row"><span class="label">Status:</span> <span class="val">{r.get('STATUS','')}</span></div>
  <div class="row"><span class="label">Resolvido:</span> <span class="val">{resolvido_txt}</span></div>
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
            f"MOTIVO_BASE: {r.get('MOTIVO_BASE','')}\n"
            f"NUCLEO_SUGERIDO: {r.get('NUCLEO_SUGERIDO','')}\n"
            f"REGRA_NUCLEO_APLICADA: {r.get('REGRA_NUCLEO_APLICADA','')}\n"
            f"REGRA_SEVERIDADE_APLICADA: {r.get('REGRA_SEVERIDADE_APLICADA','')}\n"
            f"CONFIRMADO: {confirmado_txt}\n"
            f"NUCLEO: {r.get('NUCLEO','')}\n"
            f"STATUS: {r.get('STATUS','')}\n"
            f"RESOLVIDO: {resolvido_txt}\n"
            f"OBS: {r.get('OBS_USUARIO','')}\n"
            f"HISTÓRICO: {r.get('HISTORICO_OPERACAO','')}"
        )
        st.text_area("Copiar resumo (e-mail/ticket)", value=resumo, height=230)

    # =====================================================
    # Export
    # =====================================================
    st.markdown("### Exportar")
    filtros = {"origem": origem, "ver": ver, "severidade": sev, "busca": busca.strip(), "_total_aberto": valor_aberto}

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
