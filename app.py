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
# CONFIG
# =========================================================
st.set_page_config(page_title="ConciliaMais — Conferência de Extrato Bancário", layout="wide", initial_sidebar_state="expanded")


# =========================================================
# THEME + UX CSS (Dark, clean, blue brand)
# =========================================================
st.markdown(
    """
<style>
:root{
  --bg: #0B1220;
  --card: #111827;
  --card2: #0F172A;
  --border: rgba(148,163,184,.18);
  --text: #E5E7EB;
  --muted: #94A3B8;
  --primary: #2563EB;     /* azul ConciliaMais */
  --primary2: #1D4ED8;
  --ok: #22C55E;
  --warn: #F59E0B;
  --bad: #EF4444;
  --shadow: 0 14px 28px rgba(0,0,0,.35);
}

html, body, [class*="css"]  { color: var(--text) !important; }
body { background: var(--bg) !important; }

.block-container {
  padding-top: 1.0rem;
  padding-bottom: 2.0rem;
  max-width: 1500px;
}

h1, h2, h3 { letter-spacing: -0.02em; }
small, .stCaption, .stMarkdown p { color: var(--muted); }

/* Cards / shells */
.cm-shell{
  background: linear-gradient(180deg, rgba(17,24,39,.92), rgba(15,23,42,.92));
  border: 1px solid var(--border);
  border-radius: 18px;
  padding: 14px;
  box-shadow: var(--shadow);
}
.cm-section{ margin-top: 16px; }

.cm-help { color: var(--muted); font-size: 13px; margin-top: -6px; }

.cm-cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-top: 10px; }
.cm-card{
  border-radius: 16px;
  padding: 14px 14px 12px 14px;
  background: rgba(17,24,39,.75);
  border: 1px solid var(--border);
  box-shadow: 0 10px 22px rgba(0,0,0,.25);
}
.cm-card .k{ font-size: 12px; color: var(--muted); margin-bottom: 6px; }
.cm-card .v{ font-size: 22px; font-weight: 900; color: var(--text); letter-spacing:-.01em; }
.cm-card .s{ font-size: 12px; color: var(--muted); margin-top: 6px; }

.cm-mini{
  border-radius: 14px;
  padding: 10px 12px;
  background: rgba(17,24,39,.65);
  border: 1px solid var(--border);
  text-align: right;
  box-shadow: 0 10px 22px rgba(0,0,0,.22);
}
.cm-mini .k{ font-size: 12px; color: var(--muted); margin-bottom: 4px; }
.cm-mini .v{ font-size: 20px; font-weight: 900; letter-spacing: -0.01em; color: var(--text); }

.cm-alert{
  border-radius: 16px;
  padding: 12px 14px;
  border: 1px solid var(--border);
  background: rgba(17,24,39,.72);
  box-shadow: 0 10px 22px rgba(0,0,0,.22);
  margin-top: 10px;
}
.cm-alert .t{ font-weight: 900; margin-bottom: 4px; }
.cm-alert .d{ color: var(--muted); font-size: 13px; }

.cm-alert-info{ border-left: 5px solid rgba(37,99,235,.85); }
.cm-alert-warn{ border-left: 5px solid rgba(245,158,11,.90); }
.cm-alert-ok{ border-left: 5px solid rgba(34,197,94,.90); }

.cm-tag{
  display:inline-flex; align-items:center; gap:8px;
  padding: 5px 10px;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 900;
  border: 1px solid var(--border);
  background: rgba(2,6,23,.35);
}
.cm-dot{ width:9px; height:9px; border-radius:99px; display:inline-block; }
.cm-dot-fin{ background: #3B82F6; }
.cm-dot-led{ background: #A78BFA; }
.cm-dot-ok{ background: var(--ok); }
.cm-dot-warn{ background: var(--warn); }
.cm-dot-bad{ background: var(--bad); }

.cm-breadcrumb{
  display:flex; align-items:center; gap:10px;
  color: var(--muted);
  font-size: 13px;
  margin-top: -4px;
  margin-bottom: 10px;
}
.cm-breadcrumb b{ color: var(--text); }

.cm-row{
  background: rgba(17,24,39,.62);
  border: 1px solid var(--border);
  border-radius: 16px;
  padding: 12px;
  box-shadow: 0 10px 22px rgba(0,0,0,.22);
}

.cm-detail{
  border-radius: 18px;
  padding: 14px;
  background: rgba(17,24,39,.74);
  border: 1px solid var(--border);
  box-shadow: 0 10px 22px rgba(0,0,0,.22);
}
.cm-detail .title{ font-size: 14px; font-weight: 900; margin-bottom: 10px; }
.cm-detail .row{ font-size: 13px; margin: 5px 0; }
.cm-detail .label{ color: var(--muted); }
.cm-detail .val{ font-weight: 800; color: var(--text); }

/* Primary buttons in blue */
.stButton button[kind="primary"]{
  background: var(--primary) !important;
  border: 1px solid rgba(59,130,246,.65) !important;
  color: white !important;
  border-radius: 12px !important;
}
.stButton button[kind="primary"]:hover{
  background: var(--primary2) !important;
}

/* Make secondary buttons consistent */
.stButton button{
  border-radius: 12px !important;
}

/* Sidebar toggle via CSS */
.cm-hide-sidebar section[data-testid="stSidebar"]{
  display: none !important;
}
.cm-hide-sidebar div[data-testid="stSidebarNav"]{
  display: none !important;
}

/* Data editor/table look */
div[data-testid="stDataFrame"]{
  border-radius: 14px !important;
  border: 1px solid var(--border) !important;
  overflow: hidden !important;
}
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# HELPERS (MOTOR)
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
# CLASSIFICAÇÃO (NÚCLEO SUGERIDO)
# =========================================================
NUCLEOS = ["Processo interno", "Cadastro", "Configuração RP", "Não identificado"]
STATUS_OPTS = ["Pendente", "Em análise", "Resolvido"]


def suggest_nucleo(row):
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


# =========================================================
# RECONCILIAÇÃO
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
        fin_rows.append(
            {
                "ORIGEM": "Somente Financeiro",
                "DATA": r["__date"],
                "DOCUMENTO": str(base.get(cfg.get("fin_documento"), "")) if cfg.get("fin_documento") else "",
                "PREFIXO_TITULO": str(base.get(cfg.get("fin_prefixo"), "")) if cfg.get("fin_prefixo") else "",
                "HISTORICO_OPERACAO": str(base.get(cfg.get("fin_operacao"), "")) if cfg.get("fin_operacao") else str(r["__text"]),
                "CHAVE_DOC": r["__doc_key"],
                "VALOR": round(float(r["__amount"]), 2) if pd.notna(r["__amount"]) else np.nan,
            }
        )

    led_rows = []
    led_reset = led_df.reset_index(drop=True)
    for _, r in led_only.iterrows():
        i = int(r["__idx"])
        base = led_reset.iloc[i] if 0 <= i < len(led_reset) else pd.Series(dtype="object")
        hist_val = str(base.get(cfg.get("led_historico"), "")) if cfg.get("led_historico") else str(r["__text"])
        doc_from_hist = extract_doc_from_ledger_history(hist_val)
        led_rows.append(
            {
                "ORIGEM": "Somente Contábil",
                "DATA": r["__date"],
                "DOCUMENTO": doc_from_hist,
                "PREFIXO_TITULO": "",
                "HISTORICO_OPERACAO": hist_val,
                "CHAVE_DOC": r["__doc_key"],
                "VALOR": round(float(r["__amount"]), 2) if pd.notna(r["__amount"]) else np.nan,
            }
        )

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
# UX HELPERS
# =========================================================
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


def origem_tag_html(origem: str) -> str:
    if str(origem) == "Somente Financeiro":
        return '<span class="cm-tag"><span class="cm-dot cm-dot-fin"></span>Somente Financeiro</span>'
    if str(origem) == "Somente Contábil":
        return '<span class="cm-tag"><span class="cm-dot cm-dot-led"></span>Somente Contábil</span>'
    return '<span class="cm-tag"><span class="cm-dot"></span>Todas</span>'


def severidade_tag_html(label: str) -> str:
    lab = str(label)
    if lab == "Normal":
        return '<span class="cm-tag"><span class="cm-dot cm-dot-ok"></span>Normal</span>'
    if lab == "Atenção":
        return '<span class="cm-tag"><span class="cm-dot cm-dot-warn"></span>Atenção</span>'
    return '<span class="cm-tag"><span class="cm-dot cm-dot-bad"></span>Crítica</span>'


def confirmado_txt(x) -> str:
    return "Sim" if bool(x) else "Não"


def style_table_tags(df: pd.DataFrame) -> "pd.io.formats.style.Styler":
    """Styler for read-only tables (top10/distribuições)."""
    dfx = df.copy()

    def bg_origem(val):
        if val == "Somente Financeiro":
            return "background-color: rgba(59,130,246,.15); color: #E5E7EB; font-weight: 800;"
        if val == "Somente Contábil":
            return "background-color: rgba(167,139,250,.16); color: #E5E7EB; font-weight: 800;"
        return ""

    def bg_sev(val):
        if val == "Normal":
            return "background-color: rgba(34,197,94,.14); color:#E5E7EB; font-weight:800;"
        if val == "Atenção":
            return "background-color: rgba(245,158,11,.14); color:#E5E7EB; font-weight:800;"
        if val == "Crítica":
            return "background-color: rgba(239,68,68,.12); color:#E5E7EB; font-weight:800;"
        return ""

    sty = dfx.style
    if "ORIGEM" in dfx.columns:
        sty = sty.applymap(bg_origem, subset=["ORIGEM"])
    if "SEVERIDADE" in dfx.columns:
        sty = sty.applymap(bg_sev, subset=["SEVERIDADE"])

    sty = sty.format({"VALOR": lambda v: f"R$ {fmt(v)}" if pd.notna(v) else ""})
    sty = sty.set_table_styles(
        [
            {"selector": "th", "props": [("background-color", "rgba(37,99,235,.20)"), ("color", "#E5E7EB"), ("font-weight", "900")]},
            {"selector": "td", "props": [("border-color", "rgba(148,163,184,.18)")]},
            {"selector": "table", "props": [("border-collapse", "collapse"), ("border-radius", "14px"), ("overflow", "hidden")]},
        ]
    )
    return sty


# =========================================================
# EXCEL EXPORT (Tabela + Cabeçalho azul + Total dinâmico por filtro)
# =========================================================
def _autofit_worksheet(ws, df, start_col, max_width=75, min_width=10):
    for j, col in enumerate(df.columns):
        ser = df[col].astype(str).fillna("")
        sample = ser.head(300).tolist()
        max_len = max([len(str(col))] + [len(s) for s in sample]) if sample else len(str(col))
        width = max(min_width, min(max_width, max_len + 2))
        ws.set_column(start_col + j, start_col + j, width)


def to_excel_divergencias_filtradas(df_filtrado, total_aberto, filtros, stats, generated_at):
    """
    - Sai como TABELA do Excel (filtro automático)
    - Cabeçalho azul claro
    - Linha de TOTAL com SUBTOTAL (muda com filtro)
    """
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        wb = w.book

        # Formats
        fmt_title = wb.add_format({"bold": True, "font_size": 14, "font_color": "#0F172A"})
        fmt_k = wb.add_format({"bold": True, "font_size": 10, "font_color": "#334155"})
        fmt_hdr = wb.add_format({"bold": True, "border": 1, "align": "center", "valign": "vcenter", "bg_color": "#DBEAFE", "font_color": "#0F172A"})
        fmt_txt = wb.add_format({"border": 1})
        fmt_date = wb.add_format({"num_format": "dd/mm/yyyy", "border": 1})
        fmt_money = wb.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00', "border": 1})
        fmt_money_big = wb.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00', "bold": True, "font_color": "#0F172A"})
        fmt_subhdr = wb.add_format({"bold": True, "font_size": 11, "font_color": "#0F172A"})

        sh = "Divergencias"
        df = df_filtrado.copy()

        if "DATA" in df.columns:
            df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
        if "VALOR" in df.columns:
            df["VALOR"] = df["VALOR"].map(normalize_money)

        # Ensure ID column
        df2 = df.copy()
        df2.insert(0, "ID", df2.index.astype(int))

        df2.to_excel(w, index=False, sheet_name=sh, startrow=8)
        ws = w.sheets[sh]

        # Header/top info
        ws.write(0, 0, "ConciliaMais — Divergências (Excel igual ao filtro)", fmt_title)
        ws.write(1, 0, "Processado em:", fmt_k)
        ws.write(1, 1, generated_at)

        ws.write(2, 0, "Origem:", fmt_k)
        ws.write(2, 1, filtros.get("origem", "Todas"))
        ws.write(3, 0, "Visualização:", fmt_k)
        ws.write(3, 1, filtros.get("ver", "Todas"))
        ws.write(4, 0, "Severidade:", fmt_k)
        ws.write(4, 1, filtros.get("severidade", "Todas"))
        ws.write(5, 0, "Busca:", fmt_k)
        ws.write(5, 1, filtros.get("busca", ""))

        ws.write(2, 6, "Total em aberto:", fmt_k)
        ws.write_number(2, 7, float(total_aberto or 0.0), fmt_money_big)

        # Format header row
        ws.set_row(8, 22, fmt_hdr)
        ws.freeze_panes(9, 0)

        nrows = len(df2)
        ncols = len(df2.columns)

        # Excel table (with total row dynamic)
        if nrows > 0 and ncols > 0:
            first_row = 8
            first_col = 0
            last_row = first_row + nrows
            last_col = first_col + ncols - 1

            columns = []
            for c in df2.columns:
                col_def = {"header": c}
                if c == "VALOR":
                    col_def["total_function"] = "sum"  # Excel total row uses SUBTOTAL (respects filter)
                if c == "ID":
                    col_def["total_string"] = "TOTAL (filtrado)"
                columns.append(col_def)

            ws.add_table(
                first_row,
                first_col,
                last_row,
                last_col,
                {
                    "style": "Table Style Medium 9",
                    "columns": columns,
                    "autofilter": True,
                    "total_row": True,
                },
            )

            # Apply per-column formats (data types)
            col_map = {name: idx for idx, name in enumerate(df2.columns)}
            # Re-write data cells with formats (keep table)
            for r in range(nrows):
                excel_r = 9 + r
                for c_name, c_idx in col_map.items():
                    val = df2.iloc[r, c_idx]
                    if c_name == "DATA":
                        if pd.notna(val):
                            ws.write_datetime(excel_r, c_idx, val.to_pydatetime(), fmt_date)
                        else:
                            ws.write_blank(excel_r, c_idx, None, fmt_date)
                    elif c_name == "VALOR":
                        if pd.notna(val):
                            ws.write_number(excel_r, c_idx, float(val), fmt_money)
                        else:
                            ws.write_blank(excel_r, c_idx, None, fmt_money)
                    elif c_name == "ID":
                        ws.write_number(excel_r, c_idx, int(val), fmt_txt)
                    else:
                        ws.write(excel_r, c_idx, "" if pd.isna(val) else str(val), fmt_txt)

        _autofit_worksheet(ws, df2, start_col=0)

        # Resumo sheet
        resumo = pd.DataFrame(
            [
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
            ],
            columns=["Métrica", "Valor"],
        )
        resumo.to_excel(w, index=False, sheet_name="Resumo")
        ws3 = w.sheets["Resumo"]
        ws3.freeze_panes(1, 0)
        ws3.set_row(0, 22, wb.add_format({"bold": True, "bg_color": "#DBEAFE", "font_color": "#0F172A", "border": 1}))
        ws3.set_column(0, 0, 52)
        ws3.set_column(1, 1, 26, wb.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00'}))
        ws3.write(0, 0, "Métrica", fmt_subhdr)
        ws3.write(0, 1, "Valor", fmt_subhdr)

    out.seek(0)
    return out


# =========================================================
# PDF RESUMO (mantém lógica, só consome dados atuais)
# =========================================================
def to_pdf_resumo(stats, generated_at, div_master):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("Relatório de Conciliação Bancária — ConciliaMais (Módulo Financeiro)", styles["Title"]))
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
    t1.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
            ]
        )
    )
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
    t2.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
            ]
        )
    )
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
        top_rows.append(
            [
                str(i),
                str(r.get("ORIGEM", "")),
                dt,
                str(r.get("DOCUMENTO", "")),
                fmt(float(r.get("VALOR", 0.0))),
                str(r.get("NUCLEO", "Não identificado") or "Não identificado"),
            ]
        )

    t_top = Table(top_rows, colWidths=[22, 85, 58, 175, 75, 80])
    t_top.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    story.append(t_top)
    story.append(Spacer(1, 14))

    doc.build(story)
    buf.seek(0)
    return buf


# =========================================================
# STATE
# =========================================================
if "page" not in st.session_state:
    st.session_state.page = "upload"  # upload | resultados

if "results" not in st.session_state:
    st.session_state.results = None

if "div_master" not in st.session_state:
    st.session_state.div_master = None

if "upload_step" not in st.session_state:
    st.session_state.upload_step = 1

if "nav_module" not in st.session_state:
    st.session_state.nav_module = "Financeiro"

if "nav_area" not in st.session_state:
    st.session_state.nav_area = "Extrato Bancário"

if "show_sidebar" not in st.session_state:
    st.session_state.show_sidebar = True


# =========================================================
# NAV (Sidebar + Collapse)
# =========================================================
if not st.session_state.show_sidebar:
    st.markdown('<div class="cm-hide-sidebar">', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## Navegação")
    st.caption("Estrutura preparada para novos módulos.")
    st.session_state.nav_module = st.radio("Módulo", ["Financeiro", "ProCV"], index=0 if st.session_state.nav_module == "Financeiro" else 1)
    if st.session_state.nav_module == "Financeiro":
        st.session_state.nav_area = st.radio("Área", ["Extrato Bancário", "Posição a Pagar", "Posição a Receber"], index=0)
    else:
        st.session_state.nav_area = st.radio("Área", ["(em breve)"], index=0)
    st.markdown("---")
    st.caption("Você está em:")
    st.write(f"**{st.session_state.nav_module}**  ›  **{st.session_state.nav_area}**")

if not st.session_state.show_sidebar:
    st.markdown("</div>", unsafe_allow_html=True)

# Top bar (toggle + back)
topL, topR = st.columns([1.2, 1.0])
with topL:
    st.markdown(
        f"""
<div class="cm-breadcrumb">
  <span>Você está em:</span>
  <b>{st.session_state.nav_module}</b> <span>›</span> <b>{st.session_state.nav_area}</b>
</div>
""",
        unsafe_allow_html=True,
    )

with topR:
    cA, cB = st.columns([1, 1])
    with cA:
        if st.button(("Ocultar menu" if st.session_state.show_sidebar else "Mostrar menu"), use_container_width=True):
            st.session_state.show_sidebar = not st.session_state.show_sidebar
            st.rerun()
    with cB:
        if st.session_state.page == "resultados":
            if st.button("← Voltar (Upload)", use_container_width=True):
                st.session_state.page = "upload"
                st.session_state.upload_step = 1
                st.rerun()


# =========================================================
# GATE: modules not implemented
# =========================================================
if st.session_state.nav_module != "Financeiro" or st.session_state.nav_area != "Extrato Bancário":
    st.title("Em breve")
    st.caption("Esta área está preparada na navegação, mas será implementada na sequência.")
    st.stop()


# =========================================================
# PAGE: UPLOAD (Wizard)
# =========================================================
if st.session_state.page == "upload":
    st.title("ConciliaMais — Conferência de Extrato Bancário")
    st.caption("Extrato Financeiro + Razão Contábil → Match automático → Divergências → Tratativa")

    # Wizard header
    with st.container():
        st.markdown('<div class="cm-shell">', unsafe_allow_html=True)
        st.markdown("### Etapas")
        st.progress(min(max((st.session_state.upload_step - 1) / 3, 0.0), 1.0))
        st.markdown(
            "<div class='cm-help'>1) Upload  |  2) Mapeamento  |  3) Validação de saldos  |  4) Processar</div>",
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="cm-section"></div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1) Extrato Financeiro")
        st.markdown('<div class="cm-help">Faça o upload da planilha do Extrato Financeiro.</div>', unsafe_allow_html=True)
        fin_file = st.file_uploader("Upload do Extrato Financeiro (.xlsx ou .csv)", type=["xlsx", "csv"], key="fin")
    with c2:
        st.subheader("1) Razão Contábil")
        st.markdown('<div class="cm-help">Faça o upload da planilha do Razão Contábil.</div>', unsafe_allow_html=True)
        led_file = st.file_uploader("Upload do Razão Contábil (.xlsx ou .csv)", type=["xlsx", "csv"], key="led")

    if not fin_file or not led_file:
        st.session_state.upload_step = 1
        st.markdown(
            """
<div class="cm-alert cm-alert-info">
  <div class="t">Para continuar</div>
  <div class="d">Faça o upload dos dois arquivos para liberar as próximas etapas.</div>
</div>
""",
            unsafe_allow_html=True,
        )
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
        fin_date = st.selectbox(
            "Data",
            fin_df.columns,
            index=fin_df.columns.get_loc(fin_guess["date"]) if fin_guess["date"] in fin_df.columns else 0,
        )

        fin_operacao = st.selectbox(
            "Operação/Histórico",
            ["(nenhuma)"] + list(fin_df.columns),
            index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["operacao"]) if fin_guess["operacao"] in fin_df.columns else 0,
        )

        fin_documento = st.selectbox(
            "Documento",
            ["(nenhuma)"] + list(fin_df.columns),
            index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["documento"]) if fin_guess["documento"] in fin_df.columns else 0,
        )

        fin_prefixo = st.selectbox(
            "Prefixo/Título (usaremos como DOCUMENTO na divergência)",
            ["(nenhuma)"] + list(fin_df.columns),
            index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["prefixo"]) if fin_guess["prefixo"] in fin_df.columns else 0,
        )

        fin_entradas = st.selectbox(
            "Entradas",
            ["(nenhuma)"] + list(fin_df.columns),
            index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["entradas"]) if fin_guess["entradas"] in fin_df.columns else 0,
        )

        fin_saidas = st.selectbox(
            "Saídas",
            ["(nenhuma)"] + list(fin_df.columns),
            index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["saidas"]) if fin_guess["saidas"] in fin_df.columns else 0,
        )

        fin_amount = st.selectbox(
            "OU Valor Único",
            ["(usar Entradas - Saídas)"] + list(fin_df.columns),
            index=(["(usar Entradas - Saídas)"] + list(fin_df.columns)).index(fin_guess["valor"]) if fin_guess["valor"] in fin_df.columns else 0,
        )

        fin_saldo = st.selectbox(
            "Saldo",
            ["(nenhuma)"] + list(fin_df.columns),
            index=(["(nenhuma)"] + list(fin_df.columns)).index(fin_guess["saldo"]) if fin_guess["saldo"] in fin_df.columns else 0,
        )

    with b:
        st.markdown("#### Razão Contábil")
        led_date = st.selectbox(
            "Data",
            led_df.columns,
            index=led_df.columns.get_loc(led_guess["date"]) if led_guess["date"] in led_df.columns else 0,
            key="ld",
        )

        led_historico = st.selectbox(
            "Histórico",
            ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["historico"]) if led_guess["historico"] in led_df.columns else 0,
            key="lh",
        )

        led_doc = st.selectbox(
            "Documento/Lote (opcional)",
            ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["doc"]) if led_guess["doc"] in led_df.columns else 0,
            key="ldoc",
        )

        led_conta = st.selectbox(
            "Conta (opcional)",
            ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["conta"]) if led_guess["conta"] in led_df.columns else 0,
            key="lcta",
        )

        led_debito = st.selectbox(
            "Débito",
            ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["debito"]) if led_guess["debito"] in led_df.columns else 0,
            key="ldb",
        )

        led_credito = st.selectbox(
            "Crédito",
            ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["credito"]) if led_guess["credito"] in led_df.columns else 0,
            key="lcr",
        )

        led_amount = st.selectbox(
            "OU Valor Único",
            ["(usar Débito - Crédito)"] + list(led_df.columns),
            index=(["(usar Débito - Crédito)"] + list(led_df.columns)).index(led_guess["valor"]) if led_guess["valor"] in led_df.columns else 0,
            key="lamt",
        )

        led_saldo = st.selectbox(
            "Saldo",
            ["(nenhuma)"] + list(led_df.columns),
            index=(["(nenhuma)"] + list(led_df.columns)).index(led_guess["saldo"]) if led_guess["saldo"] in led_df.columns else 0,
            key="ls",
        )

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
        st.markdown(
            """
<div class="cm-alert cm-alert-info">
  <div class="t">Validação automática não disponível</div>
  <div class="d">Não foi possível calcular o saldo anterior automaticamente. Se existir saldo nos dois arquivos, selecione corretamente a coluna de saldo.</div>
</div>
""",
            unsafe_allow_html=True,
        )
    else:
        if abs(diff_ant) > 0.01:
            st.markdown(
                f"""
<div class="cm-alert cm-alert-warn">
  <div class="t">Atenção: saldo anterior não bate</div>
  <div class="d">Diferença (Financeiro - Contábil): <b>R$ {fmt(diff_ant)}</b>. Isso pode indicar divergências de períodos anteriores.</div>
</div>
""",
                unsafe_allow_html=True,
            )
            proceed_ok = st.checkbox("Prosseguir mesmo assim", value=False)
        else:
            st.markdown(
                """
<div class="cm-alert cm-alert-ok">
  <div class="t">OK: saldo anterior consistente</div>
  <div class="d">Os saldos anteriores estão aderentes entre Financeiro e Contábil.</div>
</div>
""",
                unsafe_allow_html=True,
            )

    date_tol = st.number_input("Tolerância de dias para match por data (0 = mesma data)", min_value=0, max_value=10, value=0, step=1)

    st.session_state.upload_step = 4

    st.markdown("<div class='cm-help'>Ao processar, o sistema vai gerar a lista de divergências e habilitar a tratativa (confirmado, status e observação).</div>", unsafe_allow_html=True)

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

        # Documento preferencial do financeiro: prefixo/título
        mask_fin = div["ORIGEM"].eq("Somente Financeiro")
        if "PREFIXO_TITULO" in div.columns:
            div.loc[mask_fin, "DOCUMENTO"] = div.loc[mask_fin, "PREFIXO_TITULO"].where(
                div.loc[mask_fin, "PREFIXO_TITULO"].astype(str).str.len() > 0,
                div.loc[mask_fin, "DOCUMENTO"],
            )

 # Documento do contábil extraído do histórico quando faltar
mask_led = div["ORIGEM"].eq("Somente Contábil")
if "HISTORICO_OPERACAO" in div.columns:
    missing = div.loc[mask_led, "DOCUMENTO"].astype(str).str.strip().eq("")
    div.loc[mask_led & missing, "DOCUMENTO"] = (
        div.loc[mask_led & missing, "HISTORICO_OPERACAO"]
        .astype(str)
        .str.extract(r'([A-Z]{2,}-?\d+[-/]?\d*)', expand=False)
        .fillna(div.loc[mask_led & missing, "DOCUMENTO"])
    )

