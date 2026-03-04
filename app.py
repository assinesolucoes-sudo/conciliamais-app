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
# CSS (visual mais “produto”)
# ----------------------------
st.markdown(
    """
<style>
/* Layout geral */
.block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
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

/* Seções */
.cm-section { margin-top: 18px; }

/* Helper */
.cm-help { opacity: .78; font-size: 13px; margin-top: -6px; }

/* Mini card (Total do filtro) */
.cm-mini {
  border-radius: 14px;
  padding: 10px 12px;
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.10);
  text-align: right;
}
.cm-mini .k { font-size: 12px; opacity: .80; margin-bottom: 4px; }
.cm-mini .v { font-size: 20px; font-weight: 800; letter-spacing: -0.01em; }
</style>
""",
    unsafe_allow_html=True,
)

# ----------------------------
# Helpers (MOTOR ORIGINAL)
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

def status_pill(conferencia):
    if conferencia is None or (isinstance(conferencia, float) and np.isnan(conferencia)):
        return '<span class="cm-pill cm-warn">Conferência não calculada</span>'
    x = abs(float(conferencia))
    if x <= 0.01:
        return '<span class="cm-pill cm-ok">Fechou (0,00)</span>'
    if x <= 5:
        return '<span class="cm-pill cm-warn">Quase (verificar)</span>'
    return '<span class="cm-pill cm-bad">Não fechou</span>'

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
        "fin_conc": int(len(fin_match)),   # pares conciliados (pelo financeiro)
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

def build_tratativa(div):
    if div is None or div.empty:
        return pd.DataFrame(columns=[
            "ORIGEM","DATA","IDENTIFICADOR","HISTORICO_OPERACAO","VALOR",
            "ACAO_SUGERIDA","RESPONSAVEL","PRAZO","STATUS","OBS"
        ])
    def ident(row):
        parts = []
        doc = str(row.get("DOCUMENTO","")).strip()
        if doc: parts.append(f"DOC:{doc}")
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

# ----------------------------
# Export Excel (filtrado, formatado)
# ----------------------------
def _autofit_worksheet(ws, df, header_row, start_col, max_width=60, min_width=10):
    # Estima largura por tamanho do texto (simples e eficiente)
    for j, col in enumerate(df.columns):
        ser = df[col].astype(str).fillna("")
        sample = ser.head(200).tolist()
        max_len = max([len(str(col))] + [len(s) for s in sample]) if sample else len(str(col))
        width = max(min_width, min(max_width, max_len + 2))
        ws.set_column(start_col + j, start_col + j, width)

def to_excel_divergencias_filtradas(df_filtrado, total_filtrado, stats, generated_at):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        wb = w.book

        fmt_title = wb.add_format({"bold": True, "font_size": 14})
        fmt_k = wb.add_format({"bold": True, "font_size": 10, "font_color": "#666666"})
        fmt_money = wb.add_format({"num_format": "#,##0.00", "border": 1})
        fmt_hdr = wb.add_format({"bold": True, "border": 1, "align": "center", "valign": "vcenter"})
        fmt_txt = wb.add_format({"border": 1})
        fmt_date = wb.add_format({"num_format": "yyyy-mm-dd", "border": 1})

        # Aba Divergencias
        sh = "Divergencias"
        df = df_filtrado.copy()

        # Garantir tipos
        if "DATA" in df.columns:
            # manter como texto/ date; xlsxwriter lida melhor com datetime
            df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
        if "VALOR" in df.columns:
            df["VALOR"] = df["VALOR"].map(normalize_money)

        df.to_excel(w, index=False, sheet_name=sh, startrow=4)
        ws = w.sheets[sh]

        # Cabeçalho de contexto
        ws.write(0, 0, "ConciliaMais — Divergências (Filtrado)", fmt_title)
        ws.write(1, 0, "Processado em:", fmt_k)
        ws.write(1, 1, generated_at)
        ws.write(2, 0, "Total do filtro:", fmt_k)
        ws.write_number(2, 1, float(total_filtrado or 0.0), fmt_money)

        # Formatar cabeçalho da tabela
        ws.set_row(4, 20, fmt_hdr)
        ws.freeze_panes(5, 0)

        # Bordas/format por coluna
        start_row = 5
        start_col = 0
        nrows = len(df)
        ncols = len(df.columns)

        # Aplicar formatos em colunas específicas
        col_idx = {c: i for i, c in enumerate(df.columns)}
        for r in range(nrows):
            excel_r = start_row + r
            for c, j in col_idx.items():
                val = df.iloc[r, j]
                if c == "VALOR":
                    if pd.notna(val):
                        ws.write_number(excel_r, start_col + j, float(val), fmt_money)
                    else:
                        ws.write_blank(excel_r, start_col + j, None, fmt_money)
                elif c == "DATA":
                    if pd.notna(val):
                        ws.write_datetime(excel_r, start_col + j, val.to_pydatetime(), fmt_date)
                    else:
                        ws.write_blank(excel_r, start_col + j, None, fmt_date)
                else:
                    ws.write(excel_r, start_col + j, "" if pd.isna(val) else str(val), fmt_txt)

        # Tabela com filtro
        ws.autofilter(4, 0, 4 + nrows, max(ncols - 1, 0))

        # Auto fit
        _autofit_worksheet(ws, df, header_row=4, start_col=0)

        # Aba Tratativa
        trat = build_tratativa(df_filtrado)
        trat.to_excel(w, index=False, sheet_name="Tratativa")
        ws2 = w.sheets["Tratativa"]
        ws2.freeze_panes(1, 0)
        ws2.set_row(0, 20, fmt_hdr)
        _autofit_worksheet(ws2, trat, header_row=0, start_col=0, max_width=65)

        # Aba Resumo (executivo no Excel)
        resumo = pd.DataFrame([
            ["Saldo anterior (antes do 1º movimento) – Financeiro", stats.get("saldo_ant_fin", np.nan)],
            ["Saldo anterior (antes do 1º movimento) – Contábil", stats.get("saldo_ant_led", np.nan)],
            ["Diferença de saldo anterior (Fin - Cont)", stats.get("diff_saldo_ant", np.nan)],
            ["" , ""],
            ["Saldo final (último movimento) – Financeiro", stats.get("saldo_fin", np.nan)],
            ["Saldo final (último movimento) – Contábil", stats.get("saldo_led", np.nan)],
            ["Diferença final (Fin - Cont)", stats.get("diff_final", np.nan)],
            ["" , ""],
            ["Soma dos movimentos só no Financeiro (divergências)", stats.get("fin_pend_val", 0.0)],
            ["Soma dos movimentos só no Contábil (divergências)", stats.get("led_pend_val", 0.0)],
            ["Impacto líquido dos movimentos (Fin - Cont)", stats.get("impacto", 0.0)],
            ["Diferença esperada (Dif. saldo anterior + Impacto)", stats.get("diff_esperada", np.nan)],
            ["Conferência (Dif. final - Dif. esperada) -> precisa zerar", stats.get("conferencia", np.nan)],
        ], columns=["Métrica", "Valor"])
        resumo.to_excel(w, index=False, sheet_name="Resumo")
        ws3 = w.sheets["Resumo"]
        ws3.freeze_panes(1, 0)
        ws3.set_row(0, 20, fmt_hdr)
        ws3.set_column(0, 0, 62)
        ws3.set_column(1, 1, 22, wb.add_format({"num_format": "#,##0.00"}))

    out.seek(0)
    return out

# ----------------------------
# PDF Resumo (sem gráfico)
# ----------------------------
def to_pdf_resumo(stats, generated_at):
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36
    )
    styles = getSampleStyleSheet()
    story = []

    title = "Relatório de Resumo — ConciliaMais (Conferência de Extrato Bancário)"
    story.append(Paragraph(title, styles["Title"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"Processado em: {generated_at}", styles["Normal"]))
    story.append(Spacer(1, 14))

    # KPIs
    fin_total = int(stats.get("fin_total", 0))
    led_total = int(stats.get("led_total", 0))
    pairs = int(stats.get("fin_conc", 0))
    total_itens = fin_total + led_total
    matched_itens = pairs * 2
    pend_itens = max(total_itens - matched_itens, 0)
    pct = (matched_itens / total_itens * 100.0) if total_itens else 0.0

    kpi_data = [
        ["Indicador", "Valor"],
        ["MATCH (itens)", f"{matched_itens} / {total_itens}  ({pct:.1f}%)"],
        ["Pendentes (itens)", f"{pend_itens}  (Financeiro: {stats.get('fin_pend',0)} | Contábil: {stats.get('led_pend',0)})"],
        ["Valor a conciliar (Fin - Cont)", fmt(stats.get("impacto", 0.0))],
        ["Conferência (ideal 0,00)", fmt(stats.get("conferencia", np.nan))],
    ]

    t1 = Table(kpi_data, colWidths=[220, 280])
    t1.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#111827")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#D1D5DB")),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F9FAFB")]),
    ]))
    story.append(Paragraph("Resumo executivo", styles["Heading2"]))
    story.append(Spacer(1, 6))
    story.append(t1)
    story.append(Spacer(1, 14))

    # Composição de saldos
    comp = [
        ["Composição de saldos", "Valor"],
        ["Saldo anterior (antes do 1º movimento) – Financeiro", fmt(stats.get("saldo_ant_fin", np.nan))],
        ["Saldo anterior (antes do 1º movimento) – Contábil", fmt(stats.get("saldo_ant_led", np.nan))],
        ["Diferença de saldo anterior (Fin - Cont)", fmt(stats.get("diff_saldo_ant", np.nan))],
        ["", ""],
        ["Saldo final (último movimento) – Financeiro", fmt(stats.get("saldo_fin", np.nan))],
        ["Saldo final (último movimento) – Contábil", fmt(stats.get("saldo_led", np.nan))],
        ["Diferença final (Fin - Cont)", fmt(stats.get("diff_final", np.nan))],
        ["", ""],
        ["Soma dos movimentos só no Financeiro (divergências)", fmt(stats.get("fin_pend_val", 0.0))],
        ["Soma dos movimentos só no Contábil (divergências)", fmt(stats.get("led_pend_val", 0.0))],
        ["Impacto líquido dos movimentos (Fin - Cont)", fmt(stats.get("impacto", 0.0))],
        ["Diferença esperada (Dif. saldo anterior + Impacto)", fmt(stats.get("diff_esperada", np.nan))],
        ["Conferência (Dif. final - Dif. esperada) -> precisa zerar", fmt(stats.get("conferencia", np.nan))],
    ]
    t2 = Table(comp, colWidths=[340, 160])
    t2.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#111827")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#D1D5DB")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F9FAFB")]),
    ]))
    story.append(Paragraph("Composição de saldos", styles["Heading2"]))
    story.append(Spacer(1, 6))
    story.append(t2)

    doc.build(story)
    buf.seek(0)
    return buf

# ----------------------------
# Navegação por estado
# ----------------------------
if "page" not in st.session_state:
    st.session_state.page = "upload"
if "results" not in st.session_state:
    st.session_state.results = None

# ----------------------------
# Página: Upload
# ----------------------------
if st.session_state.page == "upload":
    st.title("ConciliaMais — Conferência de Extrato Bancário")
    st.caption("Extrato Financeiro + Razão Contábil → Match automático → Divergências → Painel de fechamento")

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

    div = res["div"].copy()
    stats = res["stats"]
    generated_at = res["generated_at"]

    # LIMPEZA: remover VALOR 0/NaN e limpar nan/None em textos
    div["VALOR"] = div["VALOR"].map(normalize_money)
    div = div[div["VALOR"].notna()].copy()
    div = div[div["VALOR"].abs() > 0.0000001].copy()

    for c in ["DOCUMENTO", "PREFIXO_TITULO", "HISTORICO_OPERACAO", "CHAVE_DOC"]:
        if c in div.columns:
            div[c] = div[c].replace({np.nan: "", "nan": "", "None": ""}).astype(str).str.strip()

    # Documento padronizado:
    # Financeiro: DOCUMENTO = PREFIXO_TITULO (quando existir)
    mask_fin = div["ORIGEM"].eq("Somente Financeiro")
    if "PREFIXO_TITULO" in div.columns:
        div.loc[mask_fin, "DOCUMENTO"] = div.loc[mask_fin, "PREFIXO_TITULO"].where(
            div.loc[mask_fin, "PREFIXO_TITULO"].astype(str).str.len() > 0,
            div.loc[mask_fin, "DOCUMENTO"]
        )

    # Contábil: mantém lógica do reconcile; reforço se vier vazio
    mask_led = div["ORIGEM"].eq("Somente Contábil")
    if "HISTORICO_OPERACAO" in div.columns:
        missing = mask_led & (div["DOCUMENTO"].astype(str).str.len() == 0)
        div.loc[missing, "DOCUMENTO"] = div.loc[missing, "HISTORICO_OPERACAO"].map(extract_doc_from_ledger_history)

    # Remover coluna PREFIXO_TITULO (não queremos mais ver)
    if "PREFIXO_TITULO" in div.columns:
        div = div.drop(columns=["PREFIXO_TITULO"])

    # Remover coluna CONTA, se existir
    if "CONTA" in div.columns:
        div = div.drop(columns=["CONTA"])

    # Processado em
    st.caption(f"Processado em: {generated_at}")

    # ----------------
    # Cards (ajustes)
    # ----------------
    pairs = int(stats.get("fin_conc", 0))
    fin_total = int(stats.get("fin_total", 0))
    led_total = int(stats.get("led_total", 0))
    total_itens = fin_total + led_total
    matched_itens = pairs * 2
    pend_itens = max(total_itens - matched_itens, 0)
    pct = (matched_itens / total_itens * 100.0) if total_itens else 0.0

    st.markdown(
        f"""
<div class="cm-cards">
  <div class="cm-card">
    <div class="k">MATCH (geral)</div>
    <div class="v">{matched_itens} / {total_itens}</div>
    <div class="s">{pct:.1f}% conciliado (itens dos dois lados)</div>
  </div>
  <div class="cm-card">
    <div class="k">Pendentes</div>
    <div class="v">{pend_itens}</div>
    <div class="s">Financeiro: {stats.get("fin_pend",0)} | Contábil: {stats.get("led_pend",0)}</div>
  </div>
  <div class="cm-card">
    <div class="k">Valor a conciliar</div>
    <div class="v">{fmt(stats.get("impacto", 0.0))}</div>
    <div class="s"></div>
  </div>
  <div class="cm-card">
    <div class="k">Conferência (ideal 0,00)</div>
    <div class="v">{fmt(stats.get("conferencia", np.nan))}</div>
    <div class="s">{status_pill(stats.get("conferencia", np.nan))}</div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    st.markdown('<div class="cm-section"></div>', unsafe_allow_html=True)

    # ----------------
    # Gráfico (mantém extração)
    # ----------------
    st.markdown("### Divergências (Financeiro x Contábil)")
    fin_val = float(stats.get("fin_pend_val", 0.0))
    led_val = float(stats.get("led_pend_val", 0.0))
    chart_df = pd.DataFrame(
        {"Valor": [fin_val, led_val]},
        index=["Somente Financeiro", "Somente Contábil"],
    )
    st.bar_chart(chart_df)

    # ----------------
    # Divergências: filtros + total alinhado
    # ----------------
    st.markdown("### Divergências (itens não pareados)")

    fcol1, fcol2, fcol3, fcol4 = st.columns([1.25, 1.0, 2.25, 1.1])
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
        cols_search = [c for c in df.columns if c in ["DOCUMENTO", "HISTORICO_OPERACAO", "CHAVE_DOC"]]
        mask = False
        for c in cols_search:
            mask = mask | df[c].astype(str).str.lower().str.contains(q, na=False)
        df = df[mask].copy()

    if ordenar in df.columns:
        df = df.sort_values(by=ordenar, ascending=True)

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

    # Garantir ordem de colunas “limpa”
    col_order = [c for c in ["ORIGEM", "DATA", "DOCUMENTO", "HISTORICO_OPERACAO", "CHAVE_DOC", "VALOR"] if c in df.columns]
    df = df[col_order].copy()

    st.dataframe(df, use_container_width=True, height=420)

    # ----------------
    # Exportar
    # ----------------
    st.markdown("### Exportar")

    excel_bytes = to_excel_divergencias_filtradas(df, total_filtrado, stats, generated_at)
    st.download_button(
        "Baixar Divergências (Excel) — exatamente como filtrado",
        data=excel_bytes,
        file_name=f"ConciliaMais_DivergenciasFiltradas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    pdf_bytes = to_pdf_resumo(stats, generated_at)
    st.download_button(
        "Baixar Relatório Resumo (PDF)",
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
            st.session_state.page = "upload"
            st.rerun()
