import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime

# PDF (relatório resumo)
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="ConciliaMais — Módulo 1", layout="wide")

# ----------------------------
# CSS (visual mais “produto”)
# ----------------------------
st.markdown(
    """
<style>
.block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
h1, h2, h3 { letter-spacing: -0.02em; }

.cm-cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-top: 10px; }
.cm-card {
  border-radius: 16px;
  padding: 14px 14px 12px 14px;
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.08);
}
.cm-card .k { font-size: 12px; opacity: .82; margin-bottom: 6px; }
.cm-card .v { font-size: 24px; font-weight: 750; }
.cm-card .s { font-size: 12px; opacity: .78; margin-top: 6px; }
.cm-pill { display: inline-block; padding: 4px 10px; border-radius: 999px; font-size: 12px; font-weight: 650; }
.cm-ok { background: rgba(34,197,94,.18); color: rgb(134,239,172); border: 1px solid rgba(34,197,94,.35); }
.cm-warn { background: rgba(245,158,11,.18); color: rgb(253,230,138); border: 1px solid rgba(245,158,11,.35); }
.cm-bad { background: rgba(239,68,68,.18); color: rgb(254,202,202); border: 1px solid rgba(239,68,68,.35); }

.cm-section { margin-top: 18px; }
.cm-subtle { opacity: .80; font-size: 13px; }

.cm-filterbar {
  display: grid;
  grid-template-columns: 1fr 1fr 2fr 1fr;
  gap: 12px;
  align-items: end;
  margin-top: 10px;
}

.cm-totalbox {
  border-radius: 16px;
  padding: 12px 14px;
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.10);
  text-align: right;
}
.cm-totalbox .k { font-size: 12px; opacity: .82; }
.cm-totalbox .v { font-size: 22px; font-weight: 750; margin-top: 4px; }

.cm-actions {
  border-radius: 16px;
  padding: 14px;
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.08);
  margin-top: 12px;
}
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

def status_pill(conferencia_calculo):
    if conferencia_calculo is None or (isinstance(conferencia_calculo, float) and np.isnan(conferencia_calculo)):
        return '<span class="cm-pill cm-warn">Cálculo não disponível</span>'
    x = abs(float(conferencia_calculo))
    if x <= 0.01:
        return '<span class="cm-pill cm-ok">Cálculo OK</span>'
    if x <= 5:
        return '<span class="cm-pill cm-warn">Quase (revisar)</span>'
    return '<span class="cm-pill cm-bad">Cálculo divergente</span>'

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
    conferencia_calculo = np.nan if (pd.isna(diff_final) or pd.isna(diff_esperada)) else round(diff_final - diff_esperada, 2)

    # Divergências humanizadas
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
            "PREFIXO_TITULO": "",  # não usado no contábil
            "HISTORICO_OPERACAO": str(base.get(cfg.get("led_historico"), "")) if cfg.get("led_historico") else str(r["__text"]),
            "CHAVE_DOC": r["__doc_key"],
            "VALOR": round(float(r["__amount"]), 2) if pd.notna(r["__amount"]) else np.nan,
        })

    div = pd.concat([pd.DataFrame(fin_rows), pd.DataFrame(led_rows)], ignore_index=True)

    stats = {
        "fin_total": int(len(f)),
        "led_total": int(len(l)),
        "fin_conc": int(len(fin_match)),
        "led_conc": int(len(led_match)),
        "fin_pend": int(len(f) - len(fin_match)),
        "led_pend": int(len(l) - len(led_match)),
        "fin_pend_val": float(fin_unmatched),
        "led_pend_val": float(led_unmatched),
        "impacto": float(impacto),
        "saldo_ant_fin": float(saldo_ant_fin) if pd.notna(saldo_ant_fin) else np.nan,
        "saldo_ant_led": float(saldo_ant_led) if pd.notna(saldo_ant_led) else np.nan,
        "saldo_fin": float(saldo_fin) if pd.notna(saldo_fin) else np.nan,
        "saldo_led": float(saldo_led) if pd.notna(saldo_led) else np.nan,
        "diff_saldo_ant": float(diff_saldo_ant) if pd.notna(diff_saldo_ant) else np.nan,
        "diff_final": float(diff_final) if pd.notna(diff_final) else np.nan,
        "diff_esperada": float(diff_esperada) if pd.notna(diff_esperada) else np.nan,
        "conferencia_calculo": float(conferencia_calculo) if pd.notna(conferencia_calculo) else np.nan,
    }

    return div, stats

# ----------------------------
# Camada UX: tratativa editável + regras
# ----------------------------
NUCLEOS_PADRAO = [
    "",  # vazio = não analisado
    "Processo interno",
    "Cadastro",
    "Configuração RP",
    "Não identificado",
]

def preparar_tratativa(div_raw: pd.DataFrame) -> pd.DataFrame:
    """Normaliza colunas, remove NaN/0 no que não interessa e cria colunas UX."""
    df = div_raw.copy()

    # limpar NaN visuais
    for c in ["DOCUMENTO", "PREFIXO_TITULO", "HISTORICO_OPERACAO", "CHAVE_DOC", "ORIGEM"]:
        if c in df.columns:
            df[c] = df[c].replace({np.nan: ""}).astype(str)

    # Remover linhas sem valor ou valor zero (não interessa)
    df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce")
    df = df[df["VALOR"].notna()].copy()
    df = df[df["VALOR"].abs() > 0].copy()

    # Unificar "DOCUMENTO": no financeiro, usar PREFIXO_TITULO como documento principal (se existir)
    df["DOCUMENTO_UNIFICADO"] = df["DOCUMENTO"].astype(str).str.strip()
    mask_fin = df["ORIGEM"].eq("Somente Financeiro")
    pre = df["PREFIXO_TITULO"].astype(str).str.strip()
    df.loc[mask_fin & pre.ne(""), "DOCUMENTO_UNIFICADO"] = pre[mask_fin & pre.ne("")]
    # no contábil, manter como está (a chave fica em CHAVE_DOC)

    # UX columns
    if "NUCLEO_CONFIRMADO" not in df.columns:
        df["NUCLEO_CONFIRMADO"] = ""
    if "RESOLVIDO" not in df.columns:
        df["RESOLVIDO"] = "Não"  # Sim/Não
    if "STATUS" not in df.columns:
        df["STATUS"] = "Pendente"
    if "OBS" not in df.columns:
        df["OBS"] = ""

    # Status coerente
    df["RESOLVIDO"] = df["RESOLVIDO"].replace({True: "Sim", False: "Não"}).astype(str)
    df["RESOLVIDO"] = df["RESOLVIDO"].apply(lambda x: "Sim" if str(x).strip().lower() in ["sim", "s", "true", "1"] else "Não")
    df["NUCLEO_CONFIRMADO"] = df["NUCLEO_CONFIRMADO"].astype(str).replace({np.nan: ""})
    df["STATUS"] = df.apply(lambda r: "Resolvido" if r["RESOLVIDO"] == "Sim" else "Pendente", axis=1)

    # ID interno para permitir ações por ID
    if "ID" not in df.columns:
        df.insert(0, "ID", np.arange(1, len(df) + 1))

    # Colunas finais da tabela (mais limpa)
    cols = ["ID", "ORIGEM", "DATA", "DOCUMENTO_UNIFICADO", "HISTORICO_OPERACAO", "CHAVE_DOC", "VALOR", "NUCLEO_CONFIRMADO", "RESOLVIDO", "STATUS", "OBS"]
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols].copy()
    df = df.rename(columns={"DOCUMENTO_UNIFICADO": "DOCUMENTO"})

    # DATA format
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce").dt.date

    return df

def calcular_kpis(trat_df: pd.DataFrame, conferencia_calculo: float):
    total_itens = int(len(trat_df))
    pendentes_df = trat_df[trat_df["RESOLVIDO"] != "Sim"].copy()
    resolvidos_df = trat_df[trat_df["RESOLVIDO"] == "Sim"].copy()

    pend_qtd = int(len(pendentes_df))
    res_qtd = int(len(resolvidos_df))

    pend_val = float(pendentes_df["VALOR"].sum()) if total_itens else 0.0

    return {
        "dif_qtd": total_itens,
        "pend_qtd": pend_qtd,
        "pend_val": pend_val,
        "res_qtd": res_qtd,
        "total_qtd": total_itens,
        "conferencia_calculo": conferencia_calculo,
    }

def excel_formatado(df_export: pd.DataFrame, titulo: str, incluir_total=True) -> BytesIO:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        df_export.to_excel(w, index=False, sheet_name="Divergencias")
        wb = w.book
        ws = w.sheets["Divergencias"]

        fmt_hdr = wb.add_format({"bold": True, "border": 1, "align": "center", "valign": "vcenter"})
        fmt_txt = wb.add_format({"border": 1})
        fmt_money = wb.add_format({"num_format": "#,##0.00", "border": 1})
        fmt_money_bold = wb.add_format({"num_format": "#,##0.00", "border": 1, "bold": True})

        ws.freeze_panes(1, 0)
        ws.set_row(0, 22, fmt_hdr)

        # Larguras (sem precisar ajustar depois)
        widths = {
            "ID": 8,
            "ORIGEM": 18,
            "DATA": 12,
            "DOCUMENTO": 20,
            "HISTORICO_OPERACAO": 46,
            "CHAVE_DOC": 16,
            "VALOR": 16,
            "NUCLEO_CONFIRMADO": 18,
            "RESOLVIDO": 10,
            "STATUS": 12,
            "OBS": 26,
        }
        for i, col in enumerate(df_export.columns):
            ws.set_column(i, i, widths.get(col, 18))

        # Formatar valores
        col_idx = {c: i for i, c in enumerate(df_export.columns)}
        if "VALOR" in col_idx:
            ws.set_column(col_idx["VALOR"], col_idx["VALOR"], widths.get("VALOR", 16), fmt_money)

        # Bordas em tudo
        for r in range(1, len(df_export) + 1):
            ws.set_row(r, 18)

        # Linha total
        if incluir_total and "VALOR" in col_idx:
            last_row = len(df_export) + 1
            ws.write(last_row, 0, "TOTAL", fmt_hdr)
            for c in range(1, len(df_export.columns)):
                ws.write(last_row, c, "", fmt_txt)
            ws.write_formula(
                last_row,
                col_idx["VALOR"],
                f"=SUM({xlsx_col(col_idx['VALOR'])}2:{xlsx_col(col_idx['VALOR'])}{len(df_export)+1})",
                fmt_money_bold,
            )

        # Título no topo
        ws.write(0, 0, df_export.columns[0], fmt_hdr)  # reforça header
        # (Título vai no nome do arquivo e no PDF; aqui mantemos simples)
    out.seek(0)
    return out

def xlsx_col(n: int) -> str:
    """0->A, 25->Z, 26->AA..."""
    s = ""
    n2 = n
    while True:
        n2, r = divmod(n2, 26)
        s = chr(65 + r) + s
        if n2 == 0:
            break
        n2 -= 1
    return s

def gerar_pdf_resumo(kpis, stats, periodo_txt: str = "") -> BytesIO:
    out = BytesIO()
    doc = SimpleDocTemplate(out, pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=1.8*cm, bottomMargin=1.6*cm)
    styles = getSampleStyleSheet()
    elems = []

    titulo = "Relatório de Conciliação — Financeiro x Contábil (Módulo 1)"
    elems.append(Paragraph(titulo, styles["Title"]))
    elems.append(Spacer(1, 10))

    meta = f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    if periodo_txt:
        meta += f" | Período: {periodo_txt}"
    elems.append(Paragraph(meta, styles["Normal"]))
    elems.append(Spacer(1, 12))

    # Bloco 1: Diagnóstico
    elems.append(Paragraph("1) Diagnóstico", styles["Heading2"]))
    diag_data = [
        ["Diferenças encontradas (itens)", str(kpis["dif_qtd"])],
        ["Pendências em aberto (itens)", str(kpis["pend_qtd"])],
        ["Pendências em aberto (valor)", f"R$ {fmt(kpis['pend_val'])}"],
        ["Progresso resolvido", f"{kpis['res_qtd']} / {kpis['total_qtd']}"],
    ]
    t = Table(diag_data, colWidths=[10*cm, 5*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
        ("BOX", (0,0), (-1,-1), 0.6, colors.grey),
        ("INNERGRID", (0,0), (-1,-1), 0.4, colors.lightgrey),
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 10),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (1,0), (1,-1), "RIGHT"),
    ]))
    elems.append(t)
    elems.append(Spacer(1, 14))

    # Bloco 2: Composição dos saldos / cálculo
    elems.append(Paragraph("2) Composição (saldos e cálculo)", styles["Heading2"]))
    comp = [
        ["Saldo anterior (antes do 1º movimento) — Financeiro", f"R$ {fmt(stats.get('saldo_ant_fin'))}"],
        ["Saldo anterior (antes do 1º movimento) — Contábil", f"R$ {fmt(stats.get('saldo_ant_led'))}"],
        ["Diferença de saldo anterior (Fin - Cont)", f"R$ {fmt(stats.get('diff_saldo_ant'))}"],
        ["Saldo final (último movimento) — Financeiro", f"R$ {fmt(stats.get('saldo_fin'))}"],
        ["Saldo final (último movimento) — Contábil", f"R$ {fmt(stats.get('saldo_led'))}"],
        ["Diferença final (Fin - Cont)", f"R$ {fmt(stats.get('diff_final'))}"],
        ["Impacto líquido dos pendentes (Fin - Cont)", f"R$ {fmt(stats.get('impacto'))}"],
        ["Diferença esperada (Dif. saldo anterior + Impacto)", f"R$ {fmt(stats.get('diff_esperada'))}"],
        ["Conferência do cálculo (Dif. final - Dif. esperada) → ideal 0,00", f"R$ {fmt(stats.get('conferencia_calculo'))}"],
    ]
    t2 = Table(comp, colWidths=[11.3*cm, 3.7*cm])
    t2.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 0.6, colors.grey),
        ("INNERGRID", (0,0), (-1,-1), 0.4, colors.lightgrey),
        ("FONTSIZE", (0,0), (-1,-1), 9.5),
        ("ALIGN", (1,0), (1,-1), "RIGHT"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ]))
    elems.append(t2)
    elems.append(Spacer(1, 10))

    # Observação final
    obs = "Observação: 'Conferência do cálculo' valida a coerência matemática do fechamento. Já o progresso de 'Resolvido' depende da tratativa/classificação feita pelo usuário."
    elems.append(Paragraph(obs, styles["Normal"]))

    doc.build(elems)
    out.seek(0)
    return out

# ----------------------------
# Navegação por estado
# ----------------------------
if "page" not in st.session_state:
    st.session_state.page = "upload"
if "results" not in st.session_state:
    st.session_state.results = None
if "trat" not in st.session_state:
    st.session_state.trat = None

# ----------------------------
# Página: Upload
# ----------------------------
if st.session_state.page == "upload":
    st.title("ConciliaMais — Módulo 1")
    st.caption("Upload do Extrato Financeiro + Razão Contábil → Match automático → Divergências → Tratativa")

    c1, c2 = st.columns(2)
    with c1:
        fin_file = st.file_uploader("Extrato Financeiro (.xlsx ou .csv)", type=["xlsx", "csv"], key="fin")
    with c2:
        led_file = st.file_uploader("Razão Contábil (.xlsx ou .csv)", type=["xlsx", "csv"], key="led")

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

        trat = preparar_tratativa(div)

        st.session_state.results = {
            "div_raw": div,
            "stats": stats,
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        st.session_state.trat = trat
        st.session_state.page = "resultados"
        st.rerun()

# ----------------------------
# Página: Resultados
# ----------------------------
else:
    res = st.session_state.results
    if not res or st.session_state.trat is None:
        st.session_state.page = "upload"
        st.rerun()

    stats = res["stats"]
    trat = st.session_state.trat.copy()

    st.title("Resultados — Módulo 1")
    st.caption(f"Processado em: {res['generated_at']}")

    # ===== INDICADORES NO TOPO (como você pediu) =====
    kpis = calcular_kpis(trat, stats.get("conferencia_calculo", np.nan))

    perc_res = (kpis["res_qtd"] / max(kpis["total_qtd"], 1)) * 100.0
    perc_res_txt = f"{perc_res:.1f}% resolvido"

    st.markdown(
        f"""
<div class="cm-cards">
  <div class="cm-card">
    <div class="k">Diferenças encontradas</div>
    <div class="v">{kpis["dif_qtd"]}</div>
    <div class="s">itens de divergência identificados</div>
  </div>
  <div class="cm-card">
    <div class="k">Pendências em aberto</div>
    <div class="v">{fmt(kpis["pend_val"])}</div>
    <div class="s">{kpis["pend_qtd"]} itens em aberto</div>
  </div>
  <div class="cm-card">
    <div class="k">Progresso resolvido</div>
    <div class="v">{kpis["res_qtd"]} / {kpis["total_qtd"]}</div>
    <div class="s">{perc_res_txt}</div>
  </div>
  <div class="cm-card">
    <div class="k">Conferência do cálculo</div>
    <div class="v">{fmt(stats.get("conferencia_calculo"))}</div>
    <div class="s">{status_pill(stats.get("conferencia_calculo"))}</div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    st.markdown('<div class="cm-section"></div>', unsafe_allow_html=True)

    # ===== FILTROS + TOTAL DO FILTRO ALINHADO =====
    st.markdown("### Divergências (itens não pareados)")
    st.markdown('<div class="cm-subtle">Filtre, classifique o núcleo e marque como resolvido. As pendências em aberto reduzem conforme você resolve.</div>', unsafe_allow_html=True)

    # Barra de filtros em grid (inclui total alinhado)
    f1, f2, f3, f4 = st.columns([1, 1, 2, 1])

    with f1:
        origem = st.selectbox("Filtrar por origem", ["Todas", "Somente Financeiro", "Somente Contábil"], index=0)
    with f2:
        ordenar = st.selectbox("Ordenar por", ["DATA", "VALOR"], index=0)
    with f3:
        busca = st.text_input("Buscar (documento, histórico, chave)", value="")
    with f4:
        # total calculado após filtros
        pass

    df = trat.copy()

    if origem != "Todas":
        df = df[df["ORIGEM"] == origem].copy()

    if busca.strip():
        q = busca.strip().lower()
        cols_search = ["DOCUMENTO", "HISTORICO_OPERACAO", "CHAVE_DOC", "NUCLEO_CONFIRMADO", "STATUS"]
        mask = False
        for c in cols_search:
            mask = mask | df[c].astype(str).str.lower().str.contains(q, na=False)
        df = df[mask].copy()

    if ordenar in df.columns:
        df = df.sort_values(by=ordenar, ascending=True)

    total_filtro = float(df["VALOR"].sum()) if len(df) else 0.0
    with f4:
        st.markdown(
            f"""
<div class="cm-totalbox">
  <div class="k">Total do filtro</div>
  <div class="v">{fmt(total_filtro)}</div>
</div>
""",
            unsafe_allow_html=True,
        )

    # ===== TABELA EDITÁVEL (UX de tratativa) =====
    st.markdown("#### Lista de divergências (tratativa)")

    # Config do editor: núcleo como select + resolvido Sim/Não
    edited = st.data_editor(
        df,
        use_container_width=True,
        height=420,
        hide_index=True,
        column_config={
            "VALOR": st.column_config.NumberColumn("VALOR", format="%.2f"),
            "NUCLEO_CONFIRMADO": st.column_config.SelectboxColumn("NÚCLEO CONFIRMADO", options=NUCLEOS_PADRAO),
            "RESOLVIDO": st.column_config.SelectboxColumn("RESOLVIDO", options=["Não", "Sim"]),
            "STATUS": st.column_config.TextColumn("STATUS", disabled=True),
        },
        disabled=["STATUS"],  # status calculado
    )

    # Reaplicar regras: RESOLVIDO -> STATUS, e RESOLVIDO=Sim exige Núcleo
    def aplicar_regras_df(dfx: pd.DataFrame) -> pd.DataFrame:
        out = dfx.copy()
        out["RESOLVIDO"] = out["RESOLVIDO"].apply(lambda x: "Sim" if str(x).strip().lower() == "sim" else "Não")
        out["NUCLEO_CONFIRMADO"] = out["NUCLEO_CONFIRMADO"].replace({np.nan: ""}).astype(str)
        # Se resolvido, núcleo obrigatório
        invalid = out[(out["RESOLVIDO"] == "Sim") & (out["NUCLEO_CONFIRMADO"].astype(str).str.strip() == "")]
        if not invalid.empty:
            st.error("Para marcar como RESOLVIDO = Sim, é obrigatório preencher o NÚCLEO CONFIRMADO nos itens correspondentes.")
        out["STATUS"] = out.apply(lambda r: "Resolvido" if r["RESOLVIDO"] == "Sim" else "Pendente", axis=1)
        return out

    edited = aplicar_regras_df(edited)

    # Persistir no state: atualiza somente os IDs que estão na tela filtrada
    # (mantém o restante do dataset)
    base = st.session_state.trat.copy()
    base = base.set_index("ID", drop=False)
    edited_idx = edited.set_index("ID", drop=False)
    for cid in edited.columns:
        base.loc[edited_idx.index, cid] = edited_idx[cid]
    base = aplicar_regras_df(base.reset_index(drop=True))
    st.session_state.trat = base

    # ===== AÇÕES EM MASSA =====
    st.markdown("### Ações em massa")
    st.markdown('<div class="cm-subtle">Aplique núcleo/resultado para itens filtrados (tela atual) ou para IDs específicos.</div>', unsafe_allow_html=True)

    with st.container():
        st.markdown('<div class="cm-actions">', unsafe_allow_html=True)
        alvo = st.radio("Aplicar em:", ["Itens filtrados (tela atual)", "IDs informados"], horizontal=True)

        cA, cB, cC = st.columns([1, 1, 2])
        with cA:
            resolvido_bulk = st.selectbox("Marcar como Resolvido", ["(não alterar)", "Não", "Sim"], index=0)
        with cB:
            nucleo_bulk = st.selectbox("Núcleo confirmado", NUCLEOS_PADRAO, index=0)
        with cC:
            ids_txt = ""
            if alvo == "IDs informados":
                ids_txt = st.text_input("IDs (separe por vírgula, espaço ou quebra de linha)", value="")

        aplicar = st.button("Aplicar", type="primary")
        st.markdown("</div>", unsafe_allow_html=True)

    if aplicar:
        base = st.session_state.trat.copy()
        alvo_ids = None

        if alvo == "Itens filtrados (tela atual)":
            alvo_ids = df["ID"].tolist()
        else:
            nums = re.findall(r"\d+", ids_txt or "")
            alvo_ids = [int(n) for n in nums] if nums else []

        if not alvo_ids:
            st.warning("Nenhum item selecionado para aplicação.")
        else:
            upd = base.copy()
            mask = upd["ID"].isin(alvo_ids)

            # aplicar núcleo
            if str(nucleo_bulk).strip() != "":
                upd.loc[mask, "NUCLEO_CONFIRMADO"] = nucleo_bulk

            # aplicar resolvido
            if resolvido_bulk != "(não alterar)":
                if resolvido_bulk == "Sim":
                    # exigir núcleo
                    falta_nucleo = upd.loc[mask, "NUCLEO_CONFIRMADO"].astype(str).str.strip().eq("")
                    if falta_nucleo.any():
                        st.error("Para marcar como RESOLVIDO = Sim em massa, todos os itens precisam ter NÚCLEO CONFIRMADO preenchido (ou selecione um núcleo antes de aplicar).")
                    else:
                        upd.loc[mask, "RESOLVIDO"] = "Sim"
                else:
                    upd.loc[mask, "RESOLVIDO"] = "Não"

            upd["STATUS"] = upd.apply(lambda r: "Resolvido" if r["RESOLVIDO"] == "Sim" else "Pendente", axis=1)
            st.session_state.trat = upd
            st.success(f"Aplicado em {int(mask.sum())} item(ns).")
            st.rerun()

    # ===== EXPORTAÇÃO (FORMATO EXCEL, DO JEITO DA TELA) =====
    st.markdown("### Exportar")
    st.markdown('<div class="cm-subtle">Exporta exatamente o que está filtrado na tela (com colunas e larguras ajustadas).</div>', unsafe_allow_html=True)

    export_df = df.copy()
    # ordenar por padrão atual
    if ordenar in export_df.columns:
        export_df = export_df.sort_values(by=ordenar, ascending=True)

    # Excel formatado do filtro atual
    excel_bytes = excel_formatado(export_df, titulo="Divergências filtradas", incluir_total=True)
    st.download_button(
        "Baixar Excel (filtro atual)",
        data=excel_bytes,
        file_name=f"ConciliaMais_Divergencias_Filtro_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # PDF resumo (relatório para diretoria)
    pdf_bytes = gerar_pdf_resumo(kpis, stats, periodo_txt="")
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
            st.session_state.trat = None
            st.session_state.page = "upload"
            st.rerun()
