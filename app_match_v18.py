import streamlit as st
import pandas as pd
import numpy as np
import re
import json
import os
import unicodedata
import hashlib
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook

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
LEARNING_FILE = "aprendizado.json"
NUCLEOS_FILE = "nucleos.json"
ASSOCIATION_RULES_FILE = "regras_associacao_campos.json"

DEFAULT_NUCLEOS = [
    "Processo interno",
    "Cadastro",
    "Configuração ERP",
    "Não identificado",
]

STATUS_OPTS = ["Pendente", "Em análise", "Resolvido"]
SEVERIDADES = ["Normal", "Atenção", "Crítica"]
ORIGEM_RULE_OPTS = ["Qualquer", "Somente Financeiro", "Somente Contábil"]

SEVERITY_ORDER = {"Crítica": 3, "Atenção": 2, "Normal": 1}

# =========================================================
# CSS
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
  max-width: 1480px;
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

.cm-subtle{
  color: var(--muted);
  font-size: 12px;
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
# Helpers gerais
# =========================================================
def set_flash(kind, msg):
    st.session_state["_flash"] = {"kind": kind, "msg": msg}

def show_flash():
    flash = st.session_state.pop("_flash", None)
    if flash:
        if flash["kind"] == "success":
            st.success(flash["msg"])
        elif flash["kind"] == "warning":
            st.warning(flash["msg"])
        elif flash["kind"] == "error":
            st.error(flash["msg"])
        else:
            st.info(flash["msg"])

def safe_float(x, default=None):
    try:
        if pd.isna(x):
            return default
        return float(x)
    except Exception:
        return default

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

def normalize_money(x):
    if pd.isna(x):
        return np.nan

    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)

    s = str(x).strip()
    if s == "":
        return np.nan

    s = s.replace("R$", "").replace("r$", "").replace(" ", "").replace("\u00a0", "")

    if s in {"-", "--"}:
        return np.nan

    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    if s.startswith("-"):
        neg = True
        s = s[1:]

    s = re.sub(r"[^0-9,\.]", "", s)
    if s == "":
        return np.nan

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        if s.count(",") == 1:
            left, right = s.split(",")
            if len(right) in (1, 2):
                s = left + "." + right
            elif len(right) == 3 and len(left) > 3:
                s = left + right
            else:
                s = left + "." + right
        else:
            parts = s.split(",")
            if len(parts[-1]) in (1, 2):
                s = "".join(parts[:-1]) + "." + parts[-1]
            else:
                s = "".join(parts)
    elif "." in s:
        if s.count(".") == 1:
            left, right = s.split(".")
            if len(right) in (1, 2):
                s = left + "." + right
            elif len(right) == 3 and len(left) > 3:
                s = left + right
            else:
                s = left + "." + right
        else:
            parts = s.split(".")
            if len(parts[-1]) in (1, 2):
                s = "".join(parts[:-1]) + "." + parts[-1]
            else:
                s = "".join(parts)

    try:
        val = float(s)
        return -val if neg else val
    except Exception:
        return np.nan

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

def origem_visual_text(origem):
    if str(origem) == "Somente Financeiro":
        return "● Financeiro"
    if str(origem) == "Somente Contábil":
        return "● Contábil"
    return str(origem)

def get_nucleo_display_series(df):
    nuc_final = df.get("NUCLEO", pd.Series(["Não identificado"] * len(df), index=df.index)).fillna("Não identificado").astype(str).str.strip()
    nuc_sug = df.get("NUCLEO_SUGERIDO", pd.Series(["Não identificado"] * len(df), index=df.index)).fillna("Não identificado").astype(str).str.strip()
    confirmado = df.get("CONFIRMADO", pd.Series([False] * len(df), index=df.index)).fillna(False)

    out = nuc_final.copy()
    mask_empty = out.isin(["", "Não identificado"])
    out.loc[mask_empty & (~confirmado)] = nuc_sug.loc[mask_empty & (~confirmado)]
    out = out.replace("", "Não identificado").fillna("Não identificado")
    return out

def build_sort_columns(df):
    dfx = df.copy()
    dfx["__RES"] = dfx["RESOLVIDO"].fillna(False) | dfx["STATUS"].astype(str).str.lower().eq("resolvido")
    dfx["__SEV_ORD"] = dfx["SEVERIDADE"].map(SEVERITY_ORDER).fillna(0)
    dfx["__ABS_VAL"] = dfx["VALOR"].map(normalize_money).abs()
    dfx["__DATA_SORT"] = pd.to_datetime(dfx["DATA"], errors="coerce")
    return dfx

# =========================================================
# Núcleos persistentes
# =========================================================
def default_nucleos_payload():
    return {"nucleos": DEFAULT_NUCLEOS}

def save_nucleos(payload):
    with open(NUCLEOS_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

def load_nucleos():
    if not os.path.exists(NUCLEOS_FILE):
        payload = default_nucleos_payload()
        save_nucleos(payload)
        return payload["nucleos"]
    try:
        with open(NUCLEOS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        nucs = data.get("nucleos", DEFAULT_NUCLEOS)
        nucs = [str(x).strip() for x in nucs if str(x).strip()]
        for d in DEFAULT_NUCLEOS:
            if d not in nucs:
                nucs.append(d)
        nucs = ["Configuração ERP" if x == "Configuração RP" else x for x in nucs]
        nucs = [x for x in nucs if x != "Não identificado"] + ["Não identificado"]
        save_nucleos({"nucleos": nucs})
        return nucs
    except Exception:
        payload = default_nucleos_payload()
        save_nucleos(payload)
        return payload["nucleos"]

def get_nucleos():
    return load_nucleos()

def add_nucleo(nome):
    nome = str(nome).strip()
    if not nome:
        return False, "Informe o nome do núcleo."
    nucs = get_nucleos()
    if nome in nucs:
        return False, "Esse núcleo já existe."
    nucs = [x for x in nucs if x != "Não identificado"] + [nome, "Não identificado"]
    save_nucleos({"nucleos": nucs})
    return True, f'Núcleo "{nome}" criado com sucesso.'

def rename_nucleo(old_name, new_name):
    old_name = str(old_name).strip()
    new_name = str(new_name).strip()
    if old_name not in get_nucleos():
        return False, "Núcleo de origem não encontrado."
    if not new_name:
        return False, "Informe o novo nome do núcleo."
    if new_name in get_nucleos() and new_name != old_name:
        return False, "Já existe um núcleo com esse nome."
    if old_name == "Não identificado":
        return False, 'O núcleo "Não identificado" não pode ser renomeado.'

    nucs = get_nucleos()
    nucs = [new_name if x == old_name else x for x in nucs]
    save_nucleos({"nucleos": nucs})

    payload = load_rules()
    for bucket in ["nucleo_rules"]:
        for r in payload[bucket]:
            if str(r.get("resultado", "")).strip() == old_name:
                r["resultado"] = new_name
    save_rules(payload)

    learning = load_learning()
    for ex in learning["examples"]:
        if str(ex.get("nucleo_sugerido", "")).strip() == old_name:
            ex["nucleo_sugerido"] = new_name
        if str(ex.get("nucleo_final", "")).strip() == old_name:
            ex["nucleo_final"] = new_name
    save_learning(learning)

    if st.session_state.get("div_master") is not None:
        dm = st.session_state.div_master.copy()
        for c in ["NUCLEO", "NUCLEO_SUGERIDO"]:
            if c in dm.columns:
                dm[c] = dm[c].replace({old_name: new_name})
        st.session_state.div_master = dm

    return True, f'Núcleo "{old_name}" renomeado para "{new_name}".'

def delete_nucleo(nome):
    nome = str(nome).strip()
    if nome in DEFAULT_NUCLEOS:
        return False, "Esse núcleo padrão não pode ser excluído."
    nucs = get_nucleos()
    if nome not in nucs:
        return False, "Núcleo não encontrado."

    nucs = [x for x in nucs if x != nome]
    save_nucleos({"nucleos": nucs})

    payload = load_rules()
    for r in payload["nucleo_rules"]:
        if str(r.get("resultado", "")).strip() == nome:
            r["resultado"] = "Não identificado"
    save_rules(payload)

    if st.session_state.get("div_master") is not None:
        dm = st.session_state.div_master.copy()
        for c in ["NUCLEO", "NUCLEO_SUGERIDO"]:
            if c in dm.columns:
                dm[c] = dm[c].replace({nome: "Não identificado"})
        st.session_state.div_master = dm

    return True, f'Núcleo "{nome}" excluído com sucesso.'

# =========================================================
# Persistência de regras
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
        for r in data["nucleo_rules"]:
            if str(r.get("resultado", "")) == "Configuração RP":
                r["resultado"] = "Configuração ERP"
        save_rules(data)
        return data
    except Exception:
        payload = default_rules_payload()
        save_rules(payload)
        return payload

def save_rules(payload):
    with open(RULES_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def default_association_rules_payload():
    return {"profiles": []}

def load_association_rules():
    if not os.path.exists(ASSOCIATION_RULES_FILE):
        payload = default_association_rules_payload()
        save_association_rules(payload)
        return payload
    try:
        with open(ASSOCIATION_RULES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            data = default_association_rules_payload()
        data.setdefault("profiles", [])
        return data
    except Exception:
        payload = default_association_rules_payload()
        save_association_rules(payload)
        return payload

def save_association_rules(payload):
    with open(ASSOCIATION_RULES_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

def save_association_profile(profile_name, nome_base_a, nome_base_b, specs):
    profile_name = str(profile_name or "").strip()
    if not profile_name or not specs:
        return False, "Informe um nome e confirme ao menos uma regra de associação."
    payload = load_association_rules()
    payload["profiles"] = [p for p in payload.get("profiles", []) if str(p.get("name", "")).strip().lower() != profile_name.lower()]
    payload["profiles"].append({
        "name": profile_name,
        "base_a": nome_base_a,
        "base_b": nome_base_b,
        "specs": specs,
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })
    payload["profiles"] = sorted(payload["profiles"], key=lambda x: str(x.get("name", "")).lower())
    save_association_rules(payload)
    return True, f'Regra de associação "{profile_name}" salva com sucesso.'



def export_association_profile_bytes(profile_name, nome_base_a, nome_base_b, specs):
    payload = {
        "name": str(profile_name or "").strip() or "Regra de associação",
        "base_a": nome_base_a,
        "base_b": nome_base_b,
        "specs": specs or [],
        "exported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")

def import_association_profile(uploaded_file):
    try:
        name = str(getattr(uploaded_file, "name", "") or "")
        suffix = os.path.splitext(name)[1].lower()
        if suffix == ".json":
            raw = json.loads(uploaded_file.getvalue().decode("utf-8"))
            specs = raw.get("specs", []) if isinstance(raw, dict) else []
            return True, {
                "name": raw.get("name", "Regra importada") if isinstance(raw, dict) else "Regra importada",
                "base_a": raw.get("base_a", "") if isinstance(raw, dict) else "",
                "base_b": raw.get("base_b", "") if isinstance(raw, dict) else "",
                "specs": specs,
            }
        if suffix == ".csv":
            df = pd.read_csv(uploaded_file)
            cols = {str(c).strip().lower(): c for c in df.columns}
            src_col = cols.get("source_col_a") or cols.get("campo origem") or cols.get("origem") or cols.get("campo_origem")
            tgt_col = cols.get("target_col_b") or cols.get("campo destino") or cols.get("destino") or cols.get("campo_destino")
            va_col = cols.get("valor_a") or cols.get("rm") or cols.get("valor_rm") or cols.get("source_value")
            vb_col = cols.get("valor_b") or cols.get("protheus") or cols.get("valor_protheus") or cols.get("target_value")
            if not all([src_col, tgt_col, va_col, vb_col]):
                return False, "Arquivo CSV da regra não está no formato esperado."
            specs = []
            for (src_name, tgt_name), grp in df.groupby([src_col, tgt_col], dropna=False):
                mapping = dict(zip(grp[va_col].astype(str), grp[vb_col].astype(str)))
                specs.append({
                    "source_col_a": str(src_name),
                    "target_col_b": str(tgt_name),
                    "map_col_a": f"[MAP] {src_name} -> {tgt_name}",
                    "mapping": mapping,
                })
            return True, {"name": os.path.splitext(name)[0] or "Regra importada", "base_a": "", "base_b": "", "specs": specs}
        return False, "Formato de arquivo não suportado. Use JSON ou CSV."
    except Exception as e:
        return False, f"Erro ao importar regra: {e}"

def next_rule_id(rules):
    used = []
    for r in rules:
        try:
            used.append(int(r.get("id", 0)))
        except Exception:
            pass
    return (max(used) + 1) if used else 1

def normalize_rule_origin(origem):
    origem = str(origem or "").strip()
    if origem == "" or origem.lower() == "qualquer":
        return "Qualquer"
    return origem

def rule_signature(rule):
    parts = [
        normalize_rule_origin(rule.get("origem", "Qualquer")),
        normalize_text_rule(rule.get("texto_contem", "")),
        str(rule.get("regex", "")).strip().lower(),
        str(rule.get("documento_prefixo", "")).strip().upper(),
        str(rule.get("resultado", "")).strip().lower(),
        str(rule.get("valor_min", "")).strip(),
        str(rule.get("valor_max", "")).strip(),
        str(rule.get("nome", "")).strip().lower(),
    ]
    return "||".join(parts)

def add_rule(rule_type, rule_dict):
    payload = load_rules()
    bucket = "nucleo_rules" if rule_type == "nucleo" else "criticidade_rules"

    rule_dict = dict(rule_dict)
    rule_dict["origem"] = normalize_rule_origin(rule_dict.get("origem", "Qualquer"))

    new_sig = rule_signature(rule_dict)
    for r in payload[bucket]:
        if rule_signature(r) == new_sig:
            return False, "Já existe uma regra semelhante cadastrada."

    rule_dict["id"] = next_rule_id(payload[bucket])
    payload[bucket].append(rule_dict)
    payload[bucket] = sorted(payload[bucket], key=lambda x: (int(x.get("prioridade", 9999)), int(x.get("id", 0))))
    save_rules(payload)
    return True, f'Regra "{rule_dict.get("nome", "")}" criada com sucesso.'

def update_rule_status(rule_type, rule_id, active):
    payload = load_rules()
    bucket = "nucleo_rules" if rule_type == "nucleo" else "criticidade_rules"
    found = False
    for r in payload[bucket]:
        if int(r.get("id", 0)) == int(rule_id):
            r["ativa"] = bool(active)
            found = True
            break
    if found:
        save_rules(payload)
        return True
    return False

def delete_rule(rule_type, rule_id):
    payload = load_rules()
    bucket = "nucleo_rules" if rule_type == "nucleo" else "criticidade_rules"
    old_len = len(payload[bucket])
    payload[bucket] = [r for r in payload[bucket] if int(r.get("id", 0)) != int(rule_id)]
    if len(payload[bucket]) != old_len:
        save_rules(payload)
        return True
    return False

# =========================================================
# Persistência de aprendizado
# =========================================================
def default_learning_payload():
    return {"examples": []}

def load_learning():
    if not os.path.exists(LEARNING_FILE):
        payload = default_learning_payload()
        save_learning(payload)
        return payload
    try:
        with open(LEARNING_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            data = default_learning_payload()
        data.setdefault("examples", [])
        for ex in data["examples"]:
            if str(ex.get("nucleo_sugerido", "")) == "Configuração RP":
                ex["nucleo_sugerido"] = "Configuração ERP"
            if str(ex.get("nucleo_final", "")) == "Configuração RP":
                ex["nucleo_final"] = "Configuração ERP"
        save_learning(data)
        return data
    except Exception:
        payload = default_learning_payload()
        save_learning(payload)
        return payload

def save_learning(payload):
    with open(LEARNING_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

def learning_signature(row):
    base = "||".join([
        str(row.get("ORIGEM", "")),
        str(row.get("MOTIVO_BASE", "")),
        str(row.get("NUCLEO", "")),
        str(row.get("SEVERIDADE", "")),
        str(row.get("HISTORICO_OPERACAO", ""))[:120],
    ])
    return hashlib.md5(base.encode("utf-8")).hexdigest()

def save_learning_examples(df):
    payload = load_learning()
    existing = {x.get("sig") for x in payload["examples"]}

    to_save = []
    for _, r in df.iterrows():
        nuc_final = str(r.get("NUCLEO", "")).strip()
        confirmado = bool(r.get("CONFIRMADO", False))
        if not confirmado:
            continue
        if nuc_final == "" or nuc_final == "Não identificado":
            continue

        sig = learning_signature(r)
        if sig in existing:
            continue

        to_save.append({
            "sig": sig,
            "origem": str(r.get("ORIGEM", "")),
            "motivo_base": str(r.get("MOTIVO_BASE", "")),
            "historico_operacao": str(r.get("HISTORICO_OPERACAO", "")),
            "documento": str(r.get("DOCUMENTO", "")),
            "valor": safe_float(r.get("VALOR", np.nan), 0.0),
            "nucleo_sugerido": str(r.get("NUCLEO_SUGERIDO", "")),
            "nucleo_final": nuc_final,
            "severidade_final": str(r.get("SEVERIDADE", "")),
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })

    if to_save:
        payload["examples"].extend(to_save)
        save_learning(payload)

def build_learning_suggestions(div_master):
    learning = load_learning().get("examples", [])
    rows = []

    for ex in learning:
        motivo = str(ex.get("motivo_base", "")).strip()
        origem = str(ex.get("origem", "")).strip()
        nucleo_final = str(ex.get("nucleo_final", "")).strip()
        if motivo and nucleo_final and nucleo_final != "Não identificado":
            rows.append({
                "ORIGEM": origem,
                "MOTIVO_BASE": motivo,
                "NUCLEO_FINAL": nucleo_final,
                "VALOR": safe_float(ex.get("valor", 0.0), 0.0),
            })

    df = div_master.copy()
    if len(df):
        cand = df.copy()
        cand["NUCLEO"] = cand["NUCLEO"].fillna("Não identificado")
        cand["CONFIRMADO"] = cand["CONFIRMADO"].fillna(False)
        cand = cand[cand["CONFIRMADO"]].copy()
        cand = cand[cand["NUCLEO"].ne("Não identificado")].copy()
        cand = cand[cand["MOTIVO_BASE"].astype(str).str.strip().ne("")].copy()

        for _, r in cand.iterrows():
            rows.append({
                "ORIGEM": str(r.get("ORIGEM", "")),
                "MOTIVO_BASE": str(r.get("MOTIVO_BASE", "")),
                "NUCLEO_FINAL": str(r.get("NUCLEO", "")),
                "VALOR": abs(safe_float(r.get("VALOR", 0.0), 0.0)),
            })

    if not rows:
        return pd.DataFrame()

    sug = pd.DataFrame(rows)
    sug["ABS_VALOR"] = sug["VALOR"].abs()

    out = (
        sug.groupby(["ORIGEM", "MOTIVO_BASE", "NUCLEO_FINAL"], dropna=False)
        .agg(
            Qtd=("MOTIVO_BASE", "size"),
            Impacto=("VALOR", "sum"),
            Maior_Valor=("ABS_VALOR", "max")
        )
        .reset_index()
    )

    payload = load_rules()
    current_rules = payload.get("nucleo_rules", [])

    def exists_rule(row):
        motivo = str(row["MOTIVO_BASE"]).strip()
        origem = str(row["ORIGEM"]).strip()
        resultado = str(row["NUCLEO_FINAL"]).strip()

        for rr in current_rules:
            rr_origem = normalize_rule_origin(rr.get("origem", "Qualquer"))
            rr_texto = normalize_text_rule(rr.get("texto_contem", ""))
            rr_result = str(rr.get("resultado", "")).strip()

            origem_ok = rr_origem == "Qualquer" or rr_origem == origem
            texto_ok = rr_texto == normalize_text_rule(motivo)
            resultado_ok = rr_result == resultado
            if origem_ok and texto_ok and resultado_ok:
                return True
        return False

    out["JA_EXISTE"] = out.apply(exists_rule, axis=1)
    out = out[~out["JA_EXISTE"]].copy()
    out = out.sort_values(["Qtd", "Maior_Valor"], ascending=[False, False])
    return out

# =========================================================
# Leitura / motor base
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

def read_table(uploaded):
    name = uploaded.name.lower()

    # CSV
    if name.endswith(".csv"):
        uploaded.seek(0)
        try:
            return pd.read_csv(
                uploaded,
                sep=None,
                engine="python",
                dtype=str,
                keep_default_na=False
            ).fillna("")
        except:
            uploaded.seek(0)
            return pd.read_csv(
                uploaded,
                dtype=str,
                keep_default_na=False
            ).fillna("")

    # EXCEL
    xl = pd.ExcelFile(uploaded)
    best = None

    for sh in xl.sheet_names:
        tmp = xl.parse(sh, dtype=str).fillna("")
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

    if any(k in hist for k in ["erp", "rp", "reprocess", "rotina", "processamento", "integracao", "integração"]):
        return "Configuração ERP"

    return "Não identificado"

def severidade_base(valor):
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
# Motivo base / regras
# =========================================================
STOPWORDS_MOTIVO = {
    "de", "da", "do", "das", "dos", "para", "por", "com", "sem", "na", "no",
    "em", "a", "o", "e", "ou", "um", "uma", "ao", "aos", "as", "os"
}

def build_motivo_base(text):
    txt = normalize_text_rule(text)
    toks = [t for t in txt.split() if t not in STOPWORDS_MOTIVO]
    toks = toks[:8]
    return " ".join(toks).strip()

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

    rule_origem = normalize_rule_origin(rule.get("origem", "Qualquer"))
    if rule_origem not in ["", "Qualquer"] and rule_origem != origem:
        return False

    texto_contem = str(rule.get("texto_contem", "")).strip()
    if texto_contem:
        tnorm = normalize_text_rule(texto_contem)
        if tnorm not in hist_norm and tnorm not in doc_norm:
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
            return rule.get("resultado", default_value), f'Regra #{rule.get("id")} - {rule.get("nome", "")}'
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

        fmt_title = wb.add_format({
            "bold": True, "font_size": 16, "font_color": "#0F172A",
            "align": "left", "valign": "vcenter"
        })
        fmt_subtitle = wb.add_format({
            "font_size": 10, "font_color": "#475569", "italic": True
        })
        fmt_label = wb.add_format({
            "bold": True, "font_size": 10, "font_color": "#334155"
        })
        fmt_info = wb.add_format({
            "font_size": 10, "font_color": "#334155"
        })
        fmt_hdr = wb.add_format({
            "bold": True, "border": 1, "align": "center", "valign": "vcenter",
            "bg_color": "#DBEAFE", "font_color": "#0F172A"
        })
        fmt_txt = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})
        fmt_date = wb.add_format({"num_format": "dd/mm/yyyy", "border": 1})
        fmt_money = wb.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00', "border": 1})
        fmt_money_big = wb.add_format({
            "num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00',
            "bold": True, "font_size": 12, "font_color": "#0F172A"
        })
        fmt_section = wb.add_format({
            "bold": True, "font_size": 12, "font_color": "white",
            "bg_color": "#1E293B", "align": "left", "valign": "vcenter"
        })
        fmt_metric = wb.add_format({"border": 1, "font_color": "#0F172A"})
        fmt_metric_money = wb.add_format({"border": 1, "num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00'})
        fmt_kpi_box = wb.add_format({
            "bold": True, "font_size": 10, "font_color": "#334155",
            "bg_color": "#F8FAFC", "border": 1, "text_wrap": True, "align": "center", "valign": "vcenter"
        })
        fmt_kpi_value = wb.add_format({
            "bold": True, "font_size": 14, "font_color": "#0F172A",
            "bg_color": "#EFF6FF", "border": 1, "align": "center", "valign": "vcenter",
            "num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00'
        })
        fmt_kpi_value_int = wb.add_format({
            "bold": True, "font_size": 14, "font_color": "#0F172A",
            "bg_color": "#EFF6FF", "border": 1, "align": "center", "valign": "vcenter"
        })
        fmt_obs = wb.add_format({
            "font_size": 9, "font_color": "#64748B", "italic": True
        })

        # -------------------------------------------------
        # Aba Divergências
        # -------------------------------------------------
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

        ws.write(0, 0, "ConciliaMais — Divergências Tratadas", fmt_title)
        ws.write(1, 0, "Base filtrada para análise e tratativa", fmt_subtitle)

        ws.write(3, 0, "Processado em:", fmt_label)
        ws.write(3, 1, generated_at, fmt_info)
        ws.write(4, 0, "Origem:", fmt_label)
        ws.write(4, 1, filtros.get("origem", "Todas"), fmt_info)
        ws.write(5, 0, "Visualização:", fmt_label)
        ws.write(5, 1, filtros.get("ver", "Todas"), fmt_info)
        ws.write(6, 0, "Severidade:", fmt_label)
        ws.write(6, 1, filtros.get("severidade", "Todas"), fmt_info)
        ws.write(7, 0, "Núcleo:", fmt_label)
        ws.write(7, 1, filtros.get("nucleo", "Todos"), fmt_info)
        ws.write(8, 0, "Status:", fmt_label)
        ws.write(8, 1, filtros.get("status", "Todos"), fmt_info)
        ws.write(9, 0, "Busca:", fmt_label)
        ws.write(9, 1, filtros.get("busca", ""), fmt_info)

        total_aberto = float(filtros.get("_total_aberto", 0.0) or 0.0)
        total_filtrado = float(df["VALOR"].sum()) if ("VALOR" in df.columns and len(df)) else 0.0

        ws.write(3, 6, "Total do filtro:", fmt_label)
        ws.write_number(3, 7, total_filtrado, fmt_money_big)
        ws.write(4, 6, "Total em aberto:", fmt_label)
        ws.write_number(4, 7, total_aberto, fmt_money_big)
        ws.write(5, 6, "Itens do filtro:", fmt_label)
        ws.write_number(5, 7, len(df), fmt_info)

        start_row_table = 12
        start_col_table = 0

        df2 = df.copy().reset_index(drop=True)
        df2.insert(0, "ID", np.arange(1, len(df2) + 1))

        for j, col in enumerate(df2.columns):
            ws.write(start_row_table, start_col_table + j, col, fmt_hdr)
        ws.set_row(start_row_table, 24)

        for r in range(len(df2)):
            excel_r = start_row_table + 1 + r
            ws.set_row(excel_r, 34)
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

        # -------------------------------------------------
        # Aba Resumo
        # -------------------------------------------------
        wsr = wb.add_worksheet("Resumo")
        w.sheets["Resumo"] = wsr
        wsr.set_zoom(95)
        wsr.set_column(0, 0, 42)
        wsr.set_column(1, 1, 20)
        wsr.set_column(3, 3, 28)
        wsr.set_column(4, 4, 20)
        wsr.set_column(6, 11, 18)

        wsr.merge_range("A1:F1", "ConciliaMais — Resumo Executivo", fmt_title)
        wsr.merge_range("A2:F2", f"Gerado em {generated_at}", fmt_subtitle)

        df_res = df_filtrado.copy()
        if "VALOR" in df_res.columns:
            df_res["VALOR"] = df_res["VALOR"].map(normalize_money)

        total_itens = len(df_res)
        total_valor = float(df_res["VALOR"].sum()) if total_itens else 0.0
        itens_res = int((df_res["RESOLVIDO"].astype(str).str.lower().eq("sim")).sum()) if total_itens and "RESOLVIDO" in df_res.columns else 0
        itens_ab = max(total_itens - itens_res, 0)
        pct_res = (itens_res / total_itens * 100.0) if total_itens else 0.0

        wsr.merge_range("A4:B4", "Visão geral", fmt_section)
        wsr.write("A5", "Itens do filtro", fmt_kpi_box)
        wsr.write_number("B5", total_itens, fmt_kpi_value_int)

        wsr.write("A6", "Valor total do filtro", fmt_kpi_box)
        wsr.write_number("B6", total_valor, fmt_kpi_value)

        wsr.write("D5", "Itens resolvidos", fmt_kpi_box)
        wsr.write_number("E5", itens_res, fmt_kpi_value_int)

        wsr.write("D6", "Itens em aberto", fmt_kpi_box)
        wsr.write_number("E6", itens_ab, fmt_kpi_value_int)

        wsr.write("G5", "Conferência do cálculo", fmt_kpi_box)
        if pd.notna(stats.get("conferencia", np.nan)):
            wsr.write_number("H5", float(stats.get("conferencia", np.nan)), fmt_kpi_value)
        else:
            wsr.write("H5", "-", fmt_kpi_value_int)

        wsr.write("G6", "% resolvido", fmt_kpi_box)
        wsr.write("H6", f"{pct_res:.1f}%", fmt_kpi_value_int)

        wsr.merge_range("A9:B9", "Composição de saldos", fmt_section)

        resumo_comp = [
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
        ]

        base_row = 10
        for i, (metrica, valor) in enumerate(resumo_comp):
            row = base_row + i
            if metrica == "":
                wsr.write_blank(row, 0, None)
                wsr.write_blank(row, 1, None)
            else:
                wsr.write(row, 0, metrica, fmt_metric)
                if pd.notna(valor):
                    wsr.write_number(row, 1, float(valor), fmt_metric_money)
                else:
                    wsr.write(row, 1, "-", fmt_metric)

        # Distribuição por origem
        wsr.merge_range("D9:E9", "Distribuição por origem", fmt_section)
        dist_origem = pd.DataFrame()
        if len(df_res) and "ORIGEM" in df_res.columns:
            dist_origem = (
                df_res.groupby("ORIGEM", dropna=False)
                .agg(Itens=("VALOR", "size"), Valor=("VALOR", "sum"))
                .reset_index()
                .sort_values("Valor", ascending=False)
            )

        hdr_simple = wb.add_format({
            "bold": True, "border": 1, "bg_color": "#E2E8F0", "font_color": "#0F172A", "align": "center"
        })

        start_ro = 10
        for j, col in enumerate(["Origem", "Itens", "Valor"]):
            wsr.write(start_ro, 3 + j, col, hdr_simple)

        if not dist_origem.empty:
            for i, (_, rr) in enumerate(dist_origem.iterrows(), start=1):
                wsr.write(start_ro + i, 3, rr["ORIGEM"], fmt_metric)
                wsr.write_number(start_ro + i, 4, int(rr["Itens"]), fmt_metric)
                wsr.write_number(start_ro + i, 5, float(rr["Valor"]), fmt_metric_money)

        # Top 10
        wsr.merge_range("G9:L9", "Top 10 pendências mais impactantes", fmt_section)
        top_df = pd.DataFrame()
        if len(df_res):
            dft = df_res.copy()
            dft["VALOR"] = dft["VALOR"].map(normalize_money)
            dft["ABS"] = dft["VALOR"].abs()
            mask_open = True
            if "RESOLVIDO" in dft.columns:
                mask_open = ~dft["RESOLVIDO"].astype(str).str.lower().eq("sim")
            top_df = dft.loc[mask_open].sort_values("ABS", ascending=False).head(10).copy()

        top_headers = ["Origem", "Data", "Documento", "Valor", "Núcleo", "Status"]
        for j, col in enumerate(top_headers):
            wsr.write(10, 6 + j, col, hdr_simple)

        if not top_df.empty:
            top_df = top_df.reset_index(drop=True)
            for i, (_, rr) in enumerate(top_df.iterrows(), start=1):
                dt = pd.to_datetime(rr.get("DATA"), errors="coerce")
                wsr.write(10 + i, 6, rr.get("ORIGEM", ""), fmt_metric)
                if pd.notna(dt):
                    wsr.write_datetime(10 + i, 7, dt.to_pydatetime(), fmt_date)
                else:
                    wsr.write(10 + i, 7, "", fmt_metric)
                wsr.write(10 + i, 8, str(rr.get("DOCUMENTO", "")), fmt_metric)
                if pd.notna(rr.get("VALOR", np.nan)):
                    wsr.write_number(10 + i, 9, float(rr.get("VALOR", 0.0)), fmt_metric_money)
                else:
                    wsr.write(10 + i, 9, "-", fmt_metric)
                wsr.write(10 + i, 10, str(rr.get("NUCLEO", "Não identificado")), fmt_metric)
                wsr.write(10 + i, 11, str(rr.get("STATUS", "")), fmt_metric)

        wsr.write("A25", "Observação:", fmt_label)
        wsr.write("B25", "A aba Divergências respeita exatamente o filtro aplicado na tela no momento da exportação.", fmt_obs)
        wsr.freeze_panes(4, 0)

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

get_nucleos()
load_rules()
load_learning()
# =========================================================
# Match Inteligente V3
# =========================================================
def render_cruzamento_inteligente_v2():
    st.title("Match Inteligente")
    st.caption("Cruze duas bases grandes com foco em performance, preservação dos valores originais e comparação em uma única tabela.")

    def _force_text_series(sr):
        return sr.fillna("").map(lambda x: "" if pd.isna(x) else str(x))

    def _norm_text(x):
        if pd.isna(x):
            return ""
        s = str(x).strip()
        s = re.sub(r"\s+", " ", s)
        return s

    def _norm_name(x):
        s = _norm_text(x).lower()
        s = unicodedata.normalize("NFD", s)
        s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
        s = re.sub(r"[^a-z0-9]+", " ", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    def _extract_numeric_str(x):
        s = "" if pd.isna(x) else str(x)
        return re.sub(r"\D", "", s)

    def _build_key(df, cols):
        if not cols:
            return pd.Series([""] * len(df), index=df.index)
        key = _force_text_series(df[cols[0]])
        for c in cols[1:]:
            key = key + "||" + _force_text_series(df[c])
        return key

    def _name_similarity(a, b):
        na = _norm_name(a)
        nb = _norm_name(b)
        if na == nb:
            return 1.0
        ta = set(na.split())
        tb = set(nb.split())
        if not ta or not tb:
            return 0.0
        base = len(ta.intersection(tb)) / max(len(ta.union(tb)), 1)
        grupos = [
            {"filial", "loja"},
            {"codigo", "cod", "id"},
            {"patrimonio", "patrimônio", "plaqueta", "num", "numero", "número"},
            {"aquisicao", "aquisição", "orig", "original", "valor", "val"},
            {"depreciacao", "depreciação", "depr", "acum", "mensal"},
            {"saldo"},
            {"nome", "historico", "hist", "descricao", "descrição", "grupo"},
            {"conta"},
        ]
        bonus = 0.0
        for g in grupos:
            if ta.intersection(g) and tb.intersection(g):
                bonus += 0.25
        return min(base + bonus, 1.0)

    def _content_similarity(df_a, col_a, df_b, col_b):
        sa = _force_text_series(df_a[col_a]).head(3000)
        sb = _force_text_series(df_b[col_b]).head(3000)
        set_a = set([x for x in sa if x != ""])
        set_b = set([x for x in sb if x != ""])
        if not set_a or not set_b:
            return 0.0
        raw_inter = len(set_a.intersection(set_b))
        raw_base = max(1, min(len(set_a), len(set_b)))
        raw_score = raw_inter / raw_base
        dig_a = set([_extract_numeric_str(x) for x in set_a if _extract_numeric_str(x) != ""])
        dig_b = set([_extract_numeric_str(x) for x in set_b if _extract_numeric_str(x) != ""])
        dig_inter = len(dig_a.intersection(dig_b))
        dig_base = max(1, min(len(dig_a), len(dig_b))) if dig_a and dig_b else 1
        dig_score = dig_inter / dig_base if dig_a and dig_b else 0.0
        return (raw_score * 0.45) + (dig_score * 0.55)

    def _suggest_pairs(df_a, df_b, top_n=10):
        rows = []
        for ca in df_a.columns:
            for cb in df_b.columns:
                nscore = _name_similarity(ca, cb)
                cscore = _content_similarity(df_a, ca, df_b, cb)
                score = (nscore * 0.35) + (cscore * 0.65)
                if score >= 0.18:
                    rows.append({
                        "USAR": True if score >= 0.45 else False,
                        "ORDEM": 99,
                        "CAMPO_BASE_A": ca,
                        "CAMPO_BASE_B": cb,
                        "CONFIANCA": "Alta" if score >= 0.75 else "Média" if score >= 0.45 else "Baixa",
                        "SCORE": round(score * 100, 1),
                    })
        if not rows:
            return pd.DataFrame()
        sug = pd.DataFrame(rows).sort_values(["SCORE", "CAMPO_BASE_A", "CAMPO_BASE_B"], ascending=[False, True, True]).reset_index(drop=True)
        usados_a, usados_b, escolhidos, ordem = set(), set(), [], 1
        for _, r in sug.iterrows():
            a = r["CAMPO_BASE_A"]
            b = r["CAMPO_BASE_B"]
            if a not in usados_a and b not in usados_b:
                rr = r.copy()
                rr["ORDEM"] = ordem
                escolhidos.append(rr)
                usados_a.add(a)
                usados_b.add(b)
                ordem += 1
            if len(escolhidos) >= top_n:
                break
        return pd.DataFrame(escolhidos)

    def _round_money_series(sr):
        nums = sr.map(normalize_money)
        return nums.round(2)

    def _compare_series(sa_raw, sb_raw, mode="Numérico", tol=0.01):
        sa_raw = _force_text_series(sa_raw)
        sb_raw = _force_text_series(sb_raw)
        if mode == "Numérico":
            n1 = _round_money_series(sa_raw)
            n2 = _round_money_series(sb_raw)
            diff = (n1 - n2).round(2)
            both = n1.notna() & n2.notna()
            status = pd.Series("Sem comparação", index=sa_raw.index, dtype="object")
            status.loc[both & (diff.abs() <= float(tol))] = "Coerente"
            status.loc[both & (diff.abs() > float(tol))] = "Divergente"
            return n1, n2, diff.where(both, np.nan), status
        if mode == "Texto exato":
            status = pd.Series("Divergente", index=sa_raw.index, dtype="object")
            status.loc[sa_raw.eq(sb_raw)] = "Coerente"
            return sa_raw, sb_raw, pd.Series(np.nan, index=sa_raw.index), status
        na = sa_raw.map(_norm_name)
        nb = sb_raw.map(_norm_name)
        status = pd.Series("Divergente", index=sa_raw.index, dtype="object")
        status.loc[na.eq(nb)] = "Coerente"
        return sa_raw, sb_raw, pd.Series(np.nan, index=sa_raw.index), status

    def _excel_col_type(col_name):
        up = str(col_name).upper()
        if "DIF_" in up or "VALOR_" in up or "SOMA_" in up:
            return "currency"
        if any(x in up for x in ["CHAVE", "COD", "CÓD", "FILIAL", "LOJA", "PATRIM", "PLAQUETA", "STATUS", "CENARIO", "RESULTADO"]):
            return "text"
        return "text"

    def _write_df_excel(ws, df, wb, text_priority_cols=None):
        if df is None or df.empty:
            return
        text_priority_cols = set(text_priority_cols or [])
        fmt_hdr = wb.add_format({"bold": True, "bg_color": "#DBEAFE", "border": 1, "align": "center", "valign": "vcenter"})
        fmt_text = wb.add_format({"border": 1, "num_format": "@"})
        fmt_num = wb.add_format({"border": 1, "num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00'})
        for c, col in enumerate(df.columns):
            ws.write(0, c, col, fmt_hdr)
        for r in range(len(df)):
            for c, col in enumerate(df.columns):
                val = df.iloc[r, c]
                if pd.isna(val):
                    ws.write_blank(r + 1, c, None, fmt_text)
                    continue
                col_up = str(col).upper()
                if col in text_priority_cols or _excel_col_type(col) == "text":
                    ws.write_string(r + 1, c, str(val), fmt_text)
                elif isinstance(val, (int, float, np.integer, np.floating)):
                    ws.write_number(r + 1, c, float(round(val, 2)), fmt_num)
                else:
                    ws.write_string(r + 1, c, str(val), fmt_text)
        for c, col in enumerate(df.columns):
            sample = [str(col)] + df[col].astype(str).head(200).tolist()
            width = min(max(max(len(x) for x in sample) + 2, 12), 42)
            ws.set_column(c, c, width)

    def _label_exec(raw):
        txt = str(raw)
        txt = re.sub(r"__SRC\d+$", "", txt)
        txt = txt.replace("TOTALIZADOR__", "")
        txt = txt.replace("TOTALIZADOR_", "")
        txt = txt.replace(" :: ", " - ")
        txt = txt.replace("RM - ", "")
        txt = txt.replace("Protheus - ", "")
        txt = txt.replace("RM__", "")
        txt = txt.replace("Protheus__", "")
        txt = txt.replace("__", " - ")
        txt = txt.replace("_", " ")
        txt = re.sub(r"\s+", " ", txt).strip(" -")
        return txt

    def _norm_field_label(txt):
        txt = str(txt or "")
        txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
        txt = re.sub(r"[^A-Za-z0-9]+", " ", txt).strip().lower()
        return txt

    def _canon_totalizer_label(col_a, col_b=None):
        a = str(col_a or "").strip()
        b = str(col_b or "").strip()
        if a and b:
            na, nb = _norm_field_label(a), _norm_field_label(b)
            if na == nb:
                return a
            if na and nb and (na in nb or nb in na):
                return a if len(a) <= len(b) else b
            return a
        return a or b

    def _build_totalizador(df_result, compare_meta, totalizador_cols):
        if df_result is None or df_result.empty:
            return pd.DataFrame(), []
        work = df_result.copy()
        group_cols = [c for c in totalizador_cols if c in work.columns]
        if not group_cols:
            group_cols = [c for c in work.columns if c.startswith("TOTALIZADOR__")][:3]
        if not group_cols:
            return pd.DataFrame(), []

        selected_meta = [m for m in compare_meta if m.get("selected", True)]
        if not selected_meta:
            selected_meta = compare_meta

        # KPIs financeiros sempre fecham com a base inteira, não só com o subconjunto agrupado.
        financial_kpis = []
        for meta in selected_meta:
            label = meta.get("label") or _label_exec(meta.get("diff_col"))
            va_all = pd.to_numeric(work.get(meta.get("val_a_col")), errors="coerce").fillna(0).round(2).sum()
            vb_all = pd.to_numeric(work.get(meta.get("val_b_col")), errors="coerce").fillna(0).round(2).sum()
            vd_all = round(float(va_all - vb_all), 2)
            financial_kpis.extend([
                (f"{label} {nome_base_a}", round(float(va_all), 2)),
                (f"{label} {nome_base_b}", round(float(vb_all), 2)),
                (f"Diferença {label}", vd_all),
            ])

        rename_map = {c: _label_exec(c) for c in group_cols}
        work = work.rename(columns=rename_map)
        group_labels = [rename_map[c] for c in group_cols]

        # Garante colunas agrupadoras como texto para evitar falhas com 1 único agrupador.
        for c in group_labels:
            work[c] = _force_text_series(work[c]).fillna("")

        rows = []
        grouped = work.groupby(group_labels, dropna=False, sort=True)
        for keys, grp in grouped:
            if len(group_labels) == 1:
                keys = (keys,)
            row = {group_labels[i]: keys[i] for i in range(len(group_labels))}
            row["Qtde Registros"] = int(len(grp))
            row["Qtde Divergências"] = int((grp["RESULTADO_FINAL"].astype(str) == "Match com divergência").sum()) if "RESULTADO_FINAL" in grp.columns else 0
            row["Qtde Ausentes"] = int(grp["CENARIO_MATCH"].astype(str).str.contains("Não encontrado", na=False).sum()) if "CENARIO_MATCH" in grp.columns else 0
            row["Qtde Duplicidades"] = int(grp["CENARIO_MATCH"].astype(str).str.contains("Duplicidade", na=False).sum()) if "CENARIO_MATCH" in grp.columns else 0
            for meta in selected_meta:
                label = meta.get("label") or _label_exec(meta.get("diff_col"))
                va = pd.to_numeric(grp.get(meta.get("val_a_col")), errors="coerce").fillna(0).round(2).sum()
                vb = pd.to_numeric(grp.get(meta.get("val_b_col")), errors="coerce").fillna(0).round(2).sum()
                vd = round(float(va - vb), 2)
                row[f"{label} {nome_base_a}"] = round(float(va), 2)
                row[f"{label} {nome_base_b}"] = round(float(vb), 2)
                row[f"Diferença {label}"] = vd
            rows.append(row)
        agg = pd.DataFrame(rows)
        support_cols = [c for c in ["Qtde Divergências", "Qtde Ausentes", "Qtde Duplicidades"] if c in agg.columns]
        diff_cols = [c for c in agg.columns if str(c).startswith("Diferença ")]
        numeric_focus = diff_cols + support_cols
        if numeric_focus and not agg.empty:
            mask_keep = agg[numeric_focus].apply(pd.to_numeric, errors="coerce").fillna(0).abs().sum(axis=1) > 0
            if mask_keep.any():
                agg = agg[mask_keep].copy()
        if not agg.empty:
            agg = agg.sort_values(group_labels).reset_index(drop=True)
        return agg, financial_kpis


    
    def _to_excel_package(df_result, resumo_dict, text_priority_cols=None, totalizador_df=None, financial_kpis=None):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            wb = writer.book
            fmt_hdr = wb.add_format({"bold": True, "bg_color": "#DBEAFE", "border": 1, "align": "center", "valign": "vcenter"})
            fmt_label = wb.add_format({"bold": True, "border": 1, "bg_color": "#F8FAFC"})
            fmt_value = wb.add_format({"border": 1})
            fmt_kpi = wb.add_format({"bold": True, "font_size": 14, "bg_color": "#EFF6FF", "border": 1, "align": "center", "valign": "vcenter"})
            fmt_num = wb.add_format({"border": 1, "num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00'})
            fmt_total = wb.add_format({"bold": True, "bg_color": "#E0F2FE", "border": 1, "num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00'})
            ws = wb.add_worksheet("RESUMO_EXECUTIVO")
            writer.sheets["RESUMO_EXECUTIVO"] = ws

            resumo_rows = [
                ("Tipo de análise", resumo_dict.get("objetivo", "")),
                ("Modo", resumo_dict.get("direcao", "")),
                ("Registros analisados", resumo_dict.get("total", 0)),
                ("Correspondências", resumo_dict.get("encontrados", 0)),
                ("Registros ausentes", resumo_dict.get("nao_encontrados", 0)),
                ("Duplicidades", resumo_dict.get("duplicados", 0)),
                ("Divergências", resumo_dict.get("divergentes", 0)),
                ("Aderência (%)", resumo_dict.get("aderencia", 0.0)),
            ]
            ws.write(0, 0, "Indicador", fmt_hdr)
            ws.write(0, 1, "Valor", fmt_hdr)
            for r, (lab, val) in enumerate(resumo_rows, start=1):
                ws.write(r, 0, lab, fmt_label)
                if isinstance(val, (int, float, np.integer, np.floating)) and "Aderência" not in lab:
                    ws.write_number(r, 1, float(val), fmt_value)
                else:
                    ws.write(r, 1, val, fmt_value)
            ws.set_column(0, 0, 28)
            ws.set_column(1, 1, 26)

            ws.write(1, 3, "Visão Executiva", fmt_hdr)
            ws.write(2, 3, "Registros", fmt_label); ws.write_number(2, 4, float(resumo_dict.get("total", 0)), fmt_kpi)
            ws.write(3, 3, "Divergências", fmt_label); ws.write_number(3, 4, float(resumo_dict.get("divergentes", 0)), fmt_kpi)
            ws.write(4, 3, "Ausentes", fmt_label); ws.write_number(4, 4, float(resumo_dict.get("nao_encontrados", 0)), fmt_kpi)
            ws.write(5, 3, "Duplicidades", fmt_label); ws.write_number(5, 4, float(resumo_dict.get("duplicados", 0)), fmt_kpi)

            row_fin = 1
            if financial_kpis:
                ws.write(0, 6, "Impacto financeiro total", fmt_hdr)
                for titulo, valor in financial_kpis:
                    ws.write(row_fin, 6, titulo, fmt_label)
                    ws.write_number(row_fin, 7, float(round(valor, 2)), fmt_num)
                    row_fin += 1
                ws.set_column(6, 6, 34)
                ws.set_column(7, 7, 18)

            if totalizador_df is not None and not totalizador_df.empty:
                start_row = max(9, row_fin + 2)
                ws.write(start_row, 0, "Resumo executivo por chave", fmt_hdr)
                for c, col in enumerate(totalizador_df.columns):
                    ws.write(start_row + 1, c, col, fmt_hdr)
                    width = min(max(len(str(col)) + 3, 14), 28)
                    ws.set_column(c, c, width)
                for r in range(len(totalizador_df)):
                    for c, col in enumerate(totalizador_df.columns):
                        val = totalizador_df.iloc[r, c]
                        if isinstance(val, (int, float, np.integer, np.floating)) and ("Diferença" in str(col) or nome_base_a in str(col) or nome_base_b in str(col)):
                            ws.write_number(start_row + 2 + r, c, float(round(val, 2)), fmt_num)
                        else:
                            ws.write(start_row + 2 + r, c, val, fmt_value)
                total_row = start_row + 2 + len(totalizador_df)
                ws.write(total_row, 0, "TOTAL GERAL", fmt_hdr)
                for c, col in enumerate(totalizador_df.columns[1:], start=1):
                    if "Diferença" in str(col):
                        val = float(pd.to_numeric(totalizador_df[col], errors="coerce").fillna(0).sum())
                        ws.write_number(total_row, c, round(val, 2), fmt_total)
                    elif nome_base_a in str(col) or nome_base_b in str(col):
                        val = float(pd.to_numeric(totalizador_df[col], errors="coerce").fillna(0).sum())
                        ws.write_number(total_row, c, round(val, 2), fmt_total)
                    elif "Qtde" in str(col):
                        val = float(pd.to_numeric(totalizador_df[col], errors="coerce").fillna(0).sum())
                        ws.write_number(total_row, c, val, fmt_kpi)
            ws_full = wb.add_worksheet("RESULTADO_COMPLETO")
            writer.sheets["RESULTADO_COMPLETO"] = ws_full
            _write_df_excel(ws_full, df_result, wb, text_priority_cols=text_priority_cols)
            if not df_result.empty and "RESULTADO_FINAL" in df_result.columns:
                for nome, filtro in [
                    ("DIVERGENCIAS", df_result["RESULTADO_FINAL"] == "Match com divergência"),
                    ("NAO_ENCONTRADOS", df_result["CENARIO_MATCH"].astype(str).str.contains("Não encontrado", na=False)),
                    ("DUPLICIDADES", df_result["CENARIO_MATCH"].astype(str).str.contains("Duplicidade", na=False)),
                ]:
                    df_tmp = df_result[filtro].copy()
                    ws_tmp = wb.add_worksheet(nome)
                    writer.sheets[nome] = ws_tmp
                    _write_df_excel(ws_tmp, df_tmp, wb, text_priority_cols=text_priority_cols)
        output.seek(0)
        return output


    st.markdown("### 1) Qual análise deseja realizar entre as bases?")
    objetivo_label = st.radio("Tipo de análise", [
        "Comparar valores de registros correspondentes",
        "Encontrar registros faltantes entre as bases",
        "Identificar registros duplicados",
        "Completar informações de uma base com a outra",
    ])

    st.markdown("### 2) Bases da análise")
    c1, c2 = st.columns(2)
    with c1:
        nome_base_a = st.text_input("Nome exibido da Base A", value="Base A", key="v6_nome_base_a")
        st.markdown(f"**{nome_base_a or 'Base A'} — primeira base da auditoria**")
        base_a_file = st.file_uploader("Upload Base A (.xlsx ou .csv)", type=["xlsx", "csv"], key="procx6_a")
    with c2:
        nome_base_b = st.text_input("Nome exibido da Base B", value="Base B", key="v6_nome_base_b")
        st.markdown(f"**{nome_base_b or 'Base B'} — segunda base da auditoria**")
        base_b_file = st.file_uploader("Upload Base B (.xlsx ou .csv)", type=["xlsx", "csv"], key="procx6_b")
    nome_base_a = (nome_base_a or "Base A").strip() or "Base A"
    nome_base_b = (nome_base_b or "Base B").strip() or "Base B"

    if not base_a_file or not base_b_file:
        st.info("Faça o upload das duas bases para continuar.")
        return
    try:
        df_a = read_table(base_a_file)
        df_b = read_table(base_b_file)
    except Exception as e:
        st.error(f"Erro ao ler os arquivos: {e}")
        return

    st.markdown("### 3) Tipo de auditoria")
    direcao = "Auditoria completa"
    st.info("A análise será executada em auditoria completa, usando a chave composta definida e confrontando os valores em uma visão única, sem duplicar os dois sentidos.")

    st.markdown("### 4) Visão inicial das bases")
    v1, v2 = st.columns(2)
    with v1:
        st.markdown(f"**{nome_base_a}**")
        st.caption(f"{len(df_a):,} linhas | {len(df_a.columns)} colunas")
        st.dataframe(df_a.head(8), use_container_width=True, height=220)
    with v2:
        st.markdown(f"**{nome_base_b}**")
        st.caption(f"{len(df_b):,} linhas | {len(df_b.columns)} colunas")
        st.dataframe(df_b.head(8), use_container_width=True, height=220)

    st.markdown("### 5) Regras opcionais de associação entre campos")
    st.caption("Use esta etapa quando quiser criar uma correspondência adicional entre campos das duas bases antes do match principal.")
    assoc_on = st.checkbox("Quero criar uma regra de associação entre campos", value=False, key="v8_assoc_on")
    if "v8_confirmed_assoc_specs" not in st.session_state:
        st.session_state["v8_confirmed_assoc_specs"] = []
    if "v8_assoc_grid_df" not in st.session_state:
        st.session_state["v8_assoc_grid_df"] = pd.DataFrame()
    association_specs = list(st.session_state.get("v8_confirmed_assoc_specs", []))
    if assoc_on:
        saved_payload = load_association_rules()
        profile_names = [p.get("name", "") for p in saved_payload.get("profiles", [])]
        h1, h2 = st.columns([2, 1])
        with h1:
            selected_profile = st.selectbox("Carregar regra salva", [""] + profile_names, key="v8_assoc_profile")
        with h2:
            if st.button("Aplicar regra salva", use_container_width=True, key="v8_assoc_load_btn") and selected_profile:
                prof = next((p for p in saved_payload.get("profiles", []) if p.get("name") == selected_profile), None)
                if prof:
                    st.session_state["v8_confirmed_assoc_specs"] = prof.get("specs", [])
                    association_specs = list(st.session_state["v8_confirmed_assoc_specs"])
                    st.success(f'Regra "{selected_profile}" carregada para esta análise.')

        imp1, imp2 = st.columns([2, 1])
        with imp1:
            uploaded_assoc_rule = st.file_uploader("Importar regra de associação (.json ou .csv)", type=["json", "csv"], key="v11_assoc_import_file")
        with imp2:
            if st.button("Aplicar regra importada", use_container_width=True, key="v11_assoc_import_btn"):
                if uploaded_assoc_rule is None:
                    st.warning("Selecione um arquivo de regra para importar.")
                else:
                    ok_imp, imported_payload = import_association_profile(uploaded_assoc_rule)
                    if ok_imp:
                        imported_specs = imported_payload.get("specs", []) if isinstance(imported_payload, dict) else []
                        st.session_state["v8_confirmed_assoc_specs"] = imported_specs
                        association_specs = list(imported_specs)
                        imported_name = imported_payload.get("name", "Regra importada") if isinstance(imported_payload, dict) else "Regra importada"
                        st.success(f'Regra importada e aplicada: {imported_name}.')
                    else:
                        st.error(imported_payload)

        a1, a2 = st.columns(2)
        with a1:
            assoc_col_a = st.selectbox(f"Campo descritivo da {nome_base_a}", list(df_a.columns), key="v8_assoc_col_a")
        with a2:
            assoc_col_b = st.selectbox(f"Campo de destino da {nome_base_b}", list(df_b.columns), key="v8_assoc_col_b")

        uniq_vals_a = sorted([str(x).strip() for x in _force_text_series(df_a[assoc_col_a]).dropna().unique().tolist() if str(x).strip() != ""])
        uniq_vals_b = sorted([str(x).strip() for x in _force_text_series(df_b[assoc_col_b]).dropna().unique().tolist() if str(x).strip() != ""])
        grid_key = f"v8_assoc_grid::{assoc_col_a}::{assoc_col_b}"
        if "v13_assoc_grid_loaded" not in st.session_state:
            st.session_state["v13_assoc_grid_loaded"] = False
        load1, load2 = st.columns([1, 2])
        with load1:
            if st.button("Carregar associação", key="v13_assoc_load_grid", use_container_width=True):
                current_grid = pd.DataFrame({
                    "USAR": [False] * min(len(uniq_vals_a), 300),
                    f"VALOR_{nome_base_a}": uniq_vals_a[:300],
                    f"VALOR_{nome_base_b}": [""] * min(len(uniq_vals_a), 300),
                })
                # Preenche previamente quando já existir regra confirmada para o mesmo par de campos
                prev = next((s for s in st.session_state.get("v8_confirmed_assoc_specs", []) if s.get("source_col_a") == assoc_col_a and s.get("target_col_b") == assoc_col_b), None)
                if prev:
                    mp = prev.get("mapping", {}) or {}
                    current_grid[f"VALOR_{nome_base_b}"] = [mp.get(v, "") for v in current_grid[f"VALOR_{nome_base_a}"].astype(str)]
                    current_grid["USAR"] = current_grid[f"VALOR_{nome_base_b}"].astype(str).str.strip() != ""
                st.session_state[grid_key] = current_grid
                st.session_state["v13_assoc_grid_loaded"] = True
        with load2:
            st.caption("Escolha os campos e clique em carregar para montar a grade de associação. A grade só aparece depois disso.")

        if st.session_state.get("v13_assoc_grid_loaded") and grid_key in st.session_state:
            b1, b2, b3 = st.columns([1, 1, 2])
            with b1:
                if st.button("Marcar todos", key="v8_assoc_mark_all", use_container_width=True):
                    tmp = st.session_state[grid_key].copy()
                    tmp["USAR"] = True
                    st.session_state[grid_key] = tmp
            with b2:
                if st.button("Desmarcar todos", key="v8_assoc_unmark_all", use_container_width=True):
                    tmp = st.session_state[grid_key].copy()
                    tmp["USAR"] = False
                    st.session_state[grid_key] = tmp
            with b3:
                st.caption("Confirme a regra atual para manter esta correlação e, se quiser, adicionar outras regras na mesma análise.")

            assoc_df = st.data_editor(
                st.session_state[grid_key],
                use_container_width=True,
                height=280,
                hide_index=True,
                column_config={
                    "USAR": st.column_config.CheckboxColumn("Usar"),
                    f"VALOR_{nome_base_a}": st.column_config.TextColumn(nome_base_a, disabled=True),
                    f"VALOR_{nome_base_b}": st.column_config.SelectboxColumn(nome_base_b, options=uniq_vals_b),
                },
                key=f"{grid_key}::editor",
            )
            st.session_state[grid_key] = assoc_df.copy()
        else:
            assoc_df = pd.DataFrame(columns=["USAR", f"VALOR_{nome_base_a}", f"VALOR_{nome_base_b}"])

        c1, c2, c3 = st.columns([1, 1, 2])
        with c1:
            if st.button("Confirmar regra atual", type="primary", use_container_width=True, key="v8_assoc_confirm_btn"):
                assoc_sel = assoc_df[(assoc_df["USAR"] == True) & (assoc_df[f"VALOR_{nome_base_b}"].astype(str).str.strip() != "")].copy()
                if assoc_sel.empty:
                    st.warning("Selecione ao menos uma linha e informe o valor correspondente na outra base.")
                else:
                    map_col_a = f"[MAP] {assoc_col_a} -> {assoc_col_b}"
                    new_spec = {
                        "source_col_a": assoc_col_a,
                        "target_col_b": assoc_col_b,
                        "map_col_a": map_col_a,
                        "mapping": dict(zip(assoc_sel[f"VALOR_{nome_base_a}"].astype(str), assoc_sel[f"VALOR_{nome_base_b}"].astype(str))),
                    }
                    specs = [s for s in st.session_state.get("v8_confirmed_assoc_specs", []) if not (s.get("source_col_a") == assoc_col_a and s.get("target_col_b") == assoc_col_b)]
                    specs.append(new_spec)
                    st.session_state["v8_confirmed_assoc_specs"] = specs
                    association_specs = list(specs)
                    st.success(f"Regra confirmada com {len(new_spec['mapping'])} associação(ões).")
        with c2:
            if st.button("Limpar regras", use_container_width=True, key="v8_assoc_clear_btn"):
                st.session_state["v8_confirmed_assoc_specs"] = []
                association_specs = []
                st.success("Regras de associação da análise atual limpas.")
        with c3:
            save_profile_name = st.text_input("Salvar regra para uso futuro", value="", placeholder="Ex.: RM x Protheus - Grupo Patrimônio x Conta", key="v8_assoc_save_name")
            specs_to_export = st.session_state.get("v8_confirmed_assoc_specs", [])
            export_name = (save_profile_name or "regra_associacao").strip() or "regra_associacao"
            st.download_button(
                "Baixar regra atual",
                data=export_association_profile_bytes(export_name, nome_base_a, nome_base_b, specs_to_export),
                file_name=f"{export_name.replace(' ', '_')}.json",
                mime="application/json",
                use_container_width=True,
                key="v11_assoc_download_btn",
                disabled=(len(specs_to_export) == 0),
            )
            if st.button("Salvar regra", use_container_width=True, key="v8_assoc_save_btn"):
                ok, msg = save_association_profile(save_profile_name, nome_base_a, nome_base_b, st.session_state.get("v8_confirmed_assoc_specs", []))
                (st.success if ok else st.warning)(msg)

        association_specs = list(st.session_state.get("v8_confirmed_assoc_specs", []))
        if association_specs:
            resumo_regras = pd.DataFrame([
                {
                    "Campo origem": s.get("source_col_a", ""),
                    "Campo destino": s.get("target_col_b", ""),
                    "Qtde associações": len(s.get("mapping", {}) or {}),
                }
                for s in association_specs
            ])
            st.markdown("#### Regras já confirmadas nesta análise")
            st.dataframe(resumo_regras, use_container_width=True, height=180, hide_index=True)

    work_a = df_a.copy()
    work_b = df_b.copy()
    for spec in association_specs:
        work_a[spec["map_col_a"]] = _force_text_series(work_a[spec["source_col_a"]]).map(lambda x: spec["mapping"].get(str(x), ""))

    sug_df = _suggest_pairs(work_a, work_b)

    st.markdown("### 6) Quais campos identificam o mesmo registro nas duas bases?")
    st.caption("Selecione os campos que serão usados para localizar o mesmo item nas duas bases.")
    if sug_df.empty:
        key_rows = [{"USAR": True, "ORDEM": 1, "CAMPO_BASE_A": list(work_a.columns)[0], "CAMPO_BASE_B": list(work_b.columns)[0], "CONFIANCA": "Manual", "SCORE": 0.0}]
    else:
        key_rows = sug_df.head(8).copy()
    key_df = st.data_editor(pd.DataFrame(key_rows), use_container_width=True, height=280, hide_index=True,
        column_config={
            "USAR": st.column_config.CheckboxColumn("Usar"),
            "ORDEM": st.column_config.NumberColumn("Ordem", min_value=1, step=1),
            "CAMPO_BASE_A": st.column_config.SelectboxColumn(f"Campo {nome_base_a}", options=list(work_a.columns)),
            "CAMPO_BASE_B": st.column_config.SelectboxColumn(f"Campo {nome_base_b}", options=list(work_b.columns)),
            "CONFIANCA": st.column_config.TextColumn("Confiança", disabled=True),
            "SCORE": st.column_config.NumberColumn("Score", disabled=True),
        }, key="key_fields_grid_v16")
    selected_keys = key_df[key_df["USAR"] == True].copy()
    selected_keys = selected_keys[selected_keys["CAMPO_BASE_A"].notna() & selected_keys["CAMPO_BASE_B"].notna()].copy()
    selected_keys["CAMPO_BASE_A"] = selected_keys["CAMPO_BASE_A"].astype(str).str.strip()
    selected_keys["CAMPO_BASE_B"] = selected_keys["CAMPO_BASE_B"].astype(str).str.strip()
    selected_keys = selected_keys[selected_keys["CAMPO_BASE_A"].isin(work_a.columns) & selected_keys["CAMPO_BASE_B"].isin(work_b.columns)].copy()
    selected_keys = selected_keys.sort_values("ORDEM", ascending=True)
    if selected_keys.empty:
        st.warning("Selecione pelo menos um campo para localizar o registro correspondente.")
        return

    st.markdown("### 7) Quais campos deseja confrontar para validar os valores?")
    st.caption("O resultado sairá em uma única tabela, sempre na ordem valor da base A, valor da base B e diferença.")
    used_a = set(selected_keys["CAMPO_BASE_A"].tolist())
    used_b = set(selected_keys["CAMPO_BASE_B"].tolist())
    compare_rows = []
    if not sug_df.empty:
        tmp = sug_df[(~sug_df["CAMPO_BASE_A"].isin(used_a)) & (~sug_df["CAMPO_BASE_B"].isin(used_b))].copy()
        for i, (_, r) in enumerate(tmp.head(10).iterrows(), start=1):
            compare_rows.append({"COMPARAR": True if i <= 4 else False, "ORDEM": i, "CAMPO_BASE_A": r["CAMPO_BASE_A"], "CAMPO_BASE_B": r["CAMPO_BASE_B"], "TIPO": "Numérico", "TOLERANCIA": 0.00})
    if not compare_rows:
        compare_rows = [{"COMPARAR": True, "ORDEM": 1, "CAMPO_BASE_A": list(work_a.columns)[0], "CAMPO_BASE_B": list(work_b.columns)[0], "TIPO": "Numérico", "TOLERANCIA": 0.00}]
    compare_df = st.data_editor(pd.DataFrame(compare_rows), use_container_width=True, height=260, hide_index=True,
        column_config={
            "COMPARAR": st.column_config.CheckboxColumn("Comparar"),
            "ORDEM": st.column_config.NumberColumn("Ordem", min_value=1, step=1),
            "CAMPO_BASE_A": st.column_config.SelectboxColumn(f"Campo {nome_base_a}", options=list(work_a.columns)),
            "CAMPO_BASE_B": st.column_config.SelectboxColumn(f"Campo {nome_base_b}", options=list(work_b.columns)),
            "TIPO": st.column_config.SelectboxColumn("Tipo", options=["Numérico", "Texto exato", "Texto normalizado"]),
            "TOLERANCIA": st.column_config.NumberColumn("Tolerância", min_value=0.0, step=0.01),
        }, key="compare_fields_grid_v16")
    selected_compares = compare_df[compare_df["COMPARAR"] == True].copy()
    selected_compares = selected_compares[selected_compares["CAMPO_BASE_A"].notna() & selected_compares["CAMPO_BASE_B"].notna()].copy()
    selected_compares["CAMPO_BASE_A"] = selected_compares["CAMPO_BASE_A"].astype(str).str.strip()
    selected_compares["CAMPO_BASE_B"] = selected_compares["CAMPO_BASE_B"].astype(str).str.strip()
    selected_compares = selected_compares[selected_compares["CAMPO_BASE_A"].isin(work_a.columns) & selected_compares["CAMPO_BASE_B"].isin(work_b.columns)].copy()
    selected_compares = selected_compares.sort_values("ORDEM", ascending=True)

    st.markdown("### 8) Como deseja receber o resultado?")
    r1, r2, r3 = st.columns(3)
    with r1:
        mostrar_apenas_divergencias = st.checkbox("Mostrar apenas divergências no resultado", value=False)
    with r2:
        incluir_nao_encontrados = st.checkbox("Incluir não encontrados", value=True)
    with r3:
        gerar_resumo_exec = st.checkbox("Gerar resumo executivo", value=True)

    totalizador_map = {}

    def _add_totalizador_option(label, source_tuple):
        label = str(label or "").strip()
        if not label:
            return
        totalizador_map.setdefault(label, [])
        if source_tuple not in totalizador_map[label]:
            totalizador_map[label].append(source_tuple)

    # Opções de agrupamento: somente campos-chave do match + campos de associação confirmados.
    for _, key_row in selected_keys.iterrows():
        lbl = _canon_totalizer_label(key_row["CAMPO_BASE_A"], key_row["CAMPO_BASE_B"])
        _add_totalizador_option(lbl, (nome_base_a, key_row["CAMPO_BASE_A"]))
        _add_totalizador_option(lbl, (nome_base_b, key_row["CAMPO_BASE_B"]))

    for spec in association_specs:
        dst = str(spec.get("target_col_b") or "").strip()
        map_col = str(spec.get("map_col_a") or "").strip()
        if dst and map_col:
            lbl = _canon_totalizer_label(dst, None)
            _add_totalizador_option(lbl, (nome_base_a, map_col))
            _add_totalizador_option(lbl, (nome_base_b, dst))

    # Ordem estável e sem duplicidades.
    group_options = list(dict.fromkeys(totalizador_map.keys()))
    default_group_choices = []
    for _, key_row in selected_keys.iterrows():
        lbl = _canon_totalizer_label(key_row["CAMPO_BASE_A"], key_row["CAMPO_BASE_B"])
        if lbl in group_options and lbl not in default_group_choices:
            default_group_choices.append(lbl)

    compare_value_options = []
    for _, cp in selected_compares.iterrows():
        lbl = str(cp["CAMPO_BASE_A"]).strip()
        if lbl and lbl not in compare_value_options:
            compare_value_options.append(lbl)

    # Limpa seleções inválidas entre execuções/reloads.
    if "v18_group_choices" in st.session_state:
        st.session_state["v18_group_choices"] = [x for x in st.session_state["v18_group_choices"] if x in group_options]
    if "v18_value_choices" in st.session_state:
        st.session_state["v18_value_choices"] = [x for x in st.session_state["v18_value_choices"] if x in compare_value_options]

    # Defaults inteligentes: chaves do match para agrupar; campos comparados para totalizar.
    if not st.session_state.get("v18_group_choices") and default_group_choices:
        st.session_state["v18_group_choices"] = default_group_choices[:]
    if not st.session_state.get("v18_value_choices") and compare_value_options:
        st.session_state["v18_value_choices"] = compare_value_options[:]

    st.markdown("#### Como deseja montar o resumo executivo?")
    gx, vx = st.columns(2)
    with gx:
        group_choices = st.multiselect(
            "Agrupar resumo por",
            group_options,
            default=st.session_state.get("v18_group_choices", default_group_choices),
            help="Os campos-chave do match já vêm sugeridos aqui. Eles definem como o resumo executivo será agrupado.",
            key="v18_group_choices",
        )
    with vx:
        value_choices = st.multiselect(
            "O que deseja totalizar/confrontar",
            compare_value_options,
            default=st.session_state.get("v18_value_choices", compare_value_options),
            help="Esses são os campos de valor que serão somados e confrontados dentro do agrupamento escolhido.",
            key="v18_value_choices",
        )

    # Nunca quebrar por vazio: assume defaults calculados.
    if not group_choices:
        group_choices = default_group_choices[:] if default_group_choices else group_options[:1]
    if not value_choices:
        value_choices = compare_value_options[:]

    st.markdown("### 9) Processar análise")
    processar = st.button("Executar análise", type="primary", use_container_width=True)
    if not processar:
        return

    with st.spinner("Processando análise..."):
        base_a = work_a.copy()
        base_b = work_b.copy()
        for c in base_a.columns:
            base_a[c] = _force_text_series(base_a[c])
        for c in base_b.columns:
            base_b[c] = _force_text_series(base_b[c])

        key_cols_a, key_cols_b = [], []
        for i, (_, key_row) in enumerate(selected_keys.iterrows(), start=1):
            col_key_a = f"__KEY_A_{i}"
            col_key_b = f"__KEY_B_{i}"
            base_a[col_key_a] = _force_text_series(base_a[key_row["CAMPO_BASE_A"]]).str.strip()
            base_b[col_key_b] = _force_text_series(base_b[key_row["CAMPO_BASE_B"]]).str.strip()
            key_cols_a.append(col_key_a)
            key_cols_b.append(col_key_b)
        base_a["__KEY__"] = _build_key(base_a, key_cols_a)
        base_b["__KEY__"] = _build_key(base_b, key_cols_b)
        a_counts = base_a["__KEY__"].value_counts(dropna=False)
        b_counts = base_b["__KEY__"].value_counts(dropna=False)
        a_unique = base_a.drop_duplicates(subset=["__KEY__"], keep="first").copy()
        b_unique = base_b.drop_duplicates(subset=["__KEY__"], keep="first").copy()
        base_a_pref = a_unique.rename(columns={c: f"A__{c}" for c in a_unique.columns if c != "__KEY__"})
        base_b_pref = b_unique.rename(columns={c: f"B__{c}" for c in b_unique.columns if c != "__KEY__"})

        def _classify_row(a_count, b_count, origin_label):
            a_count = int(a_count or 0)
            b_count = int(b_count or 0)
            if a_count >= 2 and b_count >= 2:
                return "Duplicidade em ambas as bases", "Duplicidade"
            if origin_label == "A":
                if a_count >= 2:
                    return f"Duplicidade na {nome_base_a}", "Duplicidade"
                if b_count >= 2:
                    return f"Duplicidade na {nome_base_b}", "Duplicidade"
                if b_count == 0:
                    return f"Não encontrado na {nome_base_b}", "Sem correspondência"
                if a_count == 0:
                    return f"Não encontrado na {nome_base_a}", "Sem correspondência"
                return "Encontrado", "Match exato"
            else:
                if b_count >= 2:
                    return f"Duplicidade na {nome_base_b}", "Duplicidade"
                if a_count >= 2:
                    return f"Duplicidade na {nome_base_a}", "Duplicidade"
                if a_count == 0:
                    return f"Não encontrado na {nome_base_a}", "Sem correspondência"
                if b_count == 0:
                    return f"Não encontrado na {nome_base_b}", "Sem correspondência"
                return "Encontrado", "Match exato"

        merged = base_a_pref.merge(base_b_pref, on="__KEY__", how="outer")
        df_result = pd.DataFrame(index=merged.index)
        df_result["CHAVE_PROCESSADA"] = merged["__KEY__"].fillna("")
        df_result["__A_COUNT__"] = merged["__KEY__"].map(a_counts).fillna(0).astype(int)
        df_result["__B_COUNT__"] = merged["__KEY__"].map(b_counts).fillna(0).astype(int)
        cen_res = df_result.apply(lambda r: _classify_row(r["__A_COUNT__"], r["__B_COUNT__"], "A"), axis=1, result_type="expand")
        df_result["CENARIO_MATCH"] = cen_res[0]
        df_result["RESULTADO_FINAL"] = cen_res[1]
        df_result["STATUS_MATCH"] = df_result["CENARIO_MATCH"]
        has_div = pd.Series(False, index=df_result.index)
        compare_meta = []

        # Campos-chave visíveis no resultado completo (duas origens) e canônicos no resumo.
        for _, key_row in selected_keys.iterrows():
            ca = f"CHAVE_{nome_base_a}_{key_row['CAMPO_BASE_A']}"
            cb = f"CHAVE_{nome_base_b}_{key_row['CAMPO_BASE_B']}"
            df_result[ca] = merged.get(f"A__{key_row['CAMPO_BASE_A']}", pd.Series("", index=merged.index)).fillna("")
            df_result[cb] = merged.get(f"B__{key_row['CAMPO_BASE_B']}", pd.Series("", index=merged.index)).fillna("")

        totalizador_result_cols = []
        for choice in group_choices:
            if choice not in totalizador_map:
                continue
            source_tuples = totalizador_map[choice]
            candidate_cols = []
            for lado, campo in source_tuples:
                if lado == nome_base_a:
                    col = merged.get(f"A__{campo}", pd.Series("", index=merged.index))
                else:
                    col = merged.get(f"B__{campo}", pd.Series("", index=merged.index))
                candidate_cols.append(_force_text_series(col))
            if not candidate_cols:
                continue
            canon_df = pd.concat(candidate_cols, axis=1).replace("", np.nan)
            res_col = f"TOTALIZADOR__{choice}"
            df_result[res_col] = canon_df.bfill(axis=1).iloc[:, 0].fillna("")
            totalizador_result_cols.append(res_col)

        for _, cp in selected_compares.iterrows():
            label = str(cp["CAMPO_BASE_A"]).strip()
            sa = merged.get(f"A__{cp['CAMPO_BASE_A']}", pd.Series("", index=merged.index))
            sb = merged.get(f"B__{cp['CAMPO_BASE_B']}", pd.Series("", index=merged.index))
            raw_a, raw_b, diff, status = _compare_series(sa, sb, mode=cp["TIPO"], tol=float(cp["TOLERANCIA"] or 0.0))
            c_va = f"VALOR_{nome_base_a}_{cp['CAMPO_BASE_A']}"
            c_vb = f"VALOR_{nome_base_b}_{cp['CAMPO_BASE_B']}"
            c_df = f"DIF_{cp['CAMPO_BASE_A']}__{cp['CAMPO_BASE_B']}"
            df_result[c_va] = raw_a
            df_result[c_vb] = raw_b
            df_result[c_df] = diff.round(2) if pd.api.types.is_numeric_dtype(diff) else diff
            has_div = has_div | (df_result["CENARIO_MATCH"].eq("Encontrado") & status.eq("Divergente"))
            compare_meta.append({
                "label": label,
                "val_a_col": c_va,
                "val_b_col": c_vb,
                "diff_col": c_df,
                "selected": label in value_choices,
            })
        df_result.loc[df_result["CENARIO_MATCH"].eq("Encontrado") & has_div, "RESULTADO_FINAL"] = "Match com divergência"
        df_result = df_result.drop(columns=["__A_COUNT__", "__B_COUNT__"], errors="ignore")
        if not incluir_nao_encontrados:
            df_result = df_result[~df_result["CENARIO_MATCH"].astype(str).str.contains("Não encontrado", na=False)].copy()
        if mostrar_apenas_divergencias:
            df_result = df_result[df_result["RESULTADO_FINAL"] == "Match com divergência"].copy()

        key_order_cols = []
        for _, key_row in selected_keys.iterrows():
            ca = f"CHAVE_{nome_base_a}_{key_row['CAMPO_BASE_A']}"
            cb = f"CHAVE_{nome_base_b}_{key_row['CAMPO_BASE_B']}"
            if ca in df_result.columns:
                key_order_cols.append(ca)
            if cb in df_result.columns:
                key_order_cols.append(cb)

        compare_order_cols = []
        for _, cp in selected_compares.iterrows():
            cols_cmp = [f"VALOR_{nome_base_a}_{cp['CAMPO_BASE_A']}", f"VALOR_{nome_base_b}_{cp['CAMPO_BASE_B']}", f"DIF_{cp['CAMPO_BASE_A']}__{cp['CAMPO_BASE_B']}"]
            for c in cols_cmp:
                if c in df_result.columns:
                    compare_order_cols.append(c)

        final_cols = [c for c in ["CHAVE_PROCESSADA"] if c in df_result.columns]
        final_cols.extend(key_order_cols)
        final_cols.extend([c for c in totalizador_result_cols if c in df_result.columns])
        final_cols.extend(compare_order_cols)
        final_cols.extend([c for c in ["CENARIO_MATCH", "RESULTADO_FINAL"] if c in df_result.columns])
        final_cols = final_cols + [c for c in df_result.columns if c not in final_cols]
        df_result = df_result[final_cols].copy()

        total = len(df_result)
        encontrados = int(df_result["CENARIO_MATCH"].astype(str).eq("Encontrado").sum()) if total else 0
        nao_encontrados = int(df_result["CENARIO_MATCH"].astype(str).str.contains("Não encontrado", na=False).sum()) if total else 0
        duplicados_result = int(df_result["CENARIO_MATCH"].astype(str).str.contains("Duplicidade", na=False).sum()) if total else 0
        divergentes = int(df_result["RESULTADO_FINAL"].astype(str).eq("Match com divergência").sum()) if total else 0
        aderencia = ((encontrados / total) * 100.0) if total else 0.0
        resumo_dict = {
            "objetivo": objetivo_label,
            "direcao": "Auditoria completa",
            "total": total,
            "encontrados": encontrados,
            "nao_encontrados": nao_encontrados,
            "duplicados": duplicados_result,
            "divergentes": divergentes,
            "aderencia": round(aderencia, 2),
        }
        if gerar_resumo_exec:
            totalizador_df, financial_kpis = _build_totalizador(df_result, compare_meta, totalizador_result_cols)
        else:
            totalizador_df, financial_kpis = pd.DataFrame(), []
        text_priority_cols = [c for c in df_result.columns if any(x in c.upper() for x in ["CHAVE", "COD", "CÓD", "FILIAL", "LOJA", "PATRIM", "PLAQUETA", "TOTALIZADOR"])]
        excel_bytes = _to_excel_package(df_result=df_result, resumo_dict=resumo_dict, text_priority_cols=text_priority_cols, totalizador_df=totalizador_df, financial_kpis=financial_kpis)

    st.markdown("### Resultado da análise")
    m1, m2, m3, m4, m5 = st.columns(5)
    with m1:
        st.metric("Registros analisados", total)
    with m2:
        st.metric("Correspondências", encontrados)
    with m3:
        st.metric("Ausentes", nao_encontrados)
    with m4:
        st.metric("Duplicidades", duplicados_result)
    with m5:
        st.metric("Divergências", divergentes)

    st.dataframe(df_result, use_container_width=True, height=520)
    if gerar_resumo_exec and totalizador_df is not None and not totalizador_df.empty:
        st.markdown("#### Resumo executivo por chave")
        st.dataframe(totalizador_df, use_container_width=True, height=260)
    st.download_button(
        "Baixar resultado em Excel",
        data=excel_bytes,
        file_name=f"Match_Inteligente_V18_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# =========================================================
# Sidebar
# =========================================================
with st.sidebar:
    st.markdown("## Navegação")
    mod = st.radio("Módulo", ["Financeiro", "Match Inteligente"], index=0)

    if mod == "Financeiro":
        area = st.radio(
            "Área",
            ["Extrato Bancário", "Posição a Pagar", "Posição a Receber"],
            index=0
        )
    else:
        area = st.radio(
            "Área",
            ["Match Inteligente"],
            index=0
        )

    st.markdown("---")
    st.caption(f"Você está em: {mod} > {area}")

if mod == "Match Inteligente" and area == "Match Inteligente":
    render_cruzamento_inteligente_v2()
    st.stop()

elif mod != "Financeiro" or area != "Extrato Bancário":
    st.title("ConciliaMais")
    st.info("Esta área ainda está em construção. Por enquanto, use Financeiro > Extrato Bancário.")
    st.stop()

    st.markdown("### 2) Visão inicial das bases")
    v1, v2 = st.columns(2)

    with v1:
        st.markdown("**Base A**")
        st.caption(f"{len(df_a):,} linhas | {len(df_a.columns)} colunas")
        st.dataframe(df_a.head(10), use_container_width=True, height=260)

    with v2:
        st.markdown("**Base B**")
        st.caption(f"{len(df_b):,} linhas | {len(df_b.columns)} colunas")
        st.dataframe(df_b.head(10), use_container_width=True, height=260)

    suggestions = _suggest_columns(list(df_a.columns), list(df_b.columns))

    st.markdown("### 3) Configuração do cruzamento")
    if suggestions:
        top_sug = pd.DataFrame(suggestions[:12], columns=["COLUNA_A", "COLUNA_B", "SCORE"])
        st.markdown("**Sugestões automáticas de colunas parecidas**")
        st.dataframe(top_sug, use_container_width=True, height=220)

    k1, k2 = st.columns(2)

    default_a = [suggestions[0][0]] if suggestions else []
    default_b = [suggestions[0][1]] if suggestions else []

    with k1:
        key_a = st.multiselect(
            "Campo(s) chave da Base A",
            options=list(df_a.columns),
            default=default_a
        )

    with k2:
        key_b = st.multiselect(
            "Campo(s) chave da Base B",
            options=list(df_b.columns),
            default=default_b
        )

    if len(key_a) != len(key_b):
        st.warning("A quantidade de campos-chave da Base A e da Base B deve ser igual.")
        st.stop()

    r1, r2 = st.columns(2)

    with r1:
        retorno_cols_b = st.multiselect(
            "Campos da Base B para retornar",
            options=list(df_b.columns)
        )

    with r2:
        comparar_valores = st.checkbox("Comparar um campo numérico entre as bases", value=False)

    col_val_a = None
    col_val_b = None

    if comparar_valores:
        cva, cvb = st.columns(2)
        with cva:
            col_val_a = st.selectbox("Campo numérico da Base A", list(df_a.columns), key="procv_val_a")
        with cvb:
            col_val_b = st.selectbox("Campo numérico da Base B", list(df_b.columns), key="procv_val_b")

    processar = st.button("Processar cruzamento", type="primary", use_container_width=True)

    if not processar:
        st.stop()

    if not key_a or not key_b:
        st.warning("Selecione pelo menos uma chave de relacionamento.")
        st.stop()

    with st.spinner("Processando cruzamento..."):
        base_a = df_a.copy()
        base_b = df_b.copy()

        base_a["__KEY__"] = _build_key(base_a, key_a)
        base_b["__KEY__"] = _build_key(base_b, key_b)

        dup_b = base_b["__KEY__"].value_counts()
        dup_keys_b = set(dup_b[dup_b > 1].index.tolist())

        b_lookup = base_b.drop_duplicates(subset="__KEY__", keep="first").set_index("__KEY__", drop=False)

        out_rows = []

        for _, row_a in base_a.iterrows():
            key_val = row_a["__KEY__"]
            row_out = {}

            for c in key_a:
                row_out[f"CHAVE_A_{c}"] = row_a.get(c, "")

            row_out["STATUS_MATCH"] = "Não encontrado"
            row_out["CHAVE_PROCESSADA"] = key_val

            if key_val in dup_keys_b:
                row_out["STATUS_MATCH"] = "Duplicidade na Base B"
                row_out["RESULTADO_FINAL"] = "Duplicidade"
            elif key_val in b_lookup.index:
                row_b = b_lookup.loc[key_val]
                row_out["STATUS_MATCH"] = "Encontrado"

                for c in retorno_cols_b:
                    row_out[f"RETORNO_B_{c}"] = row_b.get(c, "")

                if comparar_valores and col_val_a and col_val_b:
                    val_a = normalize_money(row_a.get(col_val_a, np.nan))
                    val_b = normalize_money(row_b.get(col_val_b, np.nan))

                    row_out[f"VALOR_A_{col_val_a}"] = val_a
                    row_out[f"VALOR_B_{col_val_b}"] = val_b

                    if pd.notna(val_a) and pd.notna(val_b):
                        diff = round(float(val_a) - float(val_b), 2)
                        row_out["DIFERENCA"] = diff
                        row_out["RESULTADO_FINAL"] = "Match exato" if abs(diff) <= 0.01 else "Match com divergência"
                    else:
                        row_out["DIFERENCA"] = np.nan
                        row_out["RESULTADO_FINAL"] = "Match encontrado"
                else:
                    row_out["RESULTADO_FINAL"] = "Match encontrado"
            else:
                for c in retorno_cols_b:
                    row_out[f"RETORNO_B_{c}"] = ""
                if comparar_valores and col_val_a:
                    row_out[f"VALOR_A_{col_val_a}"] = normalize_money(row_a.get(col_val_a, np.nan))
                if comparar_valores and col_val_b:
                    row_out[f"VALOR_B_{col_val_b}"] = np.nan
                    row_out["DIFERENCA"] = np.nan
                row_out["RESULTADO_FINAL"] = "Sem correspondência"

            out_rows.append(row_out)

        df_result = pd.DataFrame(out_rows)

    st.markdown("### 4) Resumo do processamento")
    total = len(df_result)
    encontrados = int((df_result["STATUS_MATCH"] == "Encontrado").sum()) if total else 0
    nao_encontrados = int((df_result["STATUS_MATCH"] == "Não encontrado").sum()) if total else 0
    duplicados = int((df_result["STATUS_MATCH"] == "Duplicidade na Base B").sum()) if total else 0
    divergentes = int((df_result["RESULTADO_FINAL"] == "Match com divergência").sum()) if total and "RESULTADO_FINAL" in df_result.columns else 0

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("Linhas processadas", total)
    with m2:
        st.metric("Encontrados", encontrados)
    with m3:
        st.metric("Não encontrados", nao_encontrados)
    with m4:
        st.metric("Divergentes", divergentes)

    if duplicados > 0:
        st.warning(f"Foram encontradas {duplicados} chave(s) duplicadas na Base B.")

    st.markdown("### 5) Resultado")
    f1, f2 = st.columns([1.2, 2.0])

    with f1:
        filtro_resultado = st.selectbox(
            "Filtrar resultado",
            ["Todos", "Match exato", "Match com divergência", "Sem correspondência", "Match encontrado", "Duplicidade"]
        )

    with f2:
        busca_procv = st.text_input("Buscar no resultado", value="")

    df_show = df_result.copy()

    if filtro_resultado != "Todos":
        if filtro_resultado == "Duplicidade":
            df_show = df_show[df_show["RESULTADO_FINAL"] == "Duplicidade"].copy()
        else:
            df_show = df_show[df_show["RESULTADO_FINAL"] == filtro_resultado].copy()

    if busca_procv.strip():
        q = busca_procv.strip().lower()
        mask = pd.Series(False, index=df_show.index)
        for c in df_show.columns:
            mask = mask | df_show[c].astype(str).str.lower().str.contains(q, na=False)
        df_show = df_show[mask].copy()

    st.dataframe(df_show, use_container_width=True, height=520)

    excel_out = BytesIO()
    with pd.ExcelWriter(excel_out, engine="xlsxwriter") as writer:
        df_show.to_excel(writer, sheet_name="Cruzamento", index=False)
    excel_out.seek(0)

    st.download_button(
        "Baixar resultado em Excel",
        data=excel_out,
        file_name=f"Cruzamento_Inteligente_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.stop()

elif mod != "Financeiro" or area != "Extrato Bancário":
    st.title("ConciliaMais")
    st.info("Esta área ainda está em construção. Por enquanto, use Financeiro > Extrato Bancário.")
    st.stop()

# =========================================================
# Upload
# =========================================================
if st.session_state.page == "upload":
    st.title("ConciliaMais — Conferência de Extrato Bancário")
    st.markdown('<div class="cm-breadcrumb">Financeiro  ›  Extrato Bancário</div>', unsafe_allow_html=True)
    st.caption("Extrato Financeiro + Razão Contábil → Match automático → Divergências → Tratativa")
    show_flash()

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
    st.markdown('<div class="cm-help">Ao processar, o sistema gera divergências e habilita tratativa.</div>', unsafe_allow_html=True)

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
# Resultados
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

    show_flash()

    div_master = st.session_state.div_master.copy()
    div_master["VALOR"] = div_master["VALOR"].map(normalize_money)
    div_master["RESOLVIDO"] = div_master["RESOLVIDO"].fillna(False)
    div_master["STATUS"] = div_master["STATUS"].fillna("Pendente").astype(str)
    div_master["CONFIRMADO"] = div_master.get("CONFIRMADO", False)
    div_master["NUCLEO"] = div_master.get("NUCLEO", "Não identificado").fillna("Não identificado")
    div_master["MOTIVO_BASE"] = div_master.get("MOTIVO_BASE", div_master["HISTORICO_OPERACAO"].map(build_motivo_base))

    if "NUCLEO_SUGERIDO" in div_master.columns:
        need = div_master["CONFIRMADO"] & (
            div_master["NUCLEO"].astype(str).str.strip().eq("") |
            div_master["NUCLEO"].eq("Não identificado")
        )
        div_master.loc[need, "NUCLEO"] = div_master.loc[need, "NUCLEO_SUGERIDO"].fillna("Não identificado")

    div_master.loc[div_master["RESOLVIDO"], "STATUS"] = "Resolvido"

    if "SEVERIDADE" not in div_master.columns:
        div_master = apply_classification_rules(div_master)
    if "SELECIONADO" not in div_master.columns:
        div_master["SELECIONADO"] = False

    div_master["NUCLEO_EXIBICAO"] = get_nucleo_display_series(div_master)
    div_master["ORIGEM_VISUAL"] = div_master["ORIGEM"].map(origem_visual_text)

    st.session_state.div_master = div_master
    current_nucleos = get_nucleos()

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
    # Biblioteca de núcleos
    # =====================================================
    with st.expander("Biblioteca de núcleos", expanded=False):
        st.markdown("#### Núcleos disponíveis")
        nuc_df = pd.DataFrame({"NUCLEO": current_nucleos})
        st.dataframe(nuc_df, use_container_width=True, height=220)

        n1, n2, n3 = st.columns(3)

        with n1:
            st.markdown("**Criar núcleo**")
            with st.form("form_add_nucleo", clear_on_submit=True):
                novo_nucleo = st.text_input("Novo núcleo")
                submit_add_nucleo = st.form_submit_button("Criar núcleo", type="primary")
            if submit_add_nucleo:
                ok, msg = add_nucleo(novo_nucleo)
                set_flash("success" if ok else "warning", msg)
                st.rerun()

        with n2:
            st.markdown("**Renomear núcleo**")
            with st.form("form_rename_nucleo", clear_on_submit=True):
                old_nucleo = st.selectbox("Núcleo atual", current_nucleos, key="old_nucleo")
                new_nucleo = st.text_input("Novo nome")
                submit_rename_nucleo = st.form_submit_button("Renomear núcleo", type="primary")
            if submit_rename_nucleo:
                ok, msg = rename_nucleo(old_nucleo, new_nucleo)
                set_flash("success" if ok else "warning", msg)
                st.rerun()

        with n3:
            st.markdown("**Excluir núcleo customizado**")
            custom_nucleos = [x for x in current_nucleos if x not in DEFAULT_NUCLEOS]
            with st.form("form_delete_nucleo", clear_on_submit=False):
                del_nucleo = st.selectbox("Núcleo para excluir", custom_nucleos if custom_nucleos else ["(nenhum)"])
                submit_delete_nucleo = st.form_submit_button("Excluir núcleo", type="primary", disabled=(len(custom_nucleos) == 0))
            if submit_delete_nucleo and custom_nucleos:
                ok, msg = delete_nucleo(del_nucleo)
                set_flash("success" if ok else "warning", msg)
                st.rerun()

    # =====================================================
    # Biblioteca de regras
    # =====================================================
    with st.expander("Biblioteca de regras (persistente)", expanded=False):
        payload = load_rules()

        st.markdown("#### Criar regra de Núcleo")
        with st.form("form_regra_nucleo", clear_on_submit=True):
            c1, c2, c3 = st.columns([1.4, 1.0, 1.0])
            with c1:
                nr_nome = st.text_input("Nome da regra (núcleo)")
                nr_texto = st.text_input("Texto contém")
                nr_regex = st.text_input("Regex (opcional)")
                nr_doc_pref = st.text_input("Prefixo do documento (opcional)")
            with c2:
                nr_origem = st.selectbox("Origem", ORIGEM_RULE_OPTS, index=0)
                nr_valor_min = st.text_input("Valor mínimo abs")
                nr_valor_max = st.text_input("Valor máximo abs")
            with c3:
                nr_resultado = st.selectbox("Resultado", get_nucleos())
                nr_prioridade = st.number_input("Prioridade", min_value=1, value=100, step=1)
                nr_ativa = st.checkbox("Ativa", value=True)
                salvar_nucleo = st.form_submit_button("Salvar regra de Núcleo", type="primary")

        if salvar_nucleo:
            ok, msg = add_rule("nucleo", {
                "nome": nr_nome.strip() or f"Núcleo {nr_resultado}",
                "prioridade": int(nr_prioridade),
                "ativa": bool(nr_ativa),
                "origem": nr_origem,
                "texto_contem": nr_texto.strip(),
                "regex": nr_regex.strip(),
                "documento_prefixo": nr_doc_pref.strip(),
                "valor_min": nr_valor_min.strip(),
                "valor_max": nr_valor_max.strip(),
                "resultado": nr_resultado,
            })
            if ok:
                dm = st.session_state.div_master.copy()
                dm = apply_classification_rules(dm)
                dm["NUCLEO_EXIBICAO"] = get_nucleo_display_series(dm)
                st.session_state.div_master = dm
                set_flash("success", msg)
            else:
                set_flash("warning", msg)
            st.rerun()

        st.markdown("---")
        st.markdown("#### Criar regra de Criticidade")
        with st.form("form_regra_criticidade", clear_on_submit=True):
            d1, d2, d3 = st.columns([1.4, 1.0, 1.0])
            with d1:
                cr_nome = st.text_input("Nome da regra (criticidade)")
                cr_texto = st.text_input("Texto contém ")
                cr_regex = st.text_input("Regex (opcional) ")
                cr_doc_pref = st.text_input("Prefixo do documento (opcional) ")
            with d2:
                cr_origem = st.selectbox("Origem ", ORIGEM_RULE_OPTS, index=0)
                cr_valor_min = st.text_input("Valor mínimo abs ")
                cr_valor_max = st.text_input("Valor máximo abs ")
            with d3:
                cr_resultado = st.selectbox("Resultado ", SEVERIDADES)
                cr_prioridade = st.number_input("Prioridade ", min_value=1, value=100, step=1)
                cr_ativa = st.checkbox("Ativa ", value=True)
                salvar_criticidade = st.form_submit_button("Salvar regra de Criticidade", type="primary")

        if salvar_criticidade:
            ok, msg = add_rule("criticidade", {
                "nome": cr_nome.strip() or f"Criticidade {cr_resultado}",
                "prioridade": int(cr_prioridade),
                "ativa": bool(cr_ativa),
                "origem": cr_origem,
                "texto_contem": cr_texto.strip(),
                "regex": cr_regex.strip(),
                "documento_prefixo": cr_doc_pref.strip(),
                "valor_min": cr_valor_min.strip(),
                "valor_max": cr_valor_max.strip(),
                "resultado": cr_resultado,
            })
            if ok:
                dm = st.session_state.div_master.copy()
                dm = apply_classification_rules(dm)
                dm["NUCLEO_EXIBICAO"] = get_nucleo_display_series(dm)
                st.session_state.div_master = dm
                set_flash("success", msg)
            else:
                set_flash("warning", msg)
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
                    if update_rule_status("nucleo", rid, True):
                        dm = st.session_state.div_master.copy()
                        dm = apply_classification_rules(dm)
                        dm["NUCLEO_EXIBICAO"] = get_nucleo_display_series(dm)
                        st.session_state.div_master = dm
                        set_flash("success", "Regra de núcleo ativada.")
                    else:
                        set_flash("warning", "ID de regra não encontrado.")
                    st.rerun()
            with colx2:
                if st.button("Inativar regra núcleo"):
                    if update_rule_status("nucleo", rid, False):
                        dm = st.session_state.div_master.copy()
                        dm = apply_classification_rules(dm)
                        dm["NUCLEO_EXIBICAO"] = get_nucleo_display_series(dm)
                        st.session_state.div_master = dm
                        set_flash("success", "Regra de núcleo inativada.")
                    else:
                        set_flash("warning", "ID de regra não encontrado.")
                    st.rerun()
            with colx3:
                if st.button("Excluir regra núcleo"):
                    if delete_rule("nucleo", rid):
                        dm = st.session_state.div_master.copy()
                        dm = apply_classification_rules(dm)
                        dm["NUCLEO_EXIBICAO"] = get_nucleo_display_series(dm)
                        st.session_state.div_master = dm
                        set_flash("success", "Regra de núcleo excluída.")
                    else:
                        set_flash("warning", "ID de regra não encontrado.")
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
                    if update_rule_status("criticidade", rid2, True):
                        dm = st.session_state.div_master.copy()
                        dm = apply_classification_rules(dm)
                        dm["NUCLEO_EXIBICAO"] = get_nucleo_display_series(dm)
                        st.session_state.div_master = dm
                        set_flash("success", "Regra de criticidade ativada.")
                    else:
                        set_flash("warning", "ID de regra não encontrado.")
                    st.rerun()
            with coly2:
                if st.button("Inativar regra criticidade"):
                    if update_rule_status("criticidade", rid2, False):
                        dm = st.session_state.div_master.copy()
                        dm = apply_classification_rules(dm)
                        dm["NUCLEO_EXIBICAO"] = get_nucleo_display_series(dm)
                        st.session_state.div_master = dm
                        set_flash("success", "Regra de criticidade inativada.")
                    else:
                        set_flash("warning", "ID de regra não encontrado.")
                    st.rerun()
            with coly3:
                if st.button("Excluir regra criticidade"):
                    if delete_rule("criticidade", rid2):
                        dm = st.session_state.div_master.copy()
                        dm = apply_classification_rules(dm)
                        dm["NUCLEO_EXIBICAO"] = get_nucleo_display_series(dm)
                        st.session_state.div_master = dm
                        set_flash("success", "Regra de criticidade excluída.")
                    else:
                        set_flash("warning", "ID de regra não encontrado.")
                    st.rerun()

    # =====================================================
    # Resumo / motor de aprendizado
    # =====================================================
    with st.expander("Resumo para priorização + motor de aprendizado", expanded=True):
        df_open = div_master.loc[~resolved_mask].copy()
        df_open["ABS"] = df_open["VALOR"].abs()
        df_open["NUCLEO_EXIBICAO"] = get_nucleo_display_series(df_open)

        st.markdown("**Top 10 em aberto por impacto**")
        t1, t2, t3 = st.columns([1.1, 1.8, 2.1], gap="large")
        with t1:
            top_origem = st.selectbox("Origem (Top 10)", ["Todas", "Somente Financeiro", "Somente Contábil"], key="top10_origem")
        with t2:
            nuc_opts = ["Todos"] + sorted([x for x in df_open["NUCLEO_EXIBICAO"].fillna("Não identificado").unique().tolist() if str(x).strip() != ""])
            top_nucleo = st.selectbox("Núcleo (Top 10)", nuc_opts, key="top10_nucleo")
        with t3:
            st.caption("Estes filtros atuam apenas no Top 10.")

        top_src = df_open.copy()
        if top_origem != "Todas":
            top_src = top_src[top_src["ORIGEM"] == top_origem].copy()
        if top_nucleo != "Todos":
            top_src = top_src[top_src["NUCLEO_EXIBICAO"] == top_nucleo].copy()

        top_open = top_src.sort_values("ABS", ascending=False).head(10)
        show_cols = ["ORIGEM", "DATA", "DOCUMENTO", "VALOR", "NUCLEO_EXIBICAO"]
        st.dataframe(
            top_open[show_cols].rename(columns={"NUCLEO_EXIBICAO": "NUCLEO"}).copy(),
            use_container_width=True,
            height=320
        )

        if len(df_open):
            st.markdown("**Distribuição por Origem (abertos)**")
            dist_origem = df_open.groupby("ORIGEM", dropna=False).agg(Itens=("VALOR","size"), Valor=("VALOR","sum")).reset_index().sort_values("Valor", ascending=False)
            st.dataframe(dist_origem, use_container_width=True, height=160)

            st.markdown("**Distribuição por Origem × Núcleo (abertos)**")
            dist_on = df_open.groupby(["ORIGEM","NUCLEO_EXIBICAO"], dropna=False).agg(Itens=("VALOR","size"), Valor=("VALOR","sum")).reset_index().sort_values(["ORIGEM","Valor"], ascending=[True, False])
            st.dataframe(dist_on.rename(columns={"NUCLEO_EXIBICAO": "NUCLEO"}), use_container_width=True, height=220)

            st.markdown("**Comparativo (abertos): Financeiro × Contábil**")
            comp = df_open.groupby("ORIGEM", dropna=False)["VALOR"].sum().reset_index()
            comp = comp[comp["ORIGEM"].isin(["Somente Financeiro", "Somente Contábil"])].copy()
            comp = comp.set_index("ORIGEM")
            st.bar_chart(comp["VALOR"])

            st.markdown("**Motivos detectados (painel técnico de apoio)**")
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

            st.markdown("**Sugestões automáticas de novas regras**")
            sug = build_learning_suggestions(st.session_state.div_master)

            if sug.empty:
                st.info("Ainda não há sugestões automáticas suficientes. Confirme e ajuste alguns itens para o motor aprender.")
            else:
                st.dataframe(sug.head(20), use_container_width=True, height=280)

                sx1, sx2 = st.columns([2.2, 1.0])
                with sx1:
                    sug_idx = st.number_input("Linha da sugestão (começando em 0)", min_value=0, max_value=max(0, len(sug.head(20)) - 1), step=1, value=0)
                with sx2:
                    sug_prio = st.number_input("Prioridade da regra sugerida", min_value=1, value=70, step=1)

                if st.button("Transformar sugestão em regra", type="primary"):
                    sug_top = sug.head(20).reset_index(drop=True)
                    if int(sug_idx) < len(sug_top):
                        row = sug_top.loc[int(sug_idx)]
                        ok, msg = add_rule("nucleo", {
                            "nome": f'Sugestão: {str(row["MOTIVO_BASE"])[:50]}',
                            "prioridade": int(sug_prio),
                            "ativa": True,
                            "origem": row["ORIGEM"],
                            "texto_contem": row["MOTIVO_BASE"],
                            "regex": "",
                            "documento_prefixo": "",
                            "valor_min": "",
                            "valor_max": "",
                            "resultado": row["NUCLEO_FINAL"],
                        })
                        if ok:
                            dm = st.session_state.div_master.copy()
                            dm = apply_classification_rules(dm)
                            dm["NUCLEO_EXIBICAO"] = get_nucleo_display_series(dm)
                            st.session_state.div_master = dm
                            set_flash("success", msg)
                        else:
                            set_flash("warning", msg)
                        st.rerun()
        else:
            st.info("Sem pendências em aberto.")

    # =====================================================
    # Filtros
    # =====================================================
    st.markdown("### Filtros de análise")

    f1, f2, f3, f4 = st.columns([1.0, 1.1, 1.0, 2.2], gap="large")
    with f1:
        origem = st.selectbox("Origem", ["Todas", "Somente Financeiro", "Somente Contábil"])
    with f2:
        ver = st.selectbox("Visualizar", ["Todas", "Somente em aberto", "Somente resolvidas"])
    with f3:
        sev = st.selectbox("Severidade", ["Todas", "Normal", "Atenção", "Crítica"])
    with f4:
        busca = st.text_input("Buscar (documento, histórico, valor, núcleo)", value="")

    f5, f6, f7, f8 = st.columns([1.2, 1.0, 1.0, 1.0], gap="large")
    with f5:
        nucleo_filtro = st.selectbox("Núcleo", ["Todos"] + get_nucleos())
    with f6:
        status_filtro = st.selectbox("Status", ["Todos"] + STATUS_OPTS)
    with f7:
        confirmado_filtro = st.selectbox("Confirmado", ["Todos", "Sim", "Não"])
    with f8:
        st.markdown("<div style='height:1px'></div>", unsafe_allow_html=True)

    df = div_master.copy()
    df["NUCLEO_EXIBICAO"] = get_nucleo_display_series(df)

    if origem != "Todas":
        df = df[df["ORIGEM"] == origem].copy()

    res_mask_df = df["RESOLVIDO"] | (df["STATUS"].astype(str).str.lower().eq("resolvido"))
    if ver == "Somente em aberto":
        df = df[~res_mask_df].copy()
    elif ver == "Somente resolvidas":
        df = df[res_mask_df].copy()

    if sev != "Todas":
        df = df[df["SEVERIDADE"] == sev].copy()

    if nucleo_filtro != "Todos":
        df = df[df["NUCLEO_EXIBICAO"] == nucleo_filtro].copy()

    if status_filtro != "Todos":
        df = df[df["STATUS"] == status_filtro].copy()

    if confirmado_filtro != "Todos":
        want = confirmado_filtro == "Sim"
        df = df[df["CONFIRMADO"].fillna(False) == want].copy()

    if busca.strip():
        q = busca.strip().lower()
        mask = pd.Series(False, index=df.index)

        cols_search = ["DOCUMENTO", "HISTORICO_OPERACAO", "CHAVE_DOC", "NUCLEO", "NUCLEO_EXIBICAO", "ORIGEM", "SEVERIDADE", "STATUS", "OBS_USUARIO"]
        for c in cols_search:
            if c in df.columns:
                mask = mask | df[c].astype(str).str.lower().str.contains(q, na=False)

        if "VALOR" in df.columns:
            mask = mask | df["VALOR"].map(lambda x: fmt(x)).astype(str).str.lower().str.contains(q, na=False)

        df = df[mask].copy()

    total_filtrado = float(df["VALOR"].sum()) if not df.empty else 0.0
    with f8:
        st.markdown(
            f"""
<div class="cm-mini">
  <div class="k">Total do filtro</div>
  <div class="v">{fmt(total_filtrado)}</div>
</div>
""",
            unsafe_allow_html=True,
        )

    dfx = build_sort_columns(df)
    dfx = dfx.sort_values(
        by=["__RES", "__SEV_ORD", "__ABS_VAL", "__DATA_SORT"],
        ascending=[True, False, False, True]
    )
    df = dfx.drop(columns=["__RES", "__SEV_ORD", "__ABS_VAL", "__DATA_SORT"], errors="ignore")

    # =====================================================
    # Ações em massa
    # =====================================================
    st.markdown("### Ações em massa")
    st.markdown('<div class="cm-help">A experiência foi reorganizada para facilitar seleção, escopo e aplicação das alterações.</div>', unsafe_allow_html=True)

    ids_filtrados = list(df.index)
    dm0 = st.session_state.div_master.copy()
    selecionados_count = int(dm0["SELECIONADO"].fillna(False).sum())

    st.markdown('<div class="cm-actionbar">', unsafe_allow_html=True)
    s1, s2, s3, s4 = st.columns([1.2, 1.2, 1.0, 2.0], gap="large")
    with s1:
        if st.button("Selecionar todos do filtro", use_container_width=True):
            dm = st.session_state.div_master.copy()
            dm.loc[ids_filtrados, "SELECIONADO"] = True
            st.session_state.div_master = dm
            st.rerun()
    with s2:
        if st.button("Limpar seleção do filtro", use_container_width=True):
            dm = st.session_state.div_master.copy()
            dm.loc[ids_filtrados, "SELECIONADO"] = False
            st.session_state.div_master = dm
            st.rerun()
    with s3:
        st.markdown(f'<span class="cm-badge">Selecionados: {selecionados_count}</span>', unsafe_allow_html=True)
    with s4:
        st.markdown(f'<div class="cm-subtle">Itens atualmente no filtro: <b>{len(ids_filtrados)}</b></div>', unsafe_allow_html=True)

    scope = st.radio("Escopo da ação", ["Selecionados", "Todos do filtro"], horizontal=True)
    target_ids = list(dm0.index[dm0["SELECIONADO"].fillna(False)]) if scope == "Selecionados" else ids_filtrados

    st.markdown(f'<div class="cm-subtle">A ação será aplicada em <b>{len(target_ids)}</b> item(ns).</div>', unsafe_allow_html=True)

    bA, bB, bC, bD, bE, bF = st.columns([1.0, 1.0, 1.2, 1.5, 1.9, 1.0], gap="large")
    with bA:
        bulk_confirm = st.selectbox("Confirmado", ["(não alterar)", "Sim", "Não"])
    with bB:
        bulk_resolvido = st.selectbox("Resolvido", ["(não alterar)", "Sim", "Não"])
    with bC:
        bulk_status = st.selectbox("Status", ["(não alterar)"] + STATUS_OPTS)
    with bD:
        bulk_nucleo = st.selectbox("Núcleo", ["(não alterar)"] + get_nucleos())
    with bE:
        bulk_obs = st.text_input("OBS (opcional)", value="")
    with bF:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        do_apply = st.button("Aplicar", type="primary", disabled=(len(target_ids) == 0), use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

    if do_apply:
        dm = st.session_state.div_master.copy()

        if bulk_confirm != "(não alterar)":
            if bulk_confirm == "Sim":
                dm.loc[target_ids, "CONFIRMADO"] = True
                need = dm.loc[target_ids, "NUCLEO"].astype(str).str.strip().isin(["", "Não identificado"])
                idx_need = list(pd.Index(target_ids)[need.values])
                if len(idx_need) > 0:
                    dm.loc[idx_need, "NUCLEO"] = dm.loc[idx_need, "NUCLEO_SUGERIDO"].fillna("Não identificado")
            else:
                dm.loc[target_ids, "CONFIRMADO"] = False

        if bulk_nucleo != "(não alterar)":
            dm.loc[target_ids, "NUCLEO"] = bulk_nucleo
            dm.loc[target_ids, "CONFIRMADO"] = True

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
        save_learning_examples(dm.loc[target_ids].copy())
        dm["NUCLEO_EXIBICAO"] = get_nucleo_display_series(dm)
        st.session_state.div_master = dm
        set_flash("success", f"Ação aplicada em {len(target_ids)} itens.")
        st.rerun()

    # =====================================================
    # Tratativa
    # =====================================================
    st.markdown("### Tratativa (tabela)")
    st.markdown('<div class="cm-help">Mantido o histórico para contexto. A coluna operacional final continua sendo NÚCLEO.</div>', unsafe_allow_html=True)

    view_cols = [
        "SELECIONADO",
        "ORIGEM_VISUAL",
        "SEVERIDADE",
        "DATA",
        "DOCUMENTO",
        "HISTORICO_OPERACAO",
        "CHAVE_DOC",
        "VALOR",
        "CONFIRMADO",
        "NUCLEO",
        "STATUS",
        "RESOLVIDO",
        "OBS_USUARIO"
    ]
    df_view = df[view_cols].copy()
    df_view_display = df_view.copy()
    df_view_display["DATA"] = pd.to_datetime(df_view_display["DATA"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")

    column_config = {
        "SELECIONADO": st.column_config.CheckboxColumn("Selecionar"),
        "ORIGEM_VISUAL": st.column_config.TextColumn("Origem", disabled=True),
        "SEVERIDADE": st.column_config.TextColumn(disabled=True),
        "DATA": st.column_config.TextColumn(disabled=True),
        "DOCUMENTO": st.column_config.TextColumn(disabled=True),
        "HISTORICO_OPERACAO": st.column_config.TextColumn(disabled=True, width="large"),
        "CHAVE_DOC": st.column_config.TextColumn(disabled=True),
        "VALOR": st.column_config.NumberColumn(format="R$ %.2f", disabled=True),
        "CONFIRMADO": st.column_config.CheckboxColumn(),
        "NUCLEO": st.column_config.SelectboxColumn(options=get_nucleos()),
        "STATUS": st.column_config.SelectboxColumn(options=STATUS_OPTS),
        "RESOLVIDO": st.column_config.CheckboxColumn(),
        "OBS_USUARIO": st.column_config.TextColumn("Observação"),
    }

    edited = st.data_editor(
        df_view_display,
        use_container_width=True,
        height=580,
        column_config=column_config,
        key="editor_tratativa",
        hide_index=False,
    )

    if edited is not None and len(edited) == len(df_view_display):
        to_update = edited.copy()
        to_update["NUCLEO"] = to_update["NUCLEO"].fillna("Não identificado").replace("", "Não identificado")

        res_col = to_update["RESOLVIDO"].fillna(False)
        to_update.loc[res_col, "STATUS"] = "Resolvido"

        upd_cols = ["SELECIONADO", "CONFIRMADO", "NUCLEO", "STATUS", "RESOLVIDO", "OBS_USUARIO"]
        dm = st.session_state.div_master.copy()
        for c in upd_cols:
            dm.loc[to_update.index, c] = to_update[c].values

        need = dm.loc[to_update.index, "CONFIRMADO"].fillna(False) & dm.loc[to_update.index, "NUCLEO"].astype(str).str.strip().isin(["", "Não identificado"])
        idx_need = list(pd.Index(to_update.index)[need.values])
        if len(idx_need) > 0:
            dm.loc[idx_need, "NUCLEO"] = dm.loc[idx_need, "NUCLEO_SUGERIDO"].fillna("Não identificado")

        save_learning_examples(dm.loc[to_update.index].copy())
        dm["NUCLEO_EXIBICAO"] = get_nucleo_display_series(dm)
        dm["ORIGEM_VISUAL"] = dm["ORIGEM"].map(origem_visual_text)
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
  <div class="row"><span class="label">Histórico:</span> <span class="val">{r.get('HISTORICO_OPERACAO','')}</span></div>
  <div class="row"><span class="label">Núcleo sugerido:</span> <span class="val">{r.get('NUCLEO_SUGERIDO','')}</span></div>
  <div class="row"><span class="label">Núcleo final:</span> <span class="val">{r.get('NUCLEO','')}</span></div>
  <div class="row"><span class="label">Confirmado:</span> <span class="val">{confirmado_txt}</span></div>
  <div class="row"><span class="label">Status:</span> <span class="val">{r.get('STATUS','')}</span></div>
  <div class="row"><span class="label">Resolvido:</span> <span class="val">{resolvido_txt}</span></div>
</div>
""",
            unsafe_allow_html=True,
        )

        with st.expander("Ver detalhes técnicos do item", expanded=False):
            st.write({
                "MOTIVO_BASE": r.get("MOTIVO_BASE", ""),
                "REGRA_NUCLEO_APLICADA": r.get("REGRA_NUCLEO_APLICADA", ""),
                "REGRA_SEVERIDADE_APLICADA": r.get("REGRA_SEVERIDADE_APLICADA", ""),
            })

        resumo = (
            f"ID: {pick_id}\n"
            f"ORIGEM: {r.get('ORIGEM','')}\n"
            f"SEVERIDADE: {r.get('SEVERIDADE','')}\n"
            f"DATA: {dt_txt}\n"
            f"DOCUMENTO: {r.get('DOCUMENTO','')}\n"
            f"CHAVE: {r.get('CHAVE_DOC','')}\n"
            f"VALOR: {fmt(r.get('VALOR', np.nan))}\n"
            f"HISTORICO: {r.get('HISTORICO_OPERACAO','')}\n"
            f"NUCLEO: {r.get('NUCLEO','')}\n"
            f"CONFIRMADO: {confirmado_txt}\n"
            f"STATUS: {r.get('STATUS','')}\n"
            f"RESOLVIDO: {resolvido_txt}\n"
            f"OBS: {r.get('OBS_USUARIO','')}\n"
        )
        st.text_area("Copiar resumo (e-mail/ticket)", value=resumo, height=230)

    # =====================================================
    # Export
    # =====================================================
    st.markdown("### Exportar")
    filtros = {
        "origem": origem,
        "ver": ver,
        "severidade": sev,
        "nucleo": nucleo_filtro,
        "status": status_filtro,
        "busca": busca.strip(),
        "_total_aberto": valor_aberto
    }

    export_cols = [
        "ORIGEM",
        "DATA",
        "DOCUMENTO",
        "HISTORICO_OPERACAO",
        "CHAVE_DOC",
        "VALOR",
        "CONFIRMADO",
        "NUCLEO",
        "STATUS",
        "RESOLVIDO",
        "OBS_USUARIO"
    ]
    df_export = df[export_cols].copy()

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
