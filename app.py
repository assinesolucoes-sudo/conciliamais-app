import json
import re
import unicodedata
from collections import Counter
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Match Inteligente V22", layout="wide")

# =========================================================
# Utilidades básicas
# =========================================================

def _norm_text(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def _norm_name(x) -> str:
    s = _norm_text(x).lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _force_text_series(sr: pd.Series) -> pd.Series:
    return sr.fillna("").map(lambda x: "" if pd.isna(x) else str(x).strip())


def _normalize_key_series(sr: pd.Series) -> pd.Series:
    return _force_text_series(sr).map(lambda x: _norm_text(x).upper())


def _extract_digits(x) -> str:
    return re.sub(r"\D", "", _norm_text(x))


def _to_number(sr: pd.Series) -> pd.Series:
    if sr is None:
        return pd.Series(dtype=float)
    s = sr.copy()
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0.0)
    s = s.astype(str)
    s = s.str.replace(r"\s", "", regex=True)
    # formato brasileiro 1.234,56
    mask_br = s.str.contains(",", na=False)
    s.loc[mask_br] = s.loc[mask_br].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    # remove qualquer coisa que não seja número/sinal/ponto
    s = s.str.replace(r"[^0-9\-\.]", "", regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def _build_join_key(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    if not cols:
        return pd.Series(["__ALL__"] * len(df), index=df.index)
    parts = [_normalize_key_series(df[c]) for c in cols]
    out = parts[0].copy()
    for p in parts[1:]:
        out = out + "||" + p
    return out


def _safe_sheet_name(name: str) -> str:
    bad = r'[]:*?/\\'
    for ch in bad:
        name = name.replace(ch, "_")
    return name[:31]


def _suggest_label(a: str, b: str) -> str:
    na = _norm_name(a)
    nb = _norm_name(b)
    if na == nb and na:
        return _norm_text(a)

    keywords = [
        ("filial", "Filial"),
        ("patrimonio", "Patrimônio"),
        ("plaqueta", "Plaqueta"),
        ("conta", "Conta"),
        ("saldo", "Saldo"),
        ("aquisicao", "Aquisição"),
        ("depreciacao", "Depreciação"),
        ("depr acumul", "Depreciação"),
        ("depr acum", "Depreciação"),
        ("grupo patrimonio", "Descrição Grupo Patrimônio"),
        ("nome bem", "Nome Bem"),
    ]
    joined = f"{na} {nb}"
    for key, label in keywords:
        if key in joined:
            return label
    return f"{_norm_text(a)} ↔ {_norm_text(b)}"


def _parse_imported_rule(uploaded) -> List[dict]:
    if uploaded is None:
        return []
    raw = uploaded.getvalue()
    name = uploaded.name.lower()

    if name.endswith(".json"):
        payload = json.loads(raw.decode("utf-8", errors="ignore"))
        if isinstance(payload, dict) and "rules" in payload:
            payload = payload["rules"]
        if not isinstance(payload, list):
            raise ValueError("JSON da regra inválido.")
        out = []
        for item in payload:
            if not isinstance(item, dict):
                continue
            out.append({
                "source_col": item.get("source_col", ""),
                "target_col": item.get("target_col", ""),
                "mapping": {str(k): str(v) for k, v in (item.get("mapping", {}) or {}).items() if str(k).strip() != ""},
            })
        return out

    # CSV robusto - aceita o formato exportado pelo próprio sistema e versões quebradas anteriores
    text = raw.decode("utf-8-sig", errors="ignore")
    lines = [ln.strip().strip('"') for ln in text.splitlines() if ln.strip()]

    # Formato novo: SOURCE_COL,TARGET_COL,SOURCE_VALUE,TARGET_VALUE,USE
    if lines and "SOURCE_COL" in lines[0].upper():
        df = pd.read_csv(BytesIO(raw), dtype=str).fillna("")
        req = {"SOURCE_COL", "TARGET_COL", "SOURCE_VALUE", "TARGET_VALUE"}
        if req.issubset(set(df.columns)):
            grouped = {}
            for _, r in df.iterrows():
                if str(r.get("USE", "true")).strip().lower() not in {"true", "1", "sim", "yes", "y", ""}:
                    continue
                sc = _norm_text(r["SOURCE_COL"])
                tc = _norm_text(r["TARGET_COL"])
                sv = _norm_text(r["SOURCE_VALUE"])
                tv = _norm_text(r["TARGET_VALUE"])
                if not sc or not tc or not sv:
                    continue
                grouped.setdefault((sc, tc), {})[sv] = tv
            return [{"source_col": k[0], "target_col": k[1], "mapping": v} for k, v in grouped.items()]

    # Formato legado quebrado: USAR,VALOR_Ativos Rm,VALOR_Ativos Protheus
    pattern = re.compile(r'true,([^,\n\r]+),\s*"?([^"\n\r,]+)"?', flags=re.IGNORECASE)
    pairs = pattern.findall(text)
    if pairs:
        mapping = {}
        for src, tgt in pairs:
            src = _norm_text(src)
            tgt = _norm_text(tgt)
            if src:
                mapping[src] = tgt
        return [{"source_col": "", "target_col": "", "mapping": mapping}]

    raise ValueError("Arquivo CSV da regra não está no formato esperado.")


def _export_rules_payload(rules: List[dict]) -> Tuple[bytes, bytes]:
    payload = {"rules": rules}
    json_bytes = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    rows = []
    for rule in rules:
        sc = rule.get("source_col", "")
        tc = rule.get("target_col", "")
        for src, tgt in (rule.get("mapping", {}) or {}).items():
            rows.append({
                "SOURCE_COL": sc,
                "TARGET_COL": tc,
                "SOURCE_VALUE": src,
                "TARGET_VALUE": tgt,
                "USE": True,
            })
    csv_bytes = pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")
    return json_bytes, csv_bytes


# =========================================================
# Leitura de arquivos
# =========================================================

def read_any_table(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        # tenta ; depois ,
        raw = uploaded_file.getvalue()
        for sep in [";", ",", None]:
            try:
                if sep is None:
                    return pd.read_csv(BytesIO(raw), dtype=str, sep=None, engine="python")
                return pd.read_csv(BytesIO(raw), dtype=str, sep=sep)
            except Exception:
                continue
        raise ValueError(f"Não foi possível ler o CSV: {uploaded_file.name}")
    return pd.read_excel(BytesIO(uploaded_file.getvalue()), dtype=str)


# =========================================================
# Núcleo do produto
# =========================================================

def apply_rules_to_base1(df_a: pd.DataFrame, rules: List[dict]) -> Tuple[pd.DataFrame, List[dict]]:
    df = df_a.copy()
    created = []
    for rule in rules:
        source_col = rule.get("source_col", "")
        target_col = rule.get("target_col", "")
        mapping = rule.get("mapping", {}) or {}
        if not mapping:
            continue
        if source_col and source_col in df.columns:
            new_col = f"[MAP] {source_col} -> {target_col or 'Destino'}"
            source_series = _force_text_series(df[source_col])
            df[new_col] = source_series.map(lambda v: mapping.get(v, mapping.get(_norm_text(v), "")))
            created.append({
                "mapped_col": new_col,
                "source_col": source_col,
                "target_col": target_col,
            })
    return df, created


def build_pair_key_options(cols_a: List[str], cols_b: List[str], mapped_info: List[dict]) -> Tuple[List[str], Dict[str, str]]:
    # label amigável -> coluna real Base1 (ou campo mapeado)
    opts_a = list(cols_a)
    for m in mapped_info:
        if m["mapped_col"] not in opts_a:
            opts_a.append(m["mapped_col"])
    base1_map = {c: c for c in opts_a}
    return opts_a, base1_map


def compose_match(df_a: pd.DataFrame,
                  df_b: pd.DataFrame,
                  pairs: List[dict],
                  value_pairs: List[dict]) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    # pares: [{a_col, b_col, label}]
    # value_pairs: [{a_val, b_val, label}]

    work_a = df_a.copy()
    work_b = df_b.copy()

    # chaves lógicas para cada dimensão
    canonical_labels = []
    for p in pairs:
        label = p["label"]
        canonical_labels.append(label)
        work_a[f"KEY_{label}"] = _normalize_key_series(work_a[p["a_col"]]) if p["a_col"] in work_a.columns else ""
        work_b[f"KEY_{label}"] = _normalize_key_series(work_b[p["b_col"]]) if p["b_col"] in work_b.columns else ""

    if not canonical_labels:
        work_a["__JOIN__"] = "__ALL__"
        work_b["__JOIN__"] = "__ALL__"
    else:
        a_parts = [work_a[f"KEY_{lbl}"] for lbl in canonical_labels]
        b_parts = [work_b[f"KEY_{lbl}"] for lbl in canonical_labels]
        work_a["__JOIN__"] = a_parts[0]
        work_b["__JOIN__"] = b_parts[0]
        for i in range(1, len(canonical_labels)):
            work_a["__JOIN__"] = work_a["__JOIN__"] + "||" + a_parts[i]
            work_b["__JOIN__"] = work_b["__JOIN__"] + "||" + b_parts[i]

    # ocorrência dentro da chave para preservar repetidos
    work_a["__OCC__"] = work_a.groupby("__JOIN__").cumcount() + 1
    work_b["__OCC__"] = work_b.groupby("__JOIN__").cumcount() + 1

    a_dups = work_a.groupby("__JOIN__").size().reset_index(name="QTD_BASE1")
    b_dups = work_b.groupby("__JOIN__").size().reset_index(name="QTD_BASE2")
    dup_view = a_dups.merge(b_dups, on="__JOIN__", how="outer").fillna(0)
    dup_view["QTD_BASE1"] = dup_view["QTD_BASE1"].astype(int)
    dup_view["QTD_BASE2"] = dup_view["QTD_BASE2"].astype(int)
    dup_view["DUPLICIDADE"] = np.where((dup_view["QTD_BASE1"] > 1) | (dup_view["QTD_BASE2"] > 1), "Sim", "Não")

    merged = work_a.merge(
        work_b,
        on=["__JOIN__", "__OCC__"],
        how="outer",
        suffixes=("_BASE1", "_BASE2"),
        indicator=True,
    )

    # dimensões canônicas visíveis
    for p in pairs:
        lbl = p["label"]
        a_col = p["a_col"]
        b_col = p["b_col"]
        col_a = f"{a_col}_BASE1" if f"{a_col}_BASE1" in merged.columns else a_col
        col_b = f"{b_col}_BASE2" if f"{b_col}_BASE2" in merged.columns else b_col
        vis = pd.Series([""] * len(merged))
        if col_a in merged.columns:
            vis = _force_text_series(merged[col_a])
        if col_b in merged.columns:
            vis = vis.where(vis.ne(""), _force_text_series(merged[col_b]))
        merged[f"DIM_{lbl}"] = vis

    # metadados de presença e motivo
    merged["PRESENCA"] = merged["_merge"].map({
        "both": "Em ambas",
        "left_only": "Somente Base 1",
        "right_only": "Somente Base 2",
    })

    # Valores
    value_labels = []
    for vp in value_pairs:
        lbl = vp["label"]
        value_labels.append(lbl)
        a_col = vp["a_val"]
        b_col = vp["b_val"]
        col_a = f"{a_col}_BASE1" if f"{a_col}_BASE1" in merged.columns else a_col
        col_b = f"{b_col}_BASE2" if f"{b_col}_BASE2" in merged.columns else b_col
        a_num = _to_number(merged[col_a]) if col_a in merged.columns else pd.Series([0] * len(merged))
        b_num = _to_number(merged[col_b]) if col_b in merged.columns else pd.Series([0] * len(merged))
        merged[f"BASE1_{lbl}"] = a_num.round(2)
        merged[f"BASE2_{lbl}"] = b_num.round(2)
        merged[f"DIF_{lbl}"] = (a_num - b_num).round(2)

    if value_labels:
        dif_abs = sum(merged[f"DIF_{lbl}"].abs() for lbl in value_labels)
    else:
        dif_abs = pd.Series([0] * len(merged))

    merged["MOTIVO"] = np.select(
        [
            merged["PRESENCA"].eq("Somente Base 1"),
            merged["PRESENCA"].eq("Somente Base 2"),
            dif_abs.gt(0.0001),
        ],
        [
            "Somente Base 1",
            "Somente Base 2",
            "Divergência de valor",
        ],
        default="Conciliado",
    )

    # incorpora duplicidade
    merged = merged.merge(dup_view[["__JOIN__", "QTD_BASE1", "QTD_BASE2", "DUPLICIDADE"]], on="__JOIN__", how="left")
    merged["QTD_BASE1"] = merged["QTD_BASE1"].fillna(0).astype(int)
    merged["QTD_BASE2"] = merged["QTD_BASE2"].fillna(0).astype(int)
    merged["DUPLICIDADE"] = merged["DUPLICIDADE"].fillna("Não")
    merged.loc[merged["DUPLICIDADE"].eq("Sim"), "MOTIVO"] = np.where(
        merged.loc[merged["DUPLICIDADE"].eq("Sim"), "MOTIVO"].eq("Conciliado"),
        "Duplicidade",
        merged.loc[merged["DUPLICIDADE"].eq("Sim"), "MOTIVO"] + " + Duplicidade",
    )

    # subconjuntos úteis
    only_a = merged[merged["PRESENCA"].eq("Somente Base 1")].copy()
    only_b = merged[merged["PRESENCA"].eq("Somente Base 2")].copy()
    dups_only = merged[merged["DUPLICIDADE"].eq("Sim")].copy()

    return merged, only_a, only_b, dups_only


def build_executive_summary(full_df: pd.DataFrame,
                            group_dims: List[str],
                            value_labels: List[str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    # topo financeiro global
    top_rows = []
    for lbl in value_labels:
        top_rows.append({
            "Campo": lbl,
            "Total Base 1": round(full_df[f"BASE1_{lbl}"].sum(), 2),
            "Total Base 2": round(full_df[f"BASE2_{lbl}"].sum(), 2),
            "Diferença": round(full_df[f"DIF_{lbl}"].sum(), 2),
        })
    top_df = pd.DataFrame(top_rows)

    if not group_dims:
        group_dims = ["MOTIVO"]

    group_cols = [f"DIM_{g}" if not g.startswith("DIM_") else g for g in group_dims]
    existing_group_cols = [c for c in group_cols if c in full_df.columns]
    if not existing_group_cols:
        full_df = full_df.copy()
        full_df["DIM_Resumo"] = "Resumo Geral"
        existing_group_cols = ["DIM_Resumo"]

    agg_spec = {
        "PRESENCA": "count",
    }
    for lbl in value_labels:
        agg_spec[f"BASE1_{lbl}"] = "sum"
        agg_spec[f"BASE2_{lbl}"] = "sum"
        agg_spec[f"DIF_{lbl}"] = "sum"

    summary = full_df.groupby(existing_group_cols, dropna=False).agg(agg_spec).reset_index()
    summary = summary.rename(columns={"PRESENCA": "Qtde Registros"})

    # indicadores por motivo
    flags = pd.DataFrame(index=full_df.index)
    flags["Qtde Divergências"] = full_df["MOTIVO"].str.contains("Divergência", na=False).astype(int)
    flags["Qtde só Base 1"] = full_df["PRESENCA"].eq("Somente Base 1").astype(int)
    flags["Qtde só Base 2"] = full_df["PRESENCA"].eq("Somente Base 2").astype(int)
    flags["Qtde Duplicidades"] = full_df["DUPLICIDADE"].eq("Sim").astype(int)
    tmp = pd.concat([full_df[existing_group_cols], flags], axis=1)
    sums = tmp.groupby(existing_group_cols, dropna=False).sum().reset_index()
    summary = summary.merge(sums, on=existing_group_cols, how="left")

    # nomes mais amigáveis para agrupadores
    rename_map = {c: c.replace("DIM_", "") for c in existing_group_cols}
    summary = summary.rename(columns=rename_map)

    # ordena pelo maior impacto
    if value_labels:
        sort_cols = [f"DIF_{value_labels[0]}"]
        summary = summary.assign(__ABS__=summary[sort_cols[0]].abs()).sort_values("__ABS__", ascending=False).drop(columns=["__ABS__"])

    return top_df, summary


# =========================================================
# Exportação Excel
# =========================================================

def _autofit_and_format(writer, df: pd.DataFrame, sheet_name: str):
    wb = writer.book
    ws = writer.sheets[sheet_name]

    fmt_header = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
    fmt_text = wb.add_format({"border": 1})
    fmt_money = wb.add_format({"border": 1, "num_format": 'R$ #,##0.00'})
    fmt_int = wb.add_format({"border": 1, "num_format": '0'})

    for c_idx, col in enumerate(df.columns):
        ws.write(0, c_idx, col, fmt_header)
        ser = df[col]
        width = max(len(str(col)), min(60, ser.astype(str).map(len).max() if len(ser) else 10)) + 2
        moneyish = any(k in col.upper() for k in ["BASE 1", "BASE 2", "DIF", "SALDO", "AQUISI", "DEPRECIA"])
        qtyish = col.upper().startswith("QTDE") or col.upper().startswith("QTD")
        if qtyish:
            ws.set_column(c_idx, c_idx, width, fmt_int)
        elif moneyish and pd.api.types.is_numeric_dtype(ser):
            ws.set_column(c_idx, c_idx, width, fmt_money)
        else:
            ws.set_column(c_idx, c_idx, width, fmt_text)

    ws.autofilter(0, 0, len(df), max(0, len(df.columns) - 1))
    ws.freeze_panes(1, 0)


def export_excel(full_df: pd.DataFrame,
                 only_a: pd.DataFrame,
                 only_b: pd.DataFrame,
                 dups: pd.DataFrame,
                 top_exec: pd.DataFrame,
                 summary_exec: pd.DataFrame,
                 base1_name: str,
                 base2_name: str) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        top_exec.to_excel(writer, index=False, sheet_name="RESUMO_EXECUTIVO")
        summary_exec.to_excel(writer, index=False, sheet_name="RESUMO_POR_CHAVE", startrow=len(top_exec) + 3)
        full_df.to_excel(writer, index=False, sheet_name="RESULTADO_COMPLETO")
        only_a.to_excel(writer, index=False, sheet_name=_safe_sheet_name(f"SOMENTE_{base1_name}"))
        only_b.to_excel(writer, index=False, sheet_name=_safe_sheet_name(f"SOMENTE_{base2_name}"))
        dups.to_excel(writer, index=False, sheet_name="DUPLICIDADES")

        _autofit_and_format(writer, top_exec, "RESUMO_EXECUTIVO")
        _autofit_and_format(writer, summary_exec, "RESUMO_POR_CHAVE")
        _autofit_and_format(writer, full_df, "RESULTADO_COMPLETO")
        _autofit_and_format(writer, only_a, _safe_sheet_name(f"SOMENTE_{base1_name}"))
        _autofit_and_format(writer, only_b, _safe_sheet_name(f"SOMENTE_{base2_name}"))
        _autofit_and_format(writer, dups, "DUPLICIDADES")

    return out.getvalue()


# =========================================================
# Estado inicial
# =========================================================

def init_state():
    defaults = {
        "v22_rules": [],
        "v22_current_mapping": {},
        "v22_loaded_rule_preview": [],
        "v22_created_mapped": [],
        "v22_pair_rows": [{"a_col": "", "b_col": "", "label": ""}],
        "v22_value_rows": [{"a_val": "", "b_val": "", "label": ""}],
        "v22_base1_name": "Base 1",
        "v22_base2_name": "Base 2",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# =========================================================
# Interface
# =========================================================

def main():
    init_state()
    st.title("Match Inteligente V22")
    st.caption("Motor genérico de confronto entre duas bases, com regra de equivalência, resumo executivo e resultado auditável.")

    # --- Entrada de bases
    st.subheader("1) Carregar bases")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state["v22_base1_name"] = st.text_input("Nome da Base 1", value=st.session_state["v22_base1_name"])
        up_a = st.file_uploader("Arquivo da Base 1", type=["xlsx", "xls", "csv"], key="v22_up_a")
    with c2:
        st.session_state["v22_base2_name"] = st.text_input("Nome da Base 2", value=st.session_state["v22_base2_name"])
        up_b = st.file_uploader("Arquivo da Base 2", type=["xlsx", "xls", "csv"], key="v22_up_b")

    if not up_a or not up_b:
        st.info("Carregue as duas bases para continuar.")
        return

    df_a = read_any_table(up_a)
    df_b = read_any_table(up_b)
    cols_a = list(df_a.columns)
    cols_b = list(df_b.columns)

    st.success(f"Base 1: {len(df_a):,} linhas | Base 2: {len(df_b):,} linhas".replace(",", "."))

    # --- Regras de equivalência
    st.subheader("2) Regras opcionais de equivalência / de-para")
    st.caption("Use quando um campo da Base 1 precisa ser traduzido para corresponder a um campo da Base 2.")

    rc1, rc2, rc3 = st.columns([1.3, 1.3, 1])
    with rc1:
        rule_source = st.selectbox("Campo original da Base 1", options=[""] + cols_a, key="v22_rule_source")
    with rc2:
        rule_target = st.selectbox("Campo correspondente da Base 2", options=[""] + cols_b, key="v22_rule_target")
    with rc3:
        load_assoc = st.button("Carregar associação")

    if load_assoc:
        if not rule_source:
            st.warning("Selecione o campo original da Base 1.")
        else:
            vals = sorted([v for v in _force_text_series(df_a[rule_source]).unique().tolist() if _norm_text(v) != ""])
            st.session_state["v22_current_mapping"] = {v: "" for v in vals}

    # importar regra
    import_file = st.file_uploader("Importar regra (.json ou .csv)", type=["json", "csv"], key="v22_rule_import")
    if import_file is not None:
        try:
            parsed = _parse_imported_rule(import_file)
            st.session_state["v22_loaded_rule_preview"] = parsed
            st.success(f"Regra importada com {len(parsed)} bloco(s).")
        except Exception as e:
            st.error(f"Falha ao importar regra: {e}")

    if st.session_state.get("v22_loaded_rule_preview"):
        if st.button("Aplicar regra importada"):
            for rule in st.session_state["v22_loaded_rule_preview"]:
                st.session_state["v22_rules"].append(rule)
            st.success("Regra importada aplicada.")
            st.session_state["v22_loaded_rule_preview"] = []

    current_map = st.session_state.get("v22_current_mapping", {})
    if current_map:
        st.markdown("**Montagem da associação atual**")
        dest_options = [""] + cols_b
        default_target_index = dest_options.index(rule_target) if rule_target in dest_options else 0

        edited_mapping = {}
        for i, src_val in enumerate(list(current_map.keys())):
            cc1, cc2, cc3 = st.columns([1.5, 1.2, 0.3])
            with cc1:
                st.text_input(f"Base 1 #{i+1}", value=src_val, disabled=True, key=f"v22_src_{i}")
            with cc2:
                edited_mapping[src_val] = st.selectbox(
                    f"Base 2 #{i+1}",
                    options=dest_options,
                    index=dest_options.index(current_map.get(src_val, "")) if current_map.get(src_val, "") in dest_options else default_target_index,
                    key=f"v22_tgt_{i}",
                )
            with cc3:
                pass

        st.session_state["v22_current_mapping"] = edited_mapping
        ac1, ac2, ac3 = st.columns([1, 1, 2])
        with ac1:
            if st.button("Confirmar regra atual"):
                clean_map = {k: v for k, v in edited_mapping.items() if _norm_text(v) != ""}
                if not clean_map:
                    st.warning("Nenhuma associação preenchida para confirmar.")
                else:
                    st.session_state["v22_rules"].append({
                        "source_col": rule_source,
                        "target_col": rule_target,
                        "mapping": clean_map,
                    })
                    st.session_state["v22_current_mapping"] = {}
                    st.success("Regra adicionada.")
        with ac2:
            if st.button("Cancelar regra atual"):
                st.session_state["v22_current_mapping"] = {}

    if st.session_state["v22_rules"]:
        st.markdown("**Regras confirmadas**")
        for idx, r in enumerate(st.session_state["v22_rules"]):
            st.write(f"{idx+1}. {r.get('source_col','(livre)')} -> {r.get('target_col','(livre)')} | {len(r.get('mapping',{}))} associação(ões)")
        jbytes, cbytes = _export_rules_payload(st.session_state["v22_rules"])
        dc1, dc2 = st.columns(2)
        with dc1:
            st.download_button("Baixar regra em JSON", data=jbytes, file_name="regra_equivalencia.json", mime="application/json")
        with dc2:
            st.download_button("Baixar regra em CSV", data=cbytes, file_name="regra_equivalencia.csv", mime="text/csv")

    # aplica regras na Base 1
    df_a_mapped, mapped_info = apply_rules_to_base1(df_a, st.session_state["v22_rules"])
    opts_a, _ = build_pair_key_options(cols_a, cols_b, mapped_info)

    # --- Campos identificadores
    st.subheader("3) Campos que identificam o mesmo registro nas duas bases")
    st.caption("Monte a chave lógica da análise. Pode ser simples ou composta.")

    pair_rows = st.session_state["v22_pair_rows"]
    for i, row in enumerate(pair_rows):
        cpa, cpb, cpl = st.columns([1.2, 1.2, 1])
        with cpa:
            a_choice = st.selectbox(
                f"Base 1 #{i+1}",
                options=[""] + opts_a,
                index=([""] + opts_a).index(row.get("a_col", "")) if row.get("a_col", "") in ([""] + opts_a) else 0,
                key=f"v22_pair_a_{i}",
            )
        with cpb:
            b_options = [""] + cols_b
            # se Base1 for [MAP], sugerir alvo da regra mas sem prender o usuário
            suggested = row.get("b_col", "")
            if a_choice.startswith("[MAP]"):
                hit = next((m for m in mapped_info if m["mapped_col"] == a_choice), None)
                if hit and not suggested:
                    suggested = hit.get("target_col", "")
            b_choice = st.selectbox(
                f"Base 2 #{i+1}",
                options=b_options,
                index=b_options.index(suggested) if suggested in b_options else 0,
                key=f"v22_pair_b_{i}",
            )
        with cpl:
            default_label = row.get("label", "") or (_suggest_label(a_choice, b_choice) if a_choice or b_choice else "")
            label = st.text_input(f"Nome da dimensão #{i+1}", value=default_label, key=f"v22_pair_lbl_{i}")

        pair_rows[i] = {"a_col": a_choice, "b_col": b_choice, "label": label}

    p1, p2 = st.columns([1, 4])
    with p1:
        if st.button("Adicionar campo identificador"):
            pair_rows.append({"a_col": "", "b_col": "", "label": ""})
    st.session_state["v22_pair_rows"] = pair_rows

    # --- Campos de valor
    st.subheader("4) Quais campos deseja confrontar para validar valores")
    value_rows = st.session_state["v22_value_rows"]
    for i, row in enumerate(value_rows):
        cva, cvb, cvl = st.columns([1.2, 1.2, 1])
        with cva:
            a_choice = st.selectbox(
                f"Valor Base 1 #{i+1}",
                options=[""] + cols_a,
                index=([""] + cols_a).index(row.get("a_val", "")) if row.get("a_val", "") in ([""] + cols_a) else 0,
                key=f"v22_val_a_{i}",
            )
        with cvb:
            b_choice = st.selectbox(
                f"Valor Base 2 #{i+1}",
                options=[""] + cols_b,
                index=([""] + cols_b).index(row.get("b_val", "")) if row.get("b_val", "") in ([""] + cols_b) else 0,
                key=f"v22_val_b_{i}",
            )
        with cvl:
            default_label = row.get("label", "") or (_suggest_label(a_choice, b_choice) if a_choice or b_choice else "")
            label = st.text_input(f"Nome do valor #{i+1}", value=default_label, key=f"v22_val_lbl_{i}")
        value_rows[i] = {"a_val": a_choice, "b_val": b_choice, "label": label}

    if st.button("Adicionar campo de valor"):
        value_rows.append({"a_val": "", "b_val": "", "label": ""})
    st.session_state["v22_value_rows"] = value_rows

    # --- Resumo executivo
    st.subheader("5) Resumo executivo")
    gen_exec = st.checkbox("Gerar resumo executivo", value=True)

    valid_pair_rows = [r for r in pair_rows if r.get("a_col") and (r.get("b_col") or str(r.get("a_col", "")).startswith("[MAP]"))]
    valid_value_rows = [r for r in value_rows if r.get("a_val") and r.get("b_val")]

    default_group_labels = [r.get("label") or _suggest_label(r["a_col"], r.get("b_col", "")) for r in valid_pair_rows]
    default_value_labels = [r.get("label") or _suggest_label(r["a_val"], r["b_val"]) for r in valid_value_rows]

    group_choices = []
    value_choices = []
    if gen_exec:
        gc1, gc2 = st.columns(2)
        with gc1:
            group_choices = st.multiselect(
                "Agrupar resumo por",
                options=default_group_labels,
                default=default_group_labels,
                key="v22_group_choices",
            )
        with gc2:
            value_choices = st.multiselect(
                "O que deseja totalizar/confrontar",
                options=default_value_labels,
                default=default_value_labels,
                key="v22_value_choices",
            )

    # --- Executar
    st.subheader("6) Processar")
    if st.button("Executar análise"):
        if not valid_pair_rows:
            st.error("Adicione pelo menos um campo identificador válido.")
            return
        if not valid_value_rows:
            st.error("Adicione pelo menos um campo de valor válido.")
            return

        # completa labels faltantes
        for r in valid_pair_rows:
            if not r.get("label"):
                r["label"] = _suggest_label(r["a_col"], r.get("b_col", ""))
        for r in valid_value_rows:
            if not r.get("label"):
                r["label"] = _suggest_label(r["a_val"], r["b_val"])

        full_df, only_a, only_b, dups = compose_match(df_a_mapped, df_b, valid_pair_rows, valid_value_rows)

        # Enriquecimento quando houver [MAP] para rastreabilidade
        for r in valid_pair_rows:
            a_col = r["a_col"]
            if a_col.startswith("[MAP]"):
                hit = next((m for m in mapped_info if m["mapped_col"] == a_col), None)
                if hit:
                    source_col = hit["source_col"]
                    target_col = hit["target_col"]
                    src_col_full = f"{source_col}_BASE1" if f"{source_col}_BASE1" in full_df.columns else source_col
                    map_col_full = f"{a_col}_BASE1" if f"{a_col}_BASE1" in full_df.columns else a_col
                    tgt_col_full = f"{target_col}_BASE2" if target_col and f"{target_col}_BASE2" in full_df.columns else target_col
                    if src_col_full in full_df.columns:
                        full_df[f"BASE1_ORIGINAL_{r['label']}"] = _force_text_series(full_df[src_col_full])
                    if map_col_full in full_df.columns:
                        full_df[f"BASE1_MAPEADO_{r['label']}"] = _force_text_series(full_df[map_col_full])
                    if tgt_col_full and tgt_col_full in full_df.columns:
                        full_df[f"BASE2_CORRESP_{r['label']}"] = _force_text_series(full_df[tgt_col_full])

        selected_group = group_choices or default_group_labels
        selected_values = value_choices or default_value_labels
        top_exec, summary_exec = build_executive_summary(full_df, selected_group, selected_values)

        excel_bytes = export_excel(
            full_df=full_df,
            only_a=only_a,
            only_b=only_b,
            dups=dups,
            top_exec=top_exec,
            summary_exec=summary_exec,
            base1_name=st.session_state["v22_base1_name"],
            base2_name=st.session_state["v22_base2_name"],
        )

        st.success("Análise concluída.")
        st.markdown("**Topo executivo**")
        st.dataframe(top_exec, use_container_width=True)
        st.markdown("**Resumo por chave**")
        st.dataframe(summary_exec, use_container_width=True)
        st.markdown("**Resultado completo**")
        st.dataframe(full_df.head(200), use_container_width=True)

        st.download_button(
            "Baixar Excel da análise",
            data=excel_bytes,
            file_name="Match_Inteligente_V22.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
