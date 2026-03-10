import json
import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Concilia Mais - Match Inteligente V23", layout="wide")

# ============================================================
# Helpers
# ============================================================

def _clean_text(x) -> str:
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x).strip())


def _norm_name(x: str) -> str:
    s = _clean_text(x).lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _to_text(sr: pd.Series) -> pd.Series:
    return sr.fillna("").astype(str).map(_clean_text)


def _to_key(sr: pd.Series) -> pd.Series:
    return _to_text(sr).str.upper()


def _to_number(sr: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(sr):
        return pd.to_numeric(sr, errors="coerce").fillna(0.0)
    s = sr.fillna("").astype(str).str.strip()
    s = s.str.replace(r"\s", "", regex=True)
    mask_br = s.str.contains(",", na=False)
    s.loc[mask_br] = s.loc[mask_br].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    s = s.str.replace(r"[^0-9\-\.]", "", regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def _safe_sheet(name: str) -> str:
    for ch in r'[]:*?/\\':
        name = name.replace(ch, "_")
    return name[:31]


def _friendly_label(a: str, b: str) -> str:
    na = _norm_name(a)
    nb = _norm_name(b)
    joined = f"{na} {nb}"
    mapping = [
        ("filial", "Filial"),
        ("patrimonio", "Patrimônio"),
        ("plaqueta", "Plaqueta"),
        ("conta", "Conta Contábil"),
        ("saldo", "Saldo"),
        ("aquisicao", "Aquisição"),
        ("depreciacao", "Depreciação"),
        ("grupo patrimonio", "Grupo Patrimônio"),
        ("nome bem", "Nome Bem"),
    ]
    for key, label in mapping:
        if key in joined:
            return label
    if _clean_text(a) == _clean_text(b) and _clean_text(a):
        return _clean_text(a)
    return f"{_clean_text(a)} ↔ {_clean_text(b)}"


def _read_file(uploaded) -> pd.DataFrame:
    name = uploaded.name.lower()
    raw = uploaded.getvalue()
    if name.endswith(".csv"):
        for sep in [";", ",", None]:
            try:
                if sep is None:
                    return pd.read_csv(BytesIO(raw), dtype=str, sep=None, engine="python")
                return pd.read_csv(BytesIO(raw), dtype=str, sep=sep)
            except Exception:
                pass
        raise ValueError(f"Não foi possível ler o CSV: {uploaded.name}")
    return pd.read_excel(BytesIO(raw), dtype=str)


def _build_key(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    if not cols:
        return pd.Series(["__ALL__"] * len(df), index=df.index)
    out = _to_key(df[cols[0]])
    for c in cols[1:]:
        out = out + "||" + _to_key(df[c])
    return out


def _export_rule_files(rules: List[dict]) -> Tuple[bytes, bytes]:
    json_bytes = json.dumps({"rules": rules}, ensure_ascii=False, indent=2).encode("utf-8")
    rows = []
    for r in rules:
        for k, v in r.get("mapping", {}).items():
            rows.append(
                {
                    "SOURCE_COL": r.get("source_col", ""),
                    "TARGET_COL": r.get("target_col", ""),
                    "SOURCE_VALUE": k,
                    "TARGET_VALUE": v,
                    "USE": True,
                }
            )
    csv_bytes = pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")
    return json_bytes, csv_bytes


def _parse_rule_upload(uploaded) -> List[dict]:
    if uploaded is None:
        return []
    name = uploaded.name.lower()
    raw = uploaded.getvalue()
    if name.endswith(".json"):
        payload = json.loads(raw.decode("utf-8", errors="ignore"))
        rules = payload.get("rules", payload if isinstance(payload, list) else [])
        out = []
        for r in rules:
            out.append(
                {
                    "source_col": r.get("source_col", ""),
                    "target_col": r.get("target_col", ""),
                    "mapping": {str(k): str(v) for k, v in (r.get("mapping", {}) or {}).items()},
                }
            )
        return out

    df = pd.read_csv(BytesIO(raw), dtype=str).fillna("")
    req = {"SOURCE_COL", "TARGET_COL", "SOURCE_VALUE", "TARGET_VALUE"}
    if not req.issubset(df.columns):
        raise ValueError("CSV da regra não está no formato esperado.")
    grouped: Dict[Tuple[str, str], Dict[str, str]] = {}
    for _, row in df.iterrows():
        sc = _clean_text(row["SOURCE_COL"])
        tc = _clean_text(row["TARGET_COL"])
        sv = _clean_text(row["SOURCE_VALUE"])
        tv = _clean_text(row["TARGET_VALUE"])
        if not sv:
            continue
        grouped.setdefault((sc, tc), {})[sv] = tv
    return [{"source_col": k[0], "target_col": k[1], "mapping": v} for k, v in grouped.items()]


# ============================================================
# Estado
# ============================================================

def _init_state():
    defaults = {
        "cm_base1_name": "Base 1",
        "cm_base2_name": "Base 2",
        "cm_rules": [],
        "cm_rule_preview": [],
        "cm_live_mapping": {},
        "cm_key_rows": [{"a": "", "b": "", "label": ""}],
        "cm_val_rows": [{"a": "", "b": "", "label": ""}],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ============================================================
# Regras / campos mapeados
# ============================================================

def _apply_rules_to_base1(df: pd.DataFrame, rules: List[dict]) -> Tuple[pd.DataFrame, List[dict]]:
    out = df.copy()
    meta = []
    for rule in rules:
        sc = rule.get("source_col", "")
        tc = rule.get("target_col", "")
        mapping = rule.get("mapping", {}) or {}
        if not sc or sc not in out.columns or not mapping:
            continue
        mapped_col = f"[MAP] {sc} -> {tc or 'Destino'}"
        original = _to_text(out[sc])
        out[mapped_col] = original.map(lambda x: mapping.get(x, mapping.get(_clean_text(x), "")))
        meta.append(
            {
                "mapped_col": mapped_col,
                "source_col": sc,
                "target_col": tc,
            }
        )
    return out, meta


# ============================================================
# Motor de conciliação
# ============================================================

def _run_reconciliation(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    key_pairs: List[dict],
    value_pairs: List[dict],
    mapped_meta: List[dict],
) -> Dict[str, pd.DataFrame]:
    a = df_a.copy()
    b = df_b.copy()

    key_labels = []
    for kp in key_pairs:
        lbl = kp["label"]
        key_labels.append(lbl)
        a[f"KEY::{lbl}"] = _to_key(a[kp["a"]]) if kp["a"] in a.columns else ""
        b[f"KEY::{lbl}"] = _to_key(b[kp["b"]]) if kp["b"] in b.columns else ""

    a["__MATCH_KEY__"] = _build_key(a, [f"KEY::{lbl}" for lbl in key_labels])
    b["__MATCH_KEY__"] = _build_key(b, [f"KEY::{lbl}" for lbl in key_labels])

    a["__OCC__"] = a.groupby("__MATCH_KEY__").cumcount() + 1
    b["__OCC__"] = b.groupby("__MATCH_KEY__").cumcount() + 1

    merged = a.merge(
        b,
        on=["__MATCH_KEY__", "__OCC__"],
        how="outer",
        suffixes=("_A", "_B"),
        indicator=True,
    )

    # dimensões executivas canônicas
    for kp in key_pairs:
        label = kp["label"]
        a_col = f"{kp['a']}_A" if f"{kp['a']}_A" in merged.columns else kp["a"]
        b_col = f"{kp['b']}_B" if f"{kp['b']}_B" in merged.columns else kp["b"]
        s = pd.Series([""] * len(merged), index=merged.index)
        if a_col in merged.columns:
            s = _to_text(merged[a_col])
        if b_col in merged.columns:
            s = s.where(s.ne(""), _to_text(merged[b_col]))
        merged[f"DIM::{label}"] = s

    # valores confrontados
    value_labels = []
    for vp in value_pairs:
        lbl = vp["label"]
        value_labels.append(lbl)
        a_col = f"{vp['a']}_A" if f"{vp['a']}_A" in merged.columns else vp["a"]
        b_col = f"{vp['b']}_B" if f"{vp['b']}_B" in merged.columns else vp["b"]
        aval = _to_number(merged[a_col]) if a_col in merged.columns else pd.Series([0.0] * len(merged))
        bval = _to_number(merged[b_col]) if b_col in merged.columns else pd.Series([0.0] * len(merged))
        merged[f"VALOR::{lbl}::A"] = aval.round(2)
        merged[f"VALOR::{lbl}::B"] = bval.round(2)
        merged[f"DIF::{lbl}"] = (aval - bval).round(2)

    # motivo
    merged["PRESENCA"] = merged["_merge"].map(
        {"both": "Em ambas", "left_only": "Somente Base 1", "right_only": "Somente Base 2"}
    )

    dup_a = a.groupby("__MATCH_KEY__").size().rename("QTD_A")
    dup_b = b.groupby("__MATCH_KEY__").size().rename("QTD_B")
    dup = dup_a.to_frame().join(dup_b, how="outer").fillna(0).astype(int).reset_index()
    dup["DUPLICIDADE"] = np.where((dup["QTD_A"] > 1) | (dup["QTD_B"] > 1), 1, 0)
    merged = merged.merge(dup[["__MATCH_KEY__", "QTD_A", "QTD_B", "DUPLICIDADE"]], on="__MATCH_KEY__", how="left")
    merged["QTD_A"] = merged["QTD_A"].fillna(0).astype(int)
    merged["QTD_B"] = merged["QTD_B"].fillna(0).astype(int)
    merged["DUPLICIDADE"] = merged["DUPLICIDADE"].fillna(0).astype(int)

    any_diff = pd.Series([0.0] * len(merged), index=merged.index)
    for lbl in value_labels:
        any_diff = any_diff + merged[f"DIF::{lbl}"].abs()

    merged["MOTIVO"] = np.select(
        [
            merged["PRESENCA"].eq("Somente Base 1"),
            merged["PRESENCA"].eq("Somente Base 2"),
            merged["DUPLICIDADE"].eq(1),
            any_diff.gt(0.0001),
        ],
        [
            "Chave só na Base 1",
            "Chave só na Base 2",
            "Duplicidade",
            "Valor divergente",
        ],
        default="Conciliado",
    )

    # rastreabilidade do mapeamento
    for meta in mapped_meta:
        mapped_col = meta["mapped_col"]
        source_col = meta["source_col"]
        target_col = meta["target_col"]
        src_full = f"{source_col}_A" if f"{source_col}_A" in merged.columns else source_col
        map_full = f"{mapped_col}_A" if f"{mapped_col}_A" in merged.columns else mapped_col
        tgt_full = f"{target_col}_B" if f"{target_col}_B" in merged.columns else target_col
        if src_full in merged.columns:
            merged[f"BASE1_ORIGINAL::{source_col}"] = _to_text(merged[src_full])
        if map_full in merged.columns:
            merged[f"BASE1_MAPEADO::{target_col}"] = _to_text(merged[map_full])
        if tgt_full in merged.columns:
            merged[f"BASE2_CORRESP::{target_col}"] = _to_text(merged[tgt_full])

    # resumo global
    global_rows = []
    for lbl in value_labels:
        total_a = round(merged[f"VALOR::{lbl}::A"].sum(), 2)
        total_b = round(merged[f"VALOR::{lbl}::B"].sum(), 2)
        global_rows.append({
            "Campo": lbl,
            f"Total {st.session_state['cm_base1_name']}": total_a,
            f"Total {st.session_state['cm_base2_name']}": total_b,
            "Diferença total": round(total_a - total_b, 2),
        })
    resumo_global = pd.DataFrame(global_rows)

    return {
        "full": merged,
        "resumo_global": resumo_global,
        "value_labels": value_labels,
        "key_labels": key_labels,
    }


def _build_executive_and_detail(results: Dict[str, pd.DataFrame], group_labels: List[str], value_labels: List[str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df = results["full"].copy()

    if not group_labels:
        group_labels = results["key_labels"][:1] if results["key_labels"] else ["MOTIVO"]

    group_cols = []
    for lbl in group_labels:
        col = f"DIM::{lbl}" if f"DIM::{lbl}" in df.columns else lbl
        if col in df.columns:
            group_cols.append(col)
    if not group_cols:
        df["DIM::Resumo"] = "Resumo Geral"
        group_cols = ["DIM::Resumo"]

    agg = {"MOTIVO": "count"}
    for lbl in value_labels:
        agg[f"VALOR::{lbl}::A"] = "sum"
        agg[f"VALOR::{lbl}::B"] = "sum"
        agg[f"DIF::{lbl}"] = "sum"

    exec_df = df.groupby(group_cols, dropna=False).agg(agg).reset_index()
    exec_df = exec_df.rename(columns={"MOTIVO": "Qtde Registros"})

    # composição da diferença
    flags = pd.DataFrame(index=df.index)
    flags["Qtde Valor divergente"] = df["MOTIVO"].eq("Valor divergente").astype(int)
    flags[f"Qtde só {st.session_state['cm_base1_name']}"] = df["MOTIVO"].eq("Chave só na Base 1").astype(int)
    flags[f"Qtde só {st.session_state['cm_base2_name']}"] = df["MOTIVO"].eq("Chave só na Base 2").astype(int)
    flags["Qtde Duplicidades"] = df["MOTIVO"].eq("Duplicidade").astype(int)
    flags["Motivo predominante"] = df["MOTIVO"]
    flag_df = pd.concat([df[group_cols], flags], axis=1)

    summed = flag_df.groupby(group_cols, dropna=False).sum(numeric_only=True).reset_index()
    exec_df = exec_df.merge(summed, on=group_cols, how="left")

    motive = (
        flag_df.groupby(group_cols + ["Motivo predominante"], dropna=False)
        .size()
        .reset_index(name="QTD")
        .sort_values([*group_cols, "QTD"], ascending=[True] * len(group_cols) + [False])
        .drop_duplicates(group_cols)
        .drop(columns=["QTD"])
    )
    exec_df = exec_df.merge(motive, on=group_cols, how="left")

    # nomes amigáveis
    rename_map = {c: c.replace("DIM::", "") for c in group_cols}
    for lbl in value_labels:
        rename_map[f"VALOR::{lbl}::A"] = f"{lbl} {st.session_state['cm_base1_name']}"
        rename_map[f"VALOR::{lbl}::B"] = f"{lbl} {st.session_state['cm_base2_name']}"
        rename_map[f"DIF::{lbl}"] = f"Diferença {lbl}"
    exec_df = exec_df.rename(columns=rename_map)

    # detalhe das diferenças apenas
    detail_cols = group_cols.copy()
    for lbl in results["key_labels"]:
        dim = f"DIM::{lbl}"
        if dim in df.columns and dim not in detail_cols:
            detail_cols.append(dim)

    extra_cols = [c for c in df.columns if c.startswith("BASE1_ORIGINAL::") or c.startswith("BASE1_MAPEADO::") or c.startswith("BASE2_CORRESP::")]
    val_cols = []
    for lbl in value_labels:
        val_cols.extend([f"VALOR::{lbl}::A", f"VALOR::{lbl}::B", f"DIF::{lbl}"])

    detail = df[df["MOTIVO"].ne("Conciliado")].copy()
    detail = detail[detail_cols + extra_cols + val_cols + ["MOTIVO"]].copy()
    detail = detail.rename(columns=rename_map)
    for lbl in value_labels:
        detail = detail.rename(columns={
            f"VALOR::{lbl}::A": f"{lbl} {st.session_state['cm_base1_name']}",
            f"VALOR::{lbl}::B": f"{lbl} {st.session_state['cm_base2_name']}",
            f"DIF::{lbl}": f"Diferença {lbl}",
        })

    # ordena pela maior diferença do primeiro campo
    if value_labels:
        diff_col = f"Diferença {value_labels[0]}"
        if diff_col in exec_df.columns:
            exec_df = exec_df.assign(__ABS__=exec_df[diff_col].abs()).sort_values("__ABS__", ascending=False).drop(columns=["__ABS__"])
        if diff_col in detail.columns:
            detail = detail.assign(__ABS__=detail[diff_col].abs()).sort_values("__ABS__", ascending=False).drop(columns=["__ABS__"])

    return exec_df, detail


# ============================================================
# Excel
# ============================================================

def _fmt_sheet(writer, sheet_name: str, df: pd.DataFrame):
    wb = writer.book
    ws = writer.sheets[sheet_name]
    fmt_head = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "align": "center"})
    fmt_text = wb.add_format({"border": 1})
    fmt_money = wb.add_format({"border": 1, "num_format": 'R$ #,##0.00'})
    fmt_int = wb.add_format({"border": 1, "num_format": '0'})
    fmt_diff = wb.add_format({"border": 1, "num_format": 'R$ #,##0.00', "font_color": "#C00000", "bold": True})

    for i, col in enumerate(df.columns):
        ws.write(0, i, col, fmt_head)
        ser = df[col]
        width = max(len(str(col)), min(50, ser.astype(str).map(len).max() if len(ser) else 10)) + 2
        uc = str(col).upper()
        if uc.startswith("QTDE") or uc.startswith("QTD"):
            ws.set_column(i, i, width, fmt_int)
        elif "DIFERENÇA" in uc and pd.api.types.is_numeric_dtype(ser):
            ws.set_column(i, i, width, fmt_diff)
        elif (st.session_state['cm_base1_name'].upper() in uc or st.session_state['cm_base2_name'].upper() in uc) and pd.api.types.is_numeric_dtype(ser):
            ws.set_column(i, i, width, fmt_money)
        elif pd.api.types.is_numeric_dtype(ser):
            ws.set_column(i, i, width, fmt_money)
        else:
            ws.set_column(i, i, width, fmt_text)

    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(df), max(0, len(df.columns) - 1))


def _export_excel(resumo_global: pd.DataFrame, resumo_exec: pd.DataFrame, detalhe: pd.DataFrame, rules_df: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        resumo_global.to_excel(writer, sheet_name="RESUMO_EXECUTIVO", index=False, startrow=0)
        resumo_exec.to_excel(writer, sheet_name="RESUMO_EXECUTIVO", index=False, startrow=len(resumo_global) + 3)
        detalhe.to_excel(writer, sheet_name="DETALHE_DIFERENCAS", index=False)
        if len(rules_df):
            rules_df.to_excel(writer, sheet_name="REGRAS_APLICADAS", index=False)

        _fmt_sheet(writer, "RESUMO_EXECUTIVO", pd.concat([resumo_global, resumo_exec], axis=0, ignore_index=True))
        _fmt_sheet(writer, "DETALHE_DIFERENCAS", detalhe)
        if len(rules_df):
            _fmt_sheet(writer, "REGRAS_APLICADAS", rules_df)

    return bio.getvalue()


# ============================================================
# App
# ============================================================

def main():
    _init_state()
    st.title("Concilia Mais - Match Inteligente V23")
    st.caption("Ferramenta de conciliação entre duas bases: primeiro fecha o total, depois mostra onde está a diferença e por quê.")

    # 1. Bases
    st.subheader("1) Bases da análise")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state["cm_base1_name"] = st.text_input("Nome da Base 1", st.session_state["cm_base1_name"])
        up_a = st.file_uploader("Arquivo Base 1", type=["xlsx", "xls", "csv"], key="cm_up_a")
    with c2:
        st.session_state["cm_base2_name"] = st.text_input("Nome da Base 2", st.session_state["cm_base2_name"])
        up_b = st.file_uploader("Arquivo Base 2", type=["xlsx", "xls", "csv"], key="cm_up_b")

    if not up_a or not up_b:
        st.info("Carregue as duas bases para continuar.")
        return

    df_a = _read_file(up_a)
    df_b = _read_file(up_b)
    cols_a = list(df_a.columns)
    cols_b = list(df_b.columns)

    # 2. Regras opcionais
    st.subheader("2) Regras opcionais de equivalência")
    r1, r2, r3 = st.columns([1.2, 1.2, 0.8])
    with r1:
        rule_src = st.selectbox("Campo original da Base 1", [""] + cols_a, key="cm_rule_src")
    with r2:
        rule_tgt = st.selectbox("Campo correspondente da Base 2", [""] + cols_b, key="cm_rule_tgt")
    with r3:
        if st.button("Carregar associação"):
            if not rule_src:
                st.warning("Selecione o campo original da Base 1.")
            else:
                vals = sorted([v for v in _to_text(df_a[rule_src]).unique().tolist() if _clean_text(v)])
                st.session_state["cm_live_mapping"] = {v: "" for v in vals}

    rule_file = st.file_uploader("Importar regra (.json/.csv)", type=["json", "csv"], key="cm_rule_file")
    if rule_file is not None:
        try:
            st.session_state["cm_rule_preview"] = _parse_rule_upload(rule_file)
            st.success("Regra carregada para aplicação.")
        except Exception as e:
            st.error(str(e))

    if st.session_state["cm_rule_preview"] and st.button("Aplicar regra importada"):
        st.session_state["cm_rules"].extend(st.session_state["cm_rule_preview"])
        st.session_state["cm_rule_preview"] = []
        st.success("Regra importada aplicada.")

    live_map = st.session_state.get("cm_live_mapping", {})
    if live_map:
        st.markdown("**Montagem da regra atual**")
        for i, src_val in enumerate(list(live_map.keys())):
            cc1, cc2 = st.columns([1.4, 1.4])
            with cc1:
                st.text_input(f"Base 1 #{i+1}", value=src_val, disabled=True, key=f"cm_map_src_{i}")
            with cc2:
                st.session_state["cm_live_mapping"][src_val] = st.selectbox(
                    f"Base 2 #{i+1}", [""] + cols_b if False else [""] + sorted([v for v in _to_text(df_b[rule_tgt]).unique().tolist() if _clean_text(v)]) if rule_tgt else [""],
                    key=f"cm_map_tgt_{i}",
                )
        if st.button("Confirmar regra atual"):
            mapping = {k: v for k, v in st.session_state["cm_live_mapping"].items() if _clean_text(v)}
            if mapping:
                st.session_state["cm_rules"].append({"source_col": rule_src, "target_col": rule_tgt, "mapping": mapping})
                st.session_state["cm_live_mapping"] = {}
                st.success(f"Regra salva com {len(mapping)} associação(ões).")

    if st.session_state["cm_rules"]:
        st.markdown("**Regras confirmadas**")
        for r in st.session_state["cm_rules"]:
            st.write(f"- {r.get('source_col','')} -> {r.get('target_col','')} | {len(r.get('mapping', {}))} associação(ões)")
        jbytes, cbytes = _export_rule_files(st.session_state["cm_rules"])
        d1, d2 = st.columns(2)
        with d1:
            st.download_button("Baixar regra (.json)", jbytes, "regra_equivalencia.json", "application/json")
        with d2:
            st.download_button("Baixar regra (.csv)", cbytes, "regra_equivalencia.csv", "text/csv")

    df_a2, mapped_meta = _apply_rules_to_base1(df_a, st.session_state["cm_rules"])
    base1_options = cols_a + [m["mapped_col"] for m in mapped_meta if m["mapped_col"] not in cols_a]

    # 3. Chave da análise
    st.subheader("3) Campos que identificam o mesmo registro nas duas bases")
    st.caption("Esses campos formam a chave de busca. A conciliação encontra o registro por aqui.")
    for i, row in enumerate(st.session_state["cm_key_rows"]):
        k1, k2, k3 = st.columns([1.2, 1.2, 0.8])
        with k1:
            a_col = st.selectbox(f"Base 1 #{i+1}", [""] + base1_options, key=f"cm_key_a_{i}")
        with k2:
            b_suggest = row.get("b", "")
            if a_col.startswith("[MAP]"):
                hit = next((m for m in mapped_meta if m["mapped_col"] == a_col), None)
                if hit and not b_suggest:
                    b_suggest = hit.get("target_col", "")
            b_col = st.selectbox(f"Base 2 #{i+1}", [""] + cols_b, index=([""] + cols_b).index(b_suggest) if b_suggest in ([""] + cols_b) else 0, key=f"cm_key_b_{i}")
        with k3:
            label = st.text_input(f"Nome da dimensão #{i+1}", value=row.get("label") or _friendly_label(a_col, b_col), key=f"cm_key_lbl_{i}")
        st.session_state["cm_key_rows"][i] = {"a": a_col, "b": b_col, "label": label}

    if st.button("Adicionar par-chave"):
        st.session_state["cm_key_rows"].append({"a": "", "b": "", "label": ""})

    # 4. Campos de valor
    st.subheader("4) Quais campos deseja confrontar para validar valores")
    st.caption("Esses campos serão comparados e explicados no resumo executivo.")
    for i, row in enumerate(st.session_state["cm_val_rows"]):
        v1, v2, v3 = st.columns([1.2, 1.2, 0.8])
        with v1:
            a_col = st.selectbox(f"Valor Base 1 #{i+1}", [""] + cols_a, key=f"cm_val_a_{i}")
        with v2:
            b_col = st.selectbox(f"Valor Base 2 #{i+1}", [""] + cols_b, key=f"cm_val_b_{i}")
        with v3:
            label = st.text_input(f"Nome do valor #{i+1}", value=row.get("label") or _friendly_label(a_col, b_col), key=f"cm_val_lbl_{i}")
        st.session_state["cm_val_rows"][i] = {"a": a_col, "b": b_col, "label": label}

    if st.button("Adicionar campo de valor"):
        st.session_state["cm_val_rows"].append({"a": "", "b": "", "label": ""})

    # 5. Saída
    st.subheader("5) Como deseja receber o resultado?")
    gerar_exec = st.checkbox("Gerar resumo executivo", value=True)
    incluir_nao_encontrados = st.checkbox("Incluir não encontrados no detalhe", value=True)

    valid_keys = [r for r in st.session_state["cm_key_rows"] if r.get("a") and (r.get("b") or str(r.get("a", "")).startswith("[MAP]"))]
    valid_vals = [r for r in st.session_state["cm_val_rows"] if r.get("a") and r.get("b")]

    st.markdown("### Como deseja montar o resumo executivo?")
    default_group = [r.get("label") or _friendly_label(r.get("a", ""), r.get("b", "")) for r in valid_keys]
    default_total = [r.get("label") or _friendly_label(r.get("a", ""), r.get("b", "")) for r in valid_vals]

    g1, g2 = st.columns(2)
    with g1:
        group_labels = st.multiselect("Agrupar resumo por", options=default_group, default=default_group[:1] if default_group else []) if gerar_exec else []
    with g2:
        total_labels = st.multiselect("O que deseja totalizar/confrontar", options=default_total, default=default_total) if gerar_exec else []

    # 6. Executar
    st.subheader("6) Processar análise")
    if st.button("Executar análise"):
        if not valid_keys:
            st.error("Informe pelo menos um par-chave válido.")
            return
        if not valid_vals:
            st.error("Informe pelo menos um campo de valor válido.")
            return

        for r in valid_keys:
            if not r.get("label"):
                r["label"] = _friendly_label(r["a"], r["b"])
        for r in valid_vals:
            if not r.get("label"):
                r["label"] = _friendly_label(r["a"], r["b"])

        results = _run_reconciliation(df_a2, df_b, valid_keys, valid_vals, mapped_meta)
        exec_df, detail_df = _build_executive_and_detail(results, group_labels, total_labels or default_total)

        if not incluir_nao_encontrados:
            detail_df = detail_df[detail_df["MOTIVO"].ne("Chave só na Base 1") & detail_df["MOTIVO"].ne("Chave só na Base 2")].copy()

        rules_rows = []
        for r in st.session_state["cm_rules"]:
            for src, tgt in r.get("mapping", {}).items():
                rules_rows.append({
                    "Campo Base 1": r.get("source_col", ""),
                    "Valor original Base 1": src,
                    "Campo Base 2": r.get("target_col", ""),
                    "Valor correspondente Base 2": tgt,
                })
        rules_df = pd.DataFrame(rules_rows)

        excel = _export_excel(results["resumo_global"], exec_df, detail_df, rules_df)

        st.success("Análise concluída.")
        st.markdown("**Resumo Executivo**")
        st.dataframe(results["resumo_global"], use_container_width=True)
        if gerar_exec:
            st.dataframe(exec_df, use_container_width=True)
        st.markdown("**Detalhe das Diferenças**")
        st.dataframe(detail_df.head(200), use_container_width=True)
        st.download_button(
            "Baixar Excel da análise",
            data=excel,
            file_name="ConciliaMais_V23.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
