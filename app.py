import json
import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Concilia Mais - Match Inteligente V25", layout="wide")

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
        ("depreciacao mensal", "Dep. Mensal"),
        ("depreciacao", "Dep. Acumulada"),
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


def _top_reason_diff(series: pd.Series) -> str:
    s = series.dropna().astype(str)
    s = s[~s.eq("Conciliado")]
    if s.empty:
        return "Sem diferença"
    return s.value_counts().index[0]


def _money_cols(df: pd.DataFrame) -> List[str]:
    cols = []
    for c in df.columns:
        uc = str(c).upper()
        if any(tag in uc for tag in ["TOTAL", "DIFERENÇA", "VALOR"]) or uc.startswith("SALDO ") or uc.startswith("AQUISIÇÃO ") or uc.startswith("DEP.") or uc.startswith("DEP "):
            if pd.api.types.is_numeric_dtype(df[c]):
                cols.append(c)
    return cols


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
        meta.append({"mapped_col": mapped_col, "source_col": sc, "target_col": tc})
    return out, meta


# ============================================================
# Motor
# ============================================================

def _run_reconciliation(df_a, df_b, key_pairs, value_pairs, mapped_meta, base1_name, base2_name):
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

    value_labels = []
    for vp in value_pairs:
        lbl = vp["label"]
        value_labels.append(lbl)
        a_col = f"{vp['a']}_A" if f"{vp['a']}_A" in merged.columns else vp["a"]
        b_col = f"{vp['b']}_B" if f"{vp['b']}_B" in merged.columns else vp["b"]
        aval = _to_number(merged[a_col]) if a_col in merged.columns else pd.Series([0.0] * len(merged), index=merged.index)
        bval = _to_number(merged[b_col]) if b_col in merged.columns else pd.Series([0.0] * len(merged), index=merged.index)
        merged[f"VALOR::{lbl}::A"] = aval.round(2)
        merged[f"VALOR::{lbl}::B"] = bval.round(2)
        merged[f"DIF::{lbl}"] = (aval - bval).round(2)

    merged["PRESENCA"] = merged["_merge"].map({"both": "Em ambas", "left_only": "Somente Base 1", "right_only": "Somente Base 2"})

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
            f"Chave só na {base1_name}",
            f"Chave só na {base2_name}",
            "Duplicidade",
            f"Valor divergente entre {base1_name} e {base2_name}",
        ],
        default="Conciliado",
    )

    for meta in mapped_meta:
        mapped_col = meta["mapped_col"]
        source_col = meta["source_col"]
        target_col = meta["target_col"]
        src_full = f"{source_col}_A" if f"{source_col}_A" in merged.columns else source_col
        map_full = f"{mapped_col}_A" if f"{mapped_col}_A" in merged.columns else mapped_col
        tgt_full = f"{target_col}_B" if f"{target_col}_B" in merged.columns else target_col
        if src_full in merged.columns:
            merged[f"{base1_name} original::{source_col}"] = _to_text(merged[src_full])
        if map_full in merged.columns:
            merged[f"{base1_name} mapeado::{target_col}"] = _to_text(merged[map_full])
        if tgt_full in merged.columns:
            merged[f"{base2_name}::{target_col}"] = _to_text(merged[tgt_full])

    resumo_global_rows = []
    for lbl in value_labels:
        total_a = round(merged[f"VALOR::{lbl}::A"].sum(), 2)
        total_b = round(merged[f"VALOR::{lbl}::B"].sum(), 2)
        resumo_global_rows.append({
            "Campo confrontado": lbl,
            f"Total {base1_name}": total_a,
            f"Total {base2_name}": total_b,
            "Diferença total": round(total_a - total_b, 2),
        })
    return {
        "full": merged,
        "resumo_global": pd.DataFrame(resumo_global_rows),
        "value_labels": value_labels,
        "key_labels": key_labels,
    }


def _build_outputs(results, group_labels, value_labels, base1_name, base2_name):
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

    agg_map = {}
    for lbl in value_labels:
        agg_map[f"VALOR::{lbl}::A"] = "sum"
        agg_map[f"VALOR::{lbl}::B"] = "sum"
        agg_map[f"DIF::{lbl}"] = "sum"

    resumo_exec = df.groupby(group_cols, dropna=False).agg(agg_map).reset_index()
    motive = (
        df[df["MOTIVO"].ne("Conciliado")]
        .groupby(group_cols, dropna=False)["MOTIVO"]
        .agg(_top_reason_diff)
        .reset_index()
        .rename(columns={"MOTIVO": "Motivo predominante da diferença"})
    )
    resumo_exec = resumo_exec.merge(motive, on=group_cols, how="left")
    resumo_exec["Motivo predominante da diferença"] = resumo_exec["Motivo predominante da diferença"].fillna("Sem diferença")

    rename_map = {c: c.replace("DIM::", "") for c in group_cols}
    for lbl in value_labels:
        rename_map[f"VALOR::{lbl}::A"] = f"{lbl} {base1_name}"
        rename_map[f"VALOR::{lbl}::B"] = f"{lbl} {base2_name}"
        rename_map[f"DIF::{lbl}"] = f"Diferença {lbl}"
    resumo_exec = resumo_exec.rename(columns=rename_map)

    ponte_total_rows = []
    ponte_agrup_rows = []
    for lbl in value_labels:
        motivos = [
            f"Chave só na {base1_name}",
            f"Chave só na {base2_name}",
            f"Valor divergente entre {base1_name} e {base2_name}",
            "Duplicidade",
        ]
        for motivo in motivos:
            ponte_total_rows.append({
                "Campo confrontado": lbl,
                "Componente": motivo,
                "Valor": round(df.loc[df["MOTIVO"].eq(motivo), f"DIF::{lbl}"].sum(), 2),
            })
        ponte_total_rows.append({
            "Campo confrontado": lbl,
            "Componente": "Diferença final",
            "Valor": round(df[f"DIF::{lbl}"].sum(), 2),
        })

        grp = df.groupby(group_cols + ["MOTIVO"], dropna=False)[f"DIF::{lbl}"].sum().reset_index()
        for _, row in grp.iterrows():
            if row["MOTIVO"] == "Conciliado" and abs(float(row[f"DIF::{lbl}"])) < 0.0001:
                continue
            agrupador = " | ".join([_clean_text(row[c]) for c in group_cols])
            ponte_agrup_rows.append({
                "Agrupador": agrupador,
                "Campo confrontado": lbl,
                "Componente": row["MOTIVO"],
                "Valor": round(row[f"DIF::{lbl}"], 2),
            })

    ponte_total = pd.DataFrame(ponte_total_rows)
    ponte_agrup = pd.DataFrame(ponte_agrup_rows)

    detalhe = df[df["MOTIVO"].ne("Conciliado")].copy()
    detail_cols = []
    for lbl in results["key_labels"]:
        col = f"DIM::{lbl}"
        if col in detalhe.columns:
            detail_cols.append(col)
    extra_cols = [c for c in detalhe.columns if c.startswith(f"{base1_name} original::") or c.startswith(f"{base1_name} mapeado::") or c.startswith(f"{base2_name}::")]
    value_cols = []
    for lbl in value_labels:
        value_cols.extend([f"VALOR::{lbl}::A", f"VALOR::{lbl}::B", f"DIF::{lbl}"])
    detalhe = detalhe[detail_cols + extra_cols + value_cols + ["MOTIVO"]].copy()
    detalhe = detalhe.rename(columns=rename_map)
    for lbl in value_labels:
        detalhe = detalhe.rename(columns={
            f"VALOR::{lbl}::A": f"{lbl} {base1_name}",
            f"VALOR::{lbl}::B": f"{lbl} {base2_name}",
            f"DIF::{lbl}": f"Diferença {lbl}",
        })

    if value_labels:
        diff_col = f"Diferença {value_labels[0]}"
        if diff_col in resumo_exec.columns:
            resumo_exec = resumo_exec.assign(__ABS__=resumo_exec[diff_col].abs()).sort_values("__ABS__", ascending=False).drop(columns=["__ABS__"])
        if diff_col in detalhe.columns:
            detalhe = detalhe.assign(__ABS__=detalhe[diff_col].abs()).sort_values("__ABS__", ascending=False).drop(columns=["__ABS__"])

    return resumo_exec, detalhe, ponte_total, ponte_agrup


# ============================================================
# Excel
# ============================================================

def _set_formats(ws, wb, df):
    fmt_head = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "align": "center", "valign": "vcenter"})
    fmt_text = wb.add_format({"border": 1})
    fmt_money = wb.add_format({"border": 1, "num_format": 'R$ #,##0.00'})
    fmt_diff = wb.add_format({"border": 1, "num_format": 'R$ #,##0.00', "font_color": "#C00000", "bold": True})
    fmt_int = wb.add_format({"border": 1, "num_format": '0'})
    for i, col in enumerate(df.columns):
        ws.write(0, i, col, fmt_head)
        ser = df[col]
        width = max(len(str(col)), min(60, ser.astype(str).map(len).max() if len(ser) else 10)) + 2
        uc = str(col).upper()
        if "DIFERENÇA" in uc and pd.api.types.is_numeric_dtype(ser):
            ws.set_column(i, i, width, fmt_diff)
        elif pd.api.types.is_numeric_dtype(ser):
            ws.set_column(i, i, width, fmt_money)
        elif uc.startswith("QTD") or uc.startswith("QTDE"):
            ws.set_column(i, i, width, fmt_int)
        else:
            ws.set_column(i, i, width, fmt_text)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(df), max(0, len(df.columns) - 1))


def _add_subtotal_row(ws, wb, df, start_row=1):
    if df.empty:
        return
    row = start_row + len(df)
    fmt_total_lbl = wb.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1})
    fmt_total_money = wb.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1, "num_format": 'R$ #,##0.00'})
    ws.write(row, 0, "TOTAL FILTRADO", fmt_total_lbl)
    for idx, col in enumerate(df.columns[1:], start=1):
        if pd.api.types.is_numeric_dtype(df[col]):
            col_letter = chr(65 + idx) if idx < 26 else None
            if col_letter:
                ws.write_formula(row, idx, f"=SUBTOTAL(109,{col_letter}{start_row+1}:{col_letter}{start_row+len(df)})", fmt_total_money)


def _export_excel(resumo_global, resumo_exec, detalhe, ponte_total, ponte_agrup, regras_df):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        wb = writer.book

        # RESUMO EXECUTIVO em dois blocos bem separados
        resumo_global.to_excel(writer, sheet_name="RESUMO_EXECUTIVO", index=False, startrow=1)
        resumo_exec.to_excel(writer, sheet_name="RESUMO_EXECUTIVO", index=False, startrow=len(resumo_global) + 5)
        ws_exec = writer.sheets["RESUMO_EXECUTIVO"]
        title_fmt = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1F1F1F"})
        section_fmt = wb.add_format({"bold": True, "font_size": 12, "font_color": "#1F1F1F", "bg_color": "#E2F0D9", "border": 1})
        ws_exec.write(0, 0, "Resumo Executivo", title_fmt)
        ws_exec.write(1, 0, "Fechamento global dos campos confrontados", section_fmt)
        ws_exec.write(len(resumo_global) + 4, 0, "Diferença por agrupador", section_fmt)
        _set_formats(ws_exec, wb, pd.concat([resumo_global, resumo_exec], ignore_index=True))

        detalhe.to_excel(writer, sheet_name="DETALHE_DIFERENCAS", index=False)
        ws_det = writer.sheets["DETALHE_DIFERENCAS"]
        _set_formats(ws_det, wb, detalhe)
        _add_subtotal_row(ws_det, wb, detalhe, start_row=1)

        # Ponte separada em dois blocos: total geral fora do filtro, agrupador filtrável
        ponte_total.to_excel(writer, sheet_name="PONTE_CONCILIACAO", index=False, startrow=1)
        ponte_agrup.to_excel(writer, sheet_name="PONTE_CONCILIACAO", index=False, startrow=len(ponte_total) + 5)
        ws_ponte = writer.sheets["PONTE_CONCILIACAO"]
        ws_ponte.write(0, 0, "Ponte da Conciliação - Total Geral", section_fmt)
        ws_ponte.write(len(ponte_total) + 4, 0, "Ponte da Conciliação por Agrupador", section_fmt)
        _set_formats(ws_ponte, wb, ponte_total)
        # aplica filtro apenas no bloco inferior
        if not ponte_agrup.empty:
            start = len(ponte_total) + 5
            for i, col in enumerate(ponte_agrup.columns):
                ws_ponte.write(start, i, col, wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "align": "center"}))
            ws_ponte.autofilter(start, 0, start + len(ponte_agrup), max(0, len(ponte_agrup.columns) - 1))
            ws_ponte.freeze_panes(start + 1, 0)
            _add_subtotal_row(ws_ponte, wb, ponte_agrup, start_row=start + 1)

        if len(regras_df):
            regras_df.to_excel(writer, sheet_name="REGRAS_APLICADAS", index=False)
            ws_reg = writer.sheets["REGRAS_APLICADAS"]
            _set_formats(ws_reg, wb, regras_df)

    return bio.getvalue()


# ============================================================
# App
# ============================================================

def main():
    _init_state()
    st.title("Concilia Mais - Match Inteligente V25")
    st.caption("Ferramenta de conciliação: fecha o total, mostra onde está a diferença, explica o motivo e entrega um Excel executivo.")

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

    base1_name = st.session_state["cm_base1_name"]
    base2_name = st.session_state["cm_base2_name"]
    df_a = _read_file(up_a)
    df_b = _read_file(up_b)
    cols_a = list(df_a.columns)
    cols_b = list(df_b.columns)

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
        valid_targets = [""] + sorted([v for v in _to_text(df_b[rule_tgt]).unique().tolist() if _clean_text(v)]) if rule_tgt else [""]
        for i, src_val in enumerate(list(live_map.keys())):
            cc1, cc2 = st.columns([1.4, 1.4])
            with cc1:
                st.text_input(f"Base 1 #{i+1}", value=src_val, disabled=True, key=f"cm_map_src_{i}")
            with cc2:
                st.session_state["cm_live_mapping"][src_val] = st.selectbox(f"Base 2 #{i+1}", valid_targets, key=f"cm_map_tgt_{i}")
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

    st.subheader("3) Campos que identificam o mesmo registro nas duas bases")
    st.caption("A chave encontra o registro correspondente.")
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

    st.subheader("4) Quais campos deseja confrontar para validar valores")
    st.caption("Os valores medem o impacto da conciliação.")
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

    st.subheader("5) Como deseja receber o resultado?")
    gerar_exec = st.checkbox("Gerar resumo executivo", value=True)
    incluir_so = st.checkbox("Incluir chaves só de uma base no detalhe", value=True)

    valid_keys = [r for r in st.session_state["cm_key_rows"] if r.get("a") and (r.get("b") or str(r.get("a", "")).startswith("[MAP]"))]
    valid_vals = [r for r in st.session_state["cm_val_rows"] if r.get("a") and r.get("b")]

    default_group = [r.get("label") or _friendly_label(r.get("a", ""), r.get("b", "")) for r in valid_keys]
    default_total = [r.get("label") or _friendly_label(r.get("a", ""), r.get("b", "")) for r in valid_vals]

    g1, g2 = st.columns(2)
    with g1:
        group_labels = st.multiselect("Agrupar resumo por", options=default_group, default=default_group[:1] if default_group else []) if gerar_exec else []
    with g2:
        total_labels = st.multiselect("O que deseja totalizar/confrontar", options=default_total, default=default_total) if gerar_exec else []

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

        results = _run_reconciliation(df_a2, df_b, valid_keys, valid_vals, mapped_meta, base1_name, base2_name)
        resumo_exec, detail_df, ponte_total, ponte_agrup = _build_outputs(results, group_labels, total_labels or default_total, base1_name, base2_name)

        if not incluir_so and "MOTIVO" in detail_df.columns:
            detail_df = detail_df[~detail_df["MOTIVO"].isin([f"Chave só na {base1_name}", f"Chave só na {base2_name}"])]

        rules_rows = []
        for r in st.session_state["cm_rules"]:
            for src, tgt in r.get("mapping", {}).items():
                rules_rows.append({
                    f"Campo {base1_name}": r.get("source_col", ""),
                    f"Valor original {base1_name}": src,
                    f"Campo {base2_name}": r.get("target_col", ""),
                    f"Valor correspondente {base2_name}": tgt,
                })
        rules_df = pd.DataFrame(rules_rows)

        excel = _export_excel(results["resumo_global"], resumo_exec, detail_df, ponte_total, ponte_agrup, rules_df)

        st.success("Análise concluída.")
        st.markdown("**Resumo da conciliação**")
        st.dataframe(results["resumo_global"], use_container_width=True)
        if gerar_exec:
            st.markdown("**Resumo executivo**")
            st.dataframe(resumo_exec, use_container_width=True)
            st.markdown("**Ponte da conciliação - total geral**")
            st.dataframe(ponte_total, use_container_width=True)
            st.markdown("**Ponte da conciliação - por agrupador**")
            st.dataframe(ponte_agrup, use_container_width=True)
        st.markdown("**Detalhe das diferenças**")
        st.dataframe(detail_df.head(200), use_container_width=True)
        st.download_button(
            "Baixar Excel da análise",
            data=excel,
            file_name="ConciliaMais_V25.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
