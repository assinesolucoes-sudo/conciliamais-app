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


def _top_reason(series: pd.Series) -> str:
    s = series.dropna().astype(str)
    if s.empty:
        return ""
    return s.value_counts().index[0]


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
# Motor de conciliação
# ============================================================

def _run_reconciliation(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    key_pairs: List[dict],
    value_pairs: List[dict],
    mapped_meta: List[dict],
    base1_name: str,
    base2_name: str,
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
        aval = _to_number(merged[a_col]) if a_col in merged.columns else pd.Series([0.0] * len(merged))
        bval = _to_number(merged[b_col]) if b_col in merged.columns else pd.Series([0.0] * len(merged))
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

    motivo_base = np.select(
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
    merged["MOTIVO"] = motivo_base

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
        resumo_global_rows.append(
            {
                "Campo confrontado": lbl,
                f"Total {base1_name}": total_a,
                f"Total {base2_name}": total_b,
                "Diferença total": round(total_a - total_b, 2),
            }
        )
    resumo_global = pd.DataFrame(resumo_global_rows)

    return {"full": merged, "resumo_global": resumo_global, "value_labels": value_labels, "key_labels": key_labels}


def _build_executive_and_detail(
    results: Dict[str, pd.DataFrame],
    group_labels: List[str],
    value_labels: List[str],
    base1_name: str,
    base2_name: str,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
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

    # Resumo executivo: apenas agrupador + valores + diferença + motivo predominante
    agg_map = {}
    for lbl in value_labels:
        agg_map[f"VALOR::{lbl}::A"] = "sum"
        agg_map[f"VALOR::{lbl}::B"] = "sum"
        agg_map[f"DIF::{lbl}"] = "sum"

    exec_df = df.groupby(group_cols, dropna=False).agg(agg_map).reset_index()
    motive = df.groupby(group_cols, dropna=False)["MOTIVO"].agg(_top_reason).reset_index().rename(columns={"MOTIVO": "Motivo predominante da diferença"})
    exec_df = exec_df.merge(motive, on=group_cols, how="left")

    rename_map = {c: c.replace("DIM::", "") for c in group_cols}
    for lbl in value_labels:
        rename_map[f"VALOR::{lbl}::A"] = f"{lbl} {base1_name}"
        rename_map[f"VALOR::{lbl}::B"] = f"{lbl} {base2_name}"
        rename_map[f"DIF::{lbl}"] = f"Diferença {lbl}"
    exec_df = exec_df.rename(columns=rename_map)

    # Ponte da conciliação: global e por agrupador
    ponte_rows = []
    for lbl in value_labels:
        ponte_rows.append({"Agrupador": "TOTAL GERAL", "Campo confrontado": lbl, "Componente": f"Chave só na {base1_name}", "Valor": round(df.loc[df['MOTIVO'].eq(f'Chave só na {base1_name}'), f'DIF::{lbl}'].sum(), 2)})
        ponte_rows.append({"Agrupador": "TOTAL GERAL", "Campo confrontado": lbl, "Componente": f"Chave só na {base2_name}", "Valor": round(df.loc[df['MOTIVO'].eq(f'Chave só na {base2_name}'), f'DIF::{lbl}'].sum(), 2)})
        ponte_rows.append({"Agrupador": "TOTAL GERAL", "Campo confrontado": lbl, "Componente": f"Valor divergente entre {base1_name} e {base2_name}", "Valor": round(df.loc[df['MOTIVO'].eq(f'Valor divergente entre {base1_name} e {base2_name}'), f'DIF::{lbl}'].sum(), 2)})
        ponte_rows.append({"Agrupador": "TOTAL GERAL", "Campo confrontado": lbl, "Componente": "Duplicidade", "Valor": round(df.loc[df['MOTIVO'].eq('Duplicidade'), f'DIF::{lbl}'].sum(), 2)})
        ponte_rows.append({"Agrupador": "TOTAL GERAL", "Campo confrontado": lbl, "Componente": "Diferença final", "Valor": round(df[f'DIF::{lbl}'].sum(), 2)})

        if group_cols:
            grp = df.groupby(group_cols + ["MOTIVO"], dropna=False)[f"DIF::{lbl}"].sum().reset_index()
            for _, row in grp.iterrows():
                agrupador = " | ".join([_clean_text(row[c]) for c in group_cols])
                ponte_rows.append({
                    "Agrupador": agrupador,
                    "Campo confrontado": lbl,
                    "Componente": row["MOTIVO"],
                    "Valor": round(row[f"DIF::{lbl}"], 2),
                })
    ponte_df = pd.DataFrame(ponte_rows)

    # Detalhe das diferenças
    detail = df[df["MOTIVO"].ne("Conciliado")].copy()
    detail_cols = []
    for lbl in results["key_labels"]:
        col = f"DIM::{lbl}"
        if col in detail.columns:
            detail_cols.append(col)
    extra_cols = [c for c in detail.columns if c.startswith(f"{base1_name} original::") or c.startswith(f"{base1_name} mapeado::") or c.startswith(f"{base2_name}::")]
    value_cols = []
    for lbl in value_labels:
        value_cols.extend([f"VALOR::{lbl}::A", f"VALOR::{lbl}::B", f"DIF::{lbl}"])
    detail = detail[detail_cols + extra_cols + value_cols + ["MOTIVO"]].copy()
    detail = detail.rename(columns=rename_map)
    for lbl in value_labels:
        detail = detail.rename(columns={
            f"VALOR::{lbl}::A": f"{lbl} {base1_name}",
            f"VALOR::{lbl}::B": f"{lbl} {base2_name}",
            f"DIF::{lbl}": f"Diferença {lbl}",
        })

    if value_labels:
        diff_col = f"Diferença {value_labels[0]}"
        if diff_col in exec_df.columns:
            exec_df = exec_df.assign(__ABS__=exec_df[diff_col].abs()).sort_values("__ABS__", ascending=False).drop(columns=["__ABS__"])
        if diff_col in detail.columns:
            detail = detail.assign(__ABS__=detail[diff_col].abs()).sort_values("__ABS__", ascending=False).drop(columns=["__ABS__"])

    return exec_df, detail, ponte_df


# ============================================================
# Excel
# ============================================================

def _set_column_formats(writer, sheet_name: str, df: pd.DataFrame):
    wb = writer.book
    ws = writer.sheets[sheet_name]

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
        if any(x in uc for x in ["DIFERENÇA", "TOTAL "]) and pd.api.types.is_numeric_dtype(ser):
            ws.set_column(i, i, width, fmt_diff if "DIFERENÇA" in uc else fmt_money)
        elif any(x in uc for x in ["VALOR", st.session_state['cm_base1_name'].upper(), st.session_state['cm_base2_name'].upper(), "TOTAL"]) and pd.api.types.is_numeric_dtype(ser):
            ws.set_column(i, i, width, fmt_money)
        elif uc.startswith("QTD") or uc.startswith("QTDE"):
            ws.set_column(i, i, width, fmt_int)
        else:
            ws.set_column(i, i, width, fmt_text)

    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(df), max(0, len(df.columns) - 1))


def _add_total_row(writer, sheet_name: str, df: pd.DataFrame, skip_when_only_total_geral: bool = False):
    if df.empty:
        return
    if skip_when_only_total_geral and "Agrupador" in df.columns:
        vals = {_clean_text(v) for v in df["Agrupador"].dropna().astype(str).unique().tolist() if _clean_text(v)}
        if vals and vals == {"TOTAL GERAL"}:
            return
    wb = writer.book
    ws = writer.sheets[sheet_name]
    total_row = len(df) + 1
    fmt_total_txt = wb.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1})
    fmt_total_money = wb.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1, "num_format": 'R$ #,##0.00'})
    fmt_total_int = wb.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1, "num_format": '0'})
    label = "TOTAL FILTRADO"
    if skip_when_only_total_geral and "Agrupador" in df.columns:
        label = "TOTAL FILTRADO (visão filtrada)"
    ws.write(total_row, 0, label, fmt_total_txt)
    for col_idx, col in enumerate(df.columns[1:], start=1):
        if pd.api.types.is_numeric_dtype(df[col]):
            col_letter = chr(65 + col_idx) if col_idx < 26 else None
            if col_letter:
                fmt = fmt_total_int if pd.api.types.is_integer_dtype(df[col]) else fmt_total_money
                ws.write_formula(total_row, col_idx, f"=SUBTOTAL(109,{col_letter}2:{col_letter}{len(df)+1})", fmt)


def _write_dataframe_block(ws, wb, start_row: int, start_col: int, title: str, df: pd.DataFrame, money_cols=None, int_cols=None):
    money_cols = set(list(money_cols) if money_cols is not None else [])
    int_cols = set(list(int_cols) if int_cols is not None else [])
    header_fmt = wb.add_format({"bold": True, "bg_color": "#1F2A44", "font_color": "#FFFFFF", "border": 1, "align": "left", "valign": "vcenter"})
    colhead_fmt = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "align": "center", "valign": "vcenter"})
    text_fmt = wb.add_format({"border": 1})
    money_fmt = wb.add_format({"border": 1, "num_format": 'R$ #,##0.00'})
    diff_fmt = wb.add_format({"border": 1, "num_format": 'R$ #,##0.00', "font_color": "#C00000", "bold": True})
    int_fmt = wb.add_format({"border": 1, "num_format": '0'})

    width = max(1, len(df.columns))
    ws.merge_range(start_row, start_col, start_row, start_col + width - 1, title, header_fmt)
    if df.empty:
        ws.write(start_row + 1, start_col, "Sem dados para exibir.", text_fmt)
        return start_row + 2

    for j, col in enumerate(df.columns):
        ws.write(start_row + 1, start_col + j, col, colhead_fmt)
        max_len = len(str(col))
        for i, val in enumerate(df[col].tolist(), start=0):
            cell_row = start_row + 2 + i
            v = val
            fmt = text_fmt
            if col in int_cols and pd.notna(v) and v != "":
                fmt = int_fmt
            elif col in money_cols and pd.notna(v) and v != "":
                fmt = diff_fmt if "diferen" in _norm_name(col) else money_fmt
            elif isinstance(v, (int, float, np.integer, np.floating)) and pd.notna(v):
                if "diferen" in _norm_name(col) or "valor" in _norm_name(col) or "total" in _norm_name(col):
                    fmt = diff_fmt if "diferen" in _norm_name(col) else money_fmt
            ws.write(cell_row, start_col + j, v, fmt)
            max_len = max(max_len, len(str(v)))
        ws.set_column(start_col + j, start_col + j, min(max_len + 2, 28))
    return start_row + 2 + len(df)


def _build_resumo_metricas(resumo_global: pd.DataFrame, detalhe: pd.DataFrame, ponte: pd.DataFrame) -> pd.DataFrame:
    diff_cols = [c for c in resumo_global.columns if "Diferença" in str(c)]
    total_diff = float(resumo_global[diff_cols[0]].sum()) if diff_cols else 0.0
    impacto_abs = float(resumo_global[diff_cols[0]].abs().sum()) if diff_cols else 0.0
    qtd_pend = int(len(detalhe))
    qtd_campos = int(resumo_global["Campo confrontado"].nunique()) if "Campo confrontado" in resumo_global.columns else 0
    comp_col = "Componente" if "Componente" in ponte.columns else None
    valor_col = "Valor" if "Valor" in ponte.columns else None
    def soma_comp(nome):
        if not comp_col or not valor_col:
            return 0.0
        return float(ponte.loc[(ponte["Agrupador"].eq("TOTAL GERAL")) & (ponte[comp_col].eq(nome)), valor_col].sum())
    return pd.DataFrame([
        {"Indicador": "Campos confrontados", "Valor": qtd_campos},
        {"Indicador": "Itens em divergência", "Valor": qtd_pend},
        {"Indicador": "Impacto líquido", "Valor": total_diff},
        {"Indicador": "Impacto absoluto", "Valor": impacto_abs},
        {"Indicador": "Chave só na base 1", "Valor": soma_comp("Chave só na " + st.session_state["cm_base1_name"])},
        {"Indicador": "Chave só na base 2", "Valor": soma_comp("Chave só na " + st.session_state["cm_base2_name"])},
        {"Indicador": "Divergência de valor", "Valor": soma_comp("Valor divergente entre " + st.session_state["cm_base1_name"] + " e " + st.session_state["cm_base2_name"])},
        {"Indicador": "Duplicidade", "Valor": soma_comp("Duplicidade")},
    ])


def _prepare_top_pendencias(detalhe: pd.DataFrame) -> pd.DataFrame:
    if detalhe.empty:
        return pd.DataFrame()
    diff_cols = [c for c in detalhe.columns if "Diferença" in str(c)]
    base_cols = [c for c in detalhe.columns if c not in diff_cols][:4]
    view_cols = base_cols + diff_cols[:1] + (["MOTIVO"] if "MOTIVO" in detalhe.columns else [])
    top = detalhe[view_cols].copy()
    if diff_cols:
        top = top.assign(__ABS__=top[diff_cols[0]].abs()).sort_values("__ABS__", ascending=False).drop(columns=["__ABS__"])
    return top.head(10)


def _export_excel(resumo_global: pd.DataFrame, resumo_exec: pd.DataFrame, detalhe: pd.DataFrame, ponte: pd.DataFrame, regras_df: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        wb = writer.book

        # Aba Resumo com layout executivo
        ws = wb.add_worksheet("Resumo")
        writer.sheets["Resumo"] = ws
        title_fmt = wb.add_format({"bold": True, "font_size": 18, "font_color": "#0F172A"})
        sub_fmt = wb.add_format({"italic": True, "font_color": "#5B6577"})
        ws.write(0, 0, "ConciliaMais — Resumo Executivo", title_fmt)
        ws.write(1, 0, pd.Timestamp.now().strftime("Gerado em %d/%m/%Y %H:%M:%S"), sub_fmt)

        metricas = _build_resumo_metricas(resumo_global, detalhe, ponte)
        top10 = _prepare_top_pendencias(detalhe)

        row_a = _write_dataframe_block(ws, wb, 3, 0, "Fechamento global dos campos confrontados", resumo_global, money_cols=resumo_global.select_dtypes(include=[np.number]).columns)
        row_b = _write_dataframe_block(ws, wb, 3, 6, "Indicadores executivos", metricas, money_cols=["Valor"])
        next_row = max(row_a, row_b) + 2
        next_row = _write_dataframe_block(ws, wb, next_row, 0, "Diferença por agrupador", resumo_exec, money_cols=resumo_exec.select_dtypes(include=[np.number]).columns)
        _write_dataframe_block(ws, wb, 3, 12, "Top 10 pendências mais impactantes", top10, money_cols=top10.select_dtypes(include=[np.number]).columns if not top10.empty else [])
        ws.freeze_panes(4, 0)

        # Aba Divergências
        detalhe.to_excel(writer, sheet_name="Divergencias", index=False)
        _set_column_formats(writer, "Divergencias", detalhe)
        _add_total_row(writer, "Divergencias", detalhe)

        # Aba Pontos de conciliação
        ponte.to_excel(writer, sheet_name="Pontos_Conciliacao", index=False)
        _set_column_formats(writer, "Pontos_Conciliacao", ponte)
        _add_total_row(writer, "Pontos_Conciliacao", ponte, skip_when_only_total_geral=True)

        if len(regras_df):
            regras_df.to_excel(writer, sheet_name="Regras_Aplicadas", index=False)
            _set_column_formats(writer, "Regras_Aplicadas", regras_df)

    return bio.getvalue()


# ============================================================
# App
# ============================================================

def main():
    _init_state()
    st.title("Concilia Mais - Match Inteligente V25")
    st.caption("Ferramenta de conciliação: fecha o total, mostra onde está a diferença e explica o motivo.")

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
        exec_df, detail_df, ponte_df = _build_executive_and_detail(results, group_labels, total_labels or default_total, base1_name, base2_name)

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

        excel = _export_excel(results["resumo_global"], exec_df, detail_df, ponte_df, rules_df)

        st.success("Análise concluída.")
        st.markdown("**Resumo da conciliação**")
        st.dataframe(results["resumo_global"], use_container_width=True)
        if gerar_exec:
            st.markdown("**Resumo executivo**")
            st.dataframe(exec_df, use_container_width=True)
            st.markdown("**Ponte da conciliação**")
            st.dataframe(ponte_df, use_container_width=True)
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
