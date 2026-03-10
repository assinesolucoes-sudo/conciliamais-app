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

st.set_page_config(page_title="Match Inteligente V19", layout="wide")

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
            df[new_col] = _force_text_series(df[source_col]).map(lambda x: mapping.get(_norm_text(x), mapping.get(str(x), "")))
            created.append({"label": new_col, "source_col": source_col, "target_col": target_col, "new_col": new_col})
    return df, created


def build_analysis(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    base1_name: str,
    base2_name: str,
    key_pairs: List[dict],
    compare_pairs: List[dict],
    rules: List[dict],
    executive_group_labels: List[str],
    executive_value_labels: List[str],
) -> dict:
    if not key_pairs:
        raise ValueError("Adicione pelo menos um par de campos identificadores.")
    if not compare_pairs:
        raise ValueError("Adicione pelo menos um par de campos de valor para confronto.")

    a, mapped_fields = apply_rules_to_base1(df_a, rules)
    b = df_b.copy()

    a["__row_a"] = np.arange(len(a))
    b["__row_b"] = np.arange(len(b))

    key_meta = []
    key_cols_a = []
    key_cols_b = []
    for i, pair in enumerate(key_pairs, start=1):
        col_a = pair["base1_col"]
        col_b = pair["base2_col"]
        label = pair.get("label") or _suggest_label(col_a, col_b)
        norm_a = f"__key_a_{i}"
        norm_b = f"__key_b_{i}"
        disp_a = f"__disp_a_{i}"
        disp_b = f"__disp_b_{i}"
        canon = f"DIM__{label}"

        if col_a not in a.columns:
            raise ValueError(f"Campo da Base 1 não encontrado: {col_a}")
        if col_b not in b.columns:
            raise ValueError(f"Campo da Base 2 não encontrado: {col_b}")

        a[norm_a] = _normalize_key_series(a[col_a])
        b[norm_b] = _normalize_key_series(b[col_b])
        a[disp_a] = _force_text_series(a[col_a])
        b[disp_b] = _force_text_series(b[col_b])

        key_cols_a.append(norm_a)
        key_cols_b.append(norm_b)
        key_meta.append({
            "label": label,
            "col_a": col_a,
            "col_b": col_b,
            "norm_a": norm_a,
            "norm_b": norm_b,
            "disp_a": disp_a,
            "disp_b": disp_b,
            "canon": canon,
        })

    a["__join_key"] = _build_join_key(a, key_cols_a)
    b["__join_key"] = _build_join_key(b, key_cols_b)
    a["__dup_a"] = a.groupby("__join_key")["__join_key"].transform("size")
    b["__dup_b"] = b.groupby("__join_key")["__join_key"].transform("size")
    a["__occ"] = a.groupby("__join_key").cumcount() + 1
    b["__occ"] = b.groupby("__join_key").cumcount() + 1

    merged = a.merge(
        b,
        on=["__join_key", "__occ"],
        how="outer",
        suffixes=(f"__{base1_name}", f"__{base2_name}"),
        indicator=True,
    )

    def _mcol(original_col: str, base_name: str) -> str:
        cand1 = f"{original_col}__{base_name}"
        if cand1 in merged.columns:
            return cand1
        if original_col in merged.columns:
            return original_col
        raise KeyError(original_col)

    merged["HAS_BASE1"] = merged["__row_a"].notna()
    merged["HAS_BASE2"] = merged["__row_b"].notna()
    merged["STATUS_BASE"] = np.select(
        [
            merged["HAS_BASE1"] & merged["HAS_BASE2"],
            merged["HAS_BASE1"] & ~merged["HAS_BASE2"],
            ~merged["HAS_BASE1"] & merged["HAS_BASE2"],
        ],
        ["CASADO", f"SÓ {base1_name}", f"SÓ {base2_name}"],
        default="INDEFINIDO",
    )
    merged["CHAVE_ANALISE"] = merged["__join_key"]
    merged["DUP_BASE1"] = merged.get("__dup_a", 0).fillna(0).astype(int) > 1
    merged["DUP_BASE2"] = merged.get("__dup_b", 0).fillna(0).astype(int) > 1
    merged["EH_DUPLICIDADE"] = merged["DUP_BASE1"] | merged["DUP_BASE2"]

    canonical_group_options = []
    for meta in key_meta:
        a_disp_col = _mcol(meta["disp_a"], base1_name)
        b_disp_col = _mcol(meta["disp_b"], base2_name)
        merged[meta["canon"]] = merged[a_disp_col].fillna("")
        mask_empty = merged[meta["canon"]].eq("")
        merged.loc[mask_empty, meta["canon"]] = merged.loc[mask_empty, b_disp_col].fillna("")
        canonical_group_options.append(meta["label"])

    for mf in mapped_fields:
        label = mf["label"]
        canon = f"DIM__{label}"
        a_col = _mcol(mf["new_col"], base1_name)
        merged[canon] = merged[a_col].fillna("")
        if mf.get("target_col"):
            try:
                b_col = _mcol(mf["target_col"], base2_name)
                mask_empty = merged[canon].eq("")
                merged.loc[mask_empty, canon] = merged.loc[mask_empty, b_col].fillna("")
            except KeyError:
                pass
        key_meta.append({
            "label": label,
            "canon": canon,
            "col_a": mf["new_col"],
            "col_b": mf.get("target_col", ""),
        })
        canonical_group_options.append(label)

    compare_meta = []
    divergence_cols = []
    for i, pair in enumerate(compare_pairs, start=1):
        col_a = pair["base1_col"]
        col_b = pair["base2_col"]
        label = pair.get("label") or _suggest_label(col_a, col_b)
        tol = float(pair.get("tolerance", 0.0) or 0.0)

        col_a_m = f"VALOR_{base1_name}_{label}"
        col_b_m = f"VALOR_{base2_name}_{label}"
        col_dif = f"DIF_{label}"

        merged[col_a_m] = _to_number(merged[_mcol(col_a, base1_name)])
        merged[col_b_m] = _to_number(merged[_mcol(col_b, base2_name)])
        merged[col_dif] = merged[col_a_m] - merged[col_b_m]
        div_flag = f"__DIV_{i}"
        merged[div_flag] = merged["HAS_BASE1"] & merged["HAS_BASE2"] & (merged[col_dif].abs() > tol)
        divergence_cols.append(div_flag)
        compare_meta.append({
            "label": label,
            "col_a": col_a,
            "col_b": col_b,
            "val_a": col_a_m,
            "val_b": col_b_m,
            "diff": col_dif,
            "tol": tol,
        })

    merged["EH_DIVERGENCIA"] = merged[divergence_cols].any(axis=1) if divergence_cols else False
    merged["EH_AUSENTE"] = merged["STATUS_BASE"].ne("CASADO")

    def _motivo_row(r):
        reasons = []
        if r["DUP_BASE1"]:
            reasons.append(f"Duplicidade {base1_name}")
        if r["DUP_BASE2"]:
            reasons.append(f"Duplicidade {base2_name}")
        if r["STATUS_BASE"] == f"SÓ {base1_name}":
            reasons.append(f"Somente {base1_name}")
        elif r["STATUS_BASE"] == f"SÓ {base2_name}":
            reasons.append(f"Somente {base2_name}")
        elif r["EH_DIVERGENCIA"]:
            labels = [meta["label"] for meta, dc in zip(compare_meta, divergence_cols) if r[dc]]
            reasons.append("Diferença em " + ", ".join(labels))
        if not reasons:
            reasons.append("Coerente")
        return " | ".join(reasons)

    merged["MOTIVO"] = merged.apply(_motivo_row, axis=1)

    if not executive_group_labels:
        executive_group_labels = [meta["label"] for meta in key_meta if meta["label"] in canonical_group_options]
        if not executive_group_labels and canonical_group_options:
            executive_group_labels = [canonical_group_options[0]]
    if not executive_value_labels:
        executive_value_labels = [meta["label"] for meta in compare_meta]

    label_to_canon = {meta["label"]: meta["canon"] for meta in key_meta if "canon" in meta}
    group_cols = [label_to_canon[l] for l in executive_group_labels if l in label_to_canon]
    if not group_cols:
        group_cols = [next(iter(label_to_canon.values()))]
        executive_group_labels = [next(iter(label_to_canon.keys()))]

    summary_agg = {
        "HAS_BASE1": "sum",
        "HAS_BASE2": "sum",
        "EH_DIVERGENCIA": "sum",
        "EH_AUSENTE": "sum",
        "EH_DUPLICIDADE": "sum",
        "MOTIVO": lambda s: " | ".join([f"{k} ({v})" for k, v in Counter(s).most_common(3)]),
    }
    rename_cols = {
        "HAS_BASE1": f"Qtde {base1_name}",
        "HAS_BASE2": f"Qtde {base2_name}",
        "EH_DIVERGENCIA": "Qtde Divergências",
        "EH_AUSENTE": "Qtde Ausentes",
        "EH_DUPLICIDADE": "Qtde Duplicidades",
        "MOTIVO": "Motivo",
    }
    for meta in compare_meta:
        if meta["label"] in executive_value_labels:
            summary_agg[meta["val_a"]] = "sum"
            summary_agg[meta["val_b"]] = "sum"
            summary_agg[meta["diff"]] = "sum"
            rename_cols[meta["val_a"]] = f"{meta['label']} {base1_name}"
            rename_cols[meta["val_b"]] = f"{meta['label']} {base2_name}"
            rename_cols[meta["diff"]] = f"Diferença {meta['label']}"

    resumo_exec = merged.groupby(group_cols, dropna=False).agg(summary_agg).reset_index()
    resumo_exec.rename(columns=rename_cols, inplace=True)
    label_map_exec = {label_to_canon[k]: k for k in executive_group_labels if k in label_to_canon}
    resumo_exec.rename(columns=label_map_exec, inplace=True)
    resumo_exec.insert(len(executive_group_labels), "Qtde Registros", merged.groupby(group_cols, dropna=False).size().values)

    bridge_rows = []
    kpis = []
    for meta in compare_meta:
        if meta["label"] not in executive_value_labels:
            continue
        total_a = float(merged[meta["val_a"]].sum())
        total_b = float(merged[meta["val_b"]].sum())
        diff_total = total_a - total_b
        matched_diff = float(merged.loc[merged["STATUS_BASE"].eq("CASADO"), meta["diff"]].sum())
        only_a = float(merged.loc[merged["STATUS_BASE"].eq(f"SÓ {base1_name}"), meta["val_a"]].sum())
        only_b = -float(merged.loc[merged["STATUS_BASE"].eq(f"SÓ {base2_name}"), meta["val_b"]].sum())
        bridge_rows.extend([
            {"Campo": meta["label"], "Componente": f"{base1_name} total", "Valor": total_a},
            {"Campo": meta["label"], "Componente": f"{base2_name} total", "Valor": total_b},
            {"Campo": meta["label"], "Componente": "Diferença em registros casados", "Valor": matched_diff},
            {"Campo": meta["label"], "Componente": f"Somente {base1_name}", "Valor": only_a},
            {"Campo": meta["label"], "Componente": f"Somente {base2_name}", "Valor": only_b},
            {"Campo": meta["label"], "Componente": "Diferença total", "Valor": diff_total},
        ])
        kpis.append({
            "Campo": meta["label"],
            f"Total {base1_name}": total_a,
            f"Total {base2_name}": total_b,
            "Diferença total": diff_total,
        })

    bridge_df = pd.DataFrame(bridge_rows)
    kpis_df = pd.DataFrame(kpis)

    resumo_geral = {
        "Tipo de análise": "Confronto completo de bases",
        "Modo": "Auditoria completa sem hierarquia entre bases",
        "Registros analisados": int(max(len(df_a), len(df_b))),
        "Correspondências": int((merged["STATUS_BASE"] == "CASADO").sum()),
        f"Somente {base1_name}": int((merged["STATUS_BASE"] == f"SÓ {base1_name}").sum()),
        f"Somente {base2_name}": int((merged["STATUS_BASE"] == f"SÓ {base2_name}").sum()),
        "Duplicidades": int(merged["EH_DUPLICIDADE"].sum()),
        "Divergências": int(merged["EH_DIVERGENCIA"].sum()),
        "Aderência (%)": round(float((~merged["EH_DIVERGENCIA"] & ~merged["EH_AUSENTE"]).sum()) / max(len(merged), 1) * 100, 2),
    }

    detailed = merged.copy()
    for meta in key_meta:
        if meta.get("col_a"):
            try:
                detailed[f"{base1_name}_{meta['label']}"] = detailed[_mcol(meta['col_a'], base1_name)]
            except KeyError:
                pass
        if meta.get("col_b"):
            try:
                detailed[f"{base2_name}_{meta['label']}"] = detailed[_mcol(meta['col_b'], base2_name)]
            except KeyError:
                pass

    final_cols = []
    for meta in key_meta:
        a_name = f"{base1_name}_{meta['label']}"
        b_name = f"{base2_name}_{meta['label']}"
        if a_name in detailed.columns:
            final_cols.append(a_name)
        if b_name in detailed.columns:
            final_cols.append(b_name)
    for meta in compare_meta:
        final_cols.extend([meta["val_a"], meta["val_b"], meta["diff"]])
    final_cols += ["STATUS_BASE", "EH_DIVERGENCIA", "EH_AUSENTE", "EH_DUPLICIDADE", "MOTIVO", "CHAVE_ANALISE"]
    final_cols = [c for c in final_cols if c in detailed.columns]
    df_result = detailed[final_cols].copy()

    divergencias = df_result[df_result["EH_DIVERGENCIA"] | df_result["EH_AUSENTE"] | df_result["EH_DUPLICIDADE"]].copy()
    nao_encontrados = df_result[df_result["STATUS_BASE"].ne("CASADO")].copy()
    duplicidades = df_result[df_result["EH_DUPLICIDADE"]].copy()

    return {
        "df_result": df_result,
        "resumo_exec": resumo_exec,
        "bridge_df": bridge_df,
        "kpis_df": kpis_df,
        "resumo_geral": resumo_geral,
        "divergencias": divergencias,
        "nao_encontrados": nao_encontrados,
        "duplicidades": duplicidades,
        "group_options": canonical_group_options,
        "value_options": [m["label"] for m in compare_meta],
        "mapped_fields": mapped_fields,
    }


# =========================================================
# Excel
# =========================================================

def _autofit_columns(ws, df: pd.DataFrame):
    for i, col in enumerate(df.columns):
        max_len = max(len(str(col)), *(len(str(v)) for v in df[col].head(500).fillna("")))
        ws.set_column(i, i, min(max_len + 2, 28))


def _write_table_sheet(writer, name: str, df: pd.DataFrame, money_like: List[str] = None, int_like: List[str] = None):
    money_like = set(money_like or [])
    int_like = set(int_like or [])
    df2 = df.copy()
    df2.to_excel(writer, sheet_name=name, index=False, startrow=0)
    wb = writer.book
    ws = writer.sheets[name]
    fmt_money = wb.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00'})
    fmt_int = wb.add_format({"num_format": '0'})
    fmt_hdr = wb.add_format({"bold": True, "bg_color": '#DCE6F1', "border": 1})
    fmt_text = wb.add_format({"border": 1})

    for c, col in enumerate(df2.columns):
        ws.write(0, c, col, fmt_hdr)
        if col in money_like:
            ws.set_column(c, c, 16, fmt_money)
        elif col in int_like:
            ws.set_column(c, c, 14, fmt_int)
        else:
            ws.set_column(c, c, min(max(len(str(col)) + 2, 14), 32), fmt_text)
    ws.autofilter(0, 0, max(len(df2), 1), max(len(df2.columns) - 1, 0))
    ws.freeze_panes(1, 0)

    if len(df2) > 0:
        total_row = len(df2) + 2
        ws.write(total_row, 0, "TOTAL FILTRADO", wb.add_format({"bold": True, "bg_color": '#FFF2CC'}))
        for c, col in enumerate(df2.columns):
            if col in money_like or col in int_like:
                col_letter = chr(65 + c) if c < 26 else None
                if col_letter:
                    formula = f'=SUBTOTAL(109,{col_letter}2:{col_letter}{len(df2)+1})'
                    ws.write_formula(total_row, c, formula, fmt_money if col in money_like else fmt_int)


def to_excel_package(result: dict, base1_name: str, base2_name: str) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_title = wb.add_format({"bold": True, "font_size": 14})
        fmt_sec = wb.add_format({"bold": True, "bg_color": '#DCE6F1', 'border': 1})
        fmt_money = wb.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00'})
        fmt_text = wb.add_format({"border": 1})
        fmt_int = wb.add_format({"num_format": '0'})

        ws = wb.add_worksheet("RESUMO_EXECUTIVO")
        writer.sheets["RESUMO_EXECUTIVO"] = ws
        ws.write(0, 0, "Resumo Executivo", fmt_title)

        resumo_df = pd.DataFrame({"Indicador": list(result["resumo_geral"].keys()), "Valor": list(result["resumo_geral"].values())})
        resumo_df.to_excel(writer, sheet_name="RESUMO_EXECUTIVO", index=False, startrow=2, startcol=0)
        for c, col in enumerate(resumo_df.columns):
            ws.write(2, c, col, fmt_sec)
            ws.set_column(c, c, 24 if c == 0 else 22, fmt_text)

        kpi_df = result["kpis_df"].copy()
        if not kpi_df.empty:
            start = 2
            col0 = 4
            ws.write(start - 1, col0, "Impacto financeiro total", fmt_title)
            kpi_df.to_excel(writer, sheet_name="RESUMO_EXECUTIVO", index=False, startrow=start, startcol=col0)
            for c, col in enumerate(kpi_df.columns):
                ws.write(start, col0 + c, col, fmt_sec)
                ws.set_column(col0 + c, col0 + c, 20, fmt_money if col != "Campo" else fmt_text)

        bridge_df = result["bridge_df"].copy()
        if not bridge_df.empty:
            start = 2 + max(len(resumo_df), len(kpi_df) + 2) + 4
            ws.write(start - 1, 0, "Ponte de conciliação", fmt_title)
            bridge_df.to_excel(writer, sheet_name="RESUMO_EXECUTIVO", index=False, startrow=start, startcol=0)
            for c, col in enumerate(bridge_df.columns):
                ws.write(start, c, col, fmt_sec)
                ws.set_column(c, c, 26 if col != "Valor" else 18, fmt_money if col == "Valor" else fmt_text)

        resumo_exec = result["resumo_exec"].copy()
        if not resumo_exec.empty:
            start = max(18, 2 + max(len(resumo_df), len(kpi_df) + 2) + len(bridge_df) + 8)
            ws.write(start - 1, 0, "Resumo executivo por chave", fmt_title)
            resumo_exec.to_excel(writer, sheet_name="RESUMO_EXECUTIVO", index=False, startrow=start, startcol=0)
            money_cols = [c for c in resumo_exec.columns if c.startswith(f"{base1_name}") or c.startswith(f"{base2_name}") or c.startswith("Diferença ")]
            int_cols = [c for c in resumo_exec.columns if c.startswith("Qtde")]
            for c, col in enumerate(resumo_exec.columns):
                ws.write(start, c, col, fmt_sec)
                fmt = fmt_money if col in money_cols else (fmt_int if col in int_cols else fmt_text)
                ws.set_column(c, c, min(max(len(str(col)) + 2, 14), 28), fmt)
            ws.autofilter(start, 0, start + len(resumo_exec), max(len(resumo_exec.columns)-1, 0))
            ws.freeze_panes(start + 1, 0)
            total_row = start + len(resumo_exec) + 2
            ws.write(total_row, 0, "TOTAL FILTRADO", wb.add_format({"bold": True, "bg_color": '#FFF2CC'}))
            for c, col in enumerate(resumo_exec.columns):
                if col in money_cols or col in int_cols:
                    col_letter = chr(65 + c) if c < 26 else None
                    if col_letter:
                        formula = f'=SUBTOTAL(109,{col_letter}{start+2}:{col_letter}{start+len(resumo_exec)+1})'
                        ws.write_formula(total_row, c, formula, fmt_money if col in money_cols else fmt_int)

        # abas detalhadas
        money_cols_result = [c for c in result["df_result"].columns if c.startswith(f"VALOR_{base1_name}") or c.startswith(f"VALOR_{base2_name}") or c.startswith("DIF_")]
        int_cols_result = []
        _write_table_sheet(writer, "RESULTADO_COMPLETO", result["df_result"], money_like=money_cols_result, int_like=int_cols_result)
        _write_table_sheet(writer, "DIVERGENCIAS", result["divergencias"], money_like=money_cols_result)
        _write_table_sheet(writer, "NAO_ENCONTRADOS", result["nao_encontrados"], money_like=money_cols_result)
        _write_table_sheet(writer, "DUPLICIDADES", result["duplicidades"], money_like=money_cols_result)

    output.seek(0)
    return output.getvalue()


# =========================================================
# Estado
# =========================================================
if "v19_rules" not in st.session_state:
    st.session_state.v19_rules = []
if "v19_key_pairs" not in st.session_state:
    st.session_state.v19_key_pairs = []
if "v19_compare_pairs" not in st.session_state:
    st.session_state.v19_compare_pairs = []
if "v19_analysis" not in st.session_state:
    st.session_state.v19_analysis = None
if "v19_rule_editor_df" not in st.session_state:
    st.session_state.v19_rule_editor_df = pd.DataFrame()
if "v19_builder_loaded" not in st.session_state:
    st.session_state.v19_builder_loaded = False


# =========================================================
# Interface
# =========================================================
st.title("Match Inteligente V19")
st.caption("Confronto completo entre duas bases, sem hierarquia entre Base 1 e Base 2, usando chave da análise, regras de equivalência e fechamento financeiro consistente.")

col_a, col_b = st.columns(2)
with col_a:
    base1_name = st.text_input("Nome exibido da Base 1", value="RM")
    file_a = st.file_uploader("Upload Base 1 (.xlsx ou .csv)", type=["xlsx", "csv"], key="v19_file_a")
with col_b:
    base2_name = st.text_input("Nome exibido da Base 2", value="Protheus")
    file_b = st.file_uploader("Upload Base 2 (.xlsx ou .csv)", type=["xlsx", "csv"], key="v19_file_b")

if not (file_a and file_b):
    st.info("Envie as duas bases para configurar a auditoria.")
    st.stop()

try:
    df_a = read_any_table(file_a)
    df_b = read_any_table(file_b)
except Exception as e:
    st.error(f"Erro ao ler os arquivos: {e}")
    st.stop()

st.success(f"Base 1: {len(df_a):,} linhas | Base 2: {len(df_b):,} linhas".replace(",", "."))

# =========================================================
# Regras de equivalência
# =========================================================
st.markdown("## 1) Regras opcionais de equivalência")
st.caption("Use quando precisar transformar um valor da Base 1 em um valor equivalente da Base 2 antes do match principal.")

imp_a, imp_b = st.columns([1.5, 1])
with imp_a:
    imported_rule = st.file_uploader("Importar regra (.json ou .csv)", type=["json", "csv"], key="v19_rule_upload")
with imp_b:
    if st.button("Aplicar regra importada", use_container_width=True):
        try:
            imported = _parse_imported_rule(imported_rule)
            default_source = st.session_state.get("v19_rule_source", df_a.columns[0] if len(df_a.columns) else "")
            default_target = st.session_state.get("v19_rule_target", df_b.columns[0] if len(df_b.columns) else "")
            for r in imported:
                if not r.get("source_col"):
                    r["source_col"] = default_source
                if not r.get("target_col"):
                    r["target_col"] = default_target
            st.session_state.v19_rules = imported
            st.success(f"Regra importada com {sum(len(r.get('mapping', {})) for r in imported)} associações.")
        except Exception as e:
            st.error(str(e))

src_col, tgt_col = st.columns(2)
with src_col:
    source_rule_col = st.selectbox("Campo da Base 1 para equivalência", options=df_a.columns.tolist(), key="v19_rule_source")
with tgt_col:
    target_rule_col = st.selectbox("Campo da Base 2 de destino", options=df_b.columns.tolist(), key="v19_rule_target")

btn1, btn2 = st.columns([1, 2])
with btn1:
    if st.button("Carregar associação", use_container_width=True):
        source_values = sorted([v for v in _force_text_series(df_a[source_rule_col]).drop_duplicates().tolist() if v != ""])
        target_values = sorted([v for v in _force_text_series(df_b[target_rule_col]).drop_duplicates().tolist() if v != ""])
        existing = {}
        for r in st.session_state.v19_rules:
            if r.get("source_col") == source_rule_col and r.get("target_col") == target_rule_col:
                existing = r.get("mapping", {})
                break
        st.session_state.v19_rule_editor_df = pd.DataFrame({
            "USAR": [True if v in existing else False for v in source_values],
            "VALOR_BASE1": source_values,
            "VALOR_BASE2": [existing.get(v, "") for v in source_values],
        })
        st.session_state.v19_rule_target_options = target_values
        st.session_state.v19_builder_loaded = True
with btn2:
    if st.button("Limpar regras", use_container_width=True):
        st.session_state.v19_rules = []
        st.session_state.v19_builder_loaded = False
        st.session_state.v19_rule_editor_df = pd.DataFrame()

if st.session_state.v19_builder_loaded and not st.session_state.v19_rule_editor_df.empty:
    edited = st.data_editor(
        st.session_state.v19_rule_editor_df,
        use_container_width=True,
        hide_index=True,
        key="v19_rule_editor",
        column_config={
            "USAR": st.column_config.CheckboxColumn("Usar"),
            "VALOR_BASE1": st.column_config.TextColumn(f"{base1_name}", disabled=True),
            "VALOR_BASE2": st.column_config.SelectboxColumn(f"{base2_name}", options=st.session_state.get("v19_rule_target_options", [])),
        },
    )
    if st.button("Confirmar regra atual", use_container_width=True):
        mapping = {}
        for _, r in edited.iterrows():
            if bool(r["USAR"]) and _norm_text(r["VALOR_BASE2"]):
                mapping[_norm_text(r["VALOR_BASE1"])] = _norm_text(r["VALOR_BASE2"])
        if not mapping:
            st.warning("Selecione ao menos uma associação com valor de destino.")
        else:
            # remove regra antiga do mesmo par
            st.session_state.v19_rules = [
                r for r in st.session_state.v19_rules
                if not (r.get("source_col") == source_rule_col and r.get("target_col") == target_rule_col)
            ]
            st.session_state.v19_rules.append({
                "source_col": source_rule_col,
                "target_col": target_rule_col,
                "mapping": mapping,
            })
            st.success(f"Regra salva com {len(mapping)} associações.")

if st.session_state.v19_rules:
    st.markdown("**Regras confirmadas**")
    for i, rule in enumerate(st.session_state.v19_rules, start=1):
        st.write(f"{i}. {rule.get('source_col') or '[importado]'} → {rule.get('target_col') or '[destino]'} | {len(rule.get('mapping', {}))} associações")
    json_bytes, csv_bytes = _export_rules_payload(st.session_state.v19_rules)
    d1, d2 = st.columns(2)
    with d1:
        st.download_button("Baixar regra (.json)", data=json_bytes, file_name="regra_equivalencia.json", mime="application/json", use_container_width=True)
    with d2:
        st.download_button("Baixar regra (.csv)", data=csv_bytes, file_name="regra_equivalencia.csv", mime="text/csv", use_container_width=True)

# =========================================================
# Campos identificadores
# =========================================================
st.markdown("## 2) Campos que identificam o mesmo registro nas duas bases")
st.caption("Monte a chave da análise. Ela pode ser simples ou composta. O sistema vai confrontar usando a combinação exata escolhida.")

colk1, colk2, colk3 = st.columns([1, 1, 0.6])
# opções com mapped fields
mapped_preview_df, mapped_preview = apply_rules_to_base1(df_a, st.session_state.v19_rules)
base1_key_options = mapped_preview_df.columns.tolist()
with colk1:
    key_a_sel = st.selectbox(f"Campo da {base1_name}", options=base1_key_options, key="v19_key_a")
with colk2:
    key_b_sel = st.selectbox(f"Campo da {base2_name}", options=df_b.columns.tolist(), key="v19_key_b")
with colk3:
    key_label = st.text_input("Nome da dimensão", value=_suggest_label(key_a_sel, key_b_sel), key="v19_key_label")

if st.button("Adicionar par-chave"):
    pair = {"base1_col": key_a_sel, "base2_col": key_b_sel, "label": _norm_text(key_label) or _suggest_label(key_a_sel, key_b_sel)}
    if pair not in st.session_state.v19_key_pairs:
        st.session_state.v19_key_pairs.append(pair)

if st.session_state.v19_key_pairs:
    for i, pair in enumerate(st.session_state.v19_key_pairs):
        c1, c2 = st.columns([6, 1])
        with c1:
            st.write(f"**{pair['label']}**: {pair['base1_col']} ↔ {pair['base2_col']}")
        with c2:
            if st.button("Remover", key=f"rm_key_{i}"):
                st.session_state.v19_key_pairs.pop(i)
                st.rerun()

# =========================================================
# Campos de valor
# =========================================================
st.markdown("## 3) Campos de valor para confronto")
st.caption("Aqui entram os valores que o sistema vai comparar e totalizar no resumo executivo.")

colv1, colv2, colv3, colv4 = st.columns([1, 1, 0.8, 0.6])
with colv1:
    cmp_a_sel = st.selectbox(f"Valor da {base1_name}", options=df_a.columns.tolist(), key="v19_cmp_a")
with colv2:
    cmp_b_sel = st.selectbox(f"Valor da {base2_name}", options=df_b.columns.tolist(), key="v19_cmp_b")
with colv3:
    cmp_label = st.text_input("Nome do confronto", value=_suggest_label(cmp_a_sel, cmp_b_sel), key="v19_cmp_label")
with colv4:
    cmp_tol = st.number_input("Tolerância", min_value=0.0, value=0.0, step=0.01, key="v19_cmp_tol")

if st.button("Adicionar confronto de valor"):
    pair = {"base1_col": cmp_a_sel, "base2_col": cmp_b_sel, "label": _norm_text(cmp_label) or _suggest_label(cmp_a_sel, cmp_b_sel), "tolerance": cmp_tol}
    if pair not in st.session_state.v19_compare_pairs:
        st.session_state.v19_compare_pairs.append(pair)

if st.session_state.v19_compare_pairs:
    for i, pair in enumerate(st.session_state.v19_compare_pairs):
        c1, c2 = st.columns([6, 1])
        with c1:
            st.write(f"**{pair['label']}**: {pair['base1_col']} ↔ {pair['base2_col']} | tolerância {pair.get('tolerance', 0)}")
        with c2:
            if st.button("Remover", key=f"rm_cmp_{i}"):
                st.session_state.v19_compare_pairs.pop(i)
                st.rerun()

# Diagnóstico da chave
if st.session_state.v19_key_pairs:
    tmp_a, _ = apply_rules_to_base1(df_a, st.session_state.v19_rules)
    key_cols_tmp_a = [p['base1_col'] for p in st.session_state.v19_key_pairs]
    key_cols_tmp_b = [p['base2_col'] for p in st.session_state.v19_key_pairs]
    key_a = _build_join_key(tmp_a, key_cols_tmp_a)
    key_b = _build_join_key(df_b, key_cols_tmp_b)
    d1, d2 = st.columns(2)
    with d1:
        st.info(f"{base1_name}: {key_a.nunique():,} chaves distintas | {(key_a.duplicated(keep=False)).sum():,} linhas com repetição de chave".replace(',', '.'))
    with d2:
        st.info(f"{base2_name}: {key_b.nunique():,} chaves distintas | {(key_b.duplicated(keep=False)).sum():,} linhas com repetição de chave".replace(',', '.'))

# =========================================================
# Saída
# =========================================================
st.markdown("## 4) Como deseja receber o resultado?")
only_div = st.checkbox("Mostrar apenas divergências no resultado")
include_not_found = st.checkbox("Incluir não encontrados", value=True)
make_summary = st.checkbox("Gerar resumo executivo", value=True)

# opções de resumo - robustas
mapped_labels = [f"[MAP] {r['source_col']} -> {r['target_col']}" for r in st.session_state.v19_rules if r.get('source_col')]
default_group = [p['label'] for p in st.session_state.v19_key_pairs]
default_values = [p['label'] for p in st.session_state.v19_compare_pairs]

if make_summary:
    st.markdown("### Como deseja montar o resumo executivo?")
    executive_group = st.multiselect(
        "Agrupar resumo por",
        options=default_group,
        default=default_group,
        key="v19_exec_group",
    )
    executive_values = st.multiselect(
        "O que deseja totalizar/confrontar",
        options=default_values,
        default=default_values,
        key="v19_exec_vals",
    )
else:
    executive_group = []
    executive_values = []

if st.button("Executar análise", type="primary", use_container_width=True):
    try:
        result = build_analysis(
            df_a=df_a,
            df_b=df_b,
            base1_name=base1_name,
            base2_name=base2_name,
            key_pairs=st.session_state.v19_key_pairs,
            compare_pairs=st.session_state.v19_compare_pairs,
            rules=st.session_state.v19_rules,
            executive_group_labels=executive_group,
            executive_value_labels=executive_values,
        )
        st.session_state.v19_analysis = result
    except Exception as e:
        st.exception(e)

result = st.session_state.v19_analysis
if result:
    st.markdown("## 5) Resultado da análise")
    kpis_df = result["kpis_df"].copy()
    if not kpis_df.empty:
        st.markdown("### Impacto financeiro total")
        st.dataframe(kpis_df, use_container_width=True)

    if make_summary:
        st.markdown("### Resumo executivo por chave")
        st.dataframe(result["resumo_exec"], use_container_width=True, hide_index=True)
        st.markdown("### Ponte de conciliação")
        st.dataframe(result["bridge_df"], use_container_width=True, hide_index=True)

    view_df = result["df_result"]
    if only_div:
        view_df = result["divergencias"]
    elif include_not_found:
        view_df = result["df_result"]
    else:
        view_df = result["df_result"][result["df_result"]["STATUS_BASE"].eq("CASADO")]

    st.markdown("### Resultado completo")
    st.dataframe(view_df, use_container_width=True, hide_index=True)

    excel_bytes = to_excel_package(result, base1_name, base2_name)
    st.download_button(
        "Baixar análise (Excel)",
        data=excel_bytes,
        file_name=f"Match_Inteligente_V19_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
