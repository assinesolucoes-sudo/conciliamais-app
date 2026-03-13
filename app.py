import re
import unicodedata
import time
from io import BytesIO
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Central de Conciliações - Análise de Consistência entre Bases V34", layout="wide")


# ============================================================
# Constantes
# ============================================================

SEMANTIC_TYPES = ["texto", "numero", "moeda", "percentual", "data"]


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
    s = re.sub(r"[^a-z0-9%]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _to_text(sr) -> pd.Series:
    if isinstance(sr, pd.DataFrame):
        if sr.shape[1] == 0:
            return pd.Series([""] * len(sr), index=sr.index)
        sr = sr.iloc[:, 0]

    sr = pd.Series(sr, index=sr.index, dtype="object")
    return sr.where(pd.notna(sr), "").astype(str).map(_clean_text)


def _to_key(sr) -> pd.Series:
    return _to_text(sr).str.upper()


def _to_number(sr) -> pd.Series:
    if isinstance(sr, pd.DataFrame):
        if sr.shape[1] == 0:
            return pd.Series([0.0] * len(sr), index=sr.index)
        sr = sr.iloc[:, 0]

    if pd.api.types.is_numeric_dtype(sr):
        return pd.to_numeric(sr, errors="coerce").fillna(0.0)

    s = pd.Series(sr, index=sr.index, dtype="object")
    s = s.where(pd.notna(s), "").astype(str).str.strip()
    s = s.str.replace(r"\s", "", regex=True)

    mask_pct = s.str.contains("%", na=False)
    mask_br = s.str.contains(",", na=False)

    s.loc[mask_br] = s.loc[mask_br].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    s = s.str.replace("%", "", regex=False)
    s = s.str.replace(r"[^0-9\-\.]", "", regex=True)

    out = pd.to_numeric(s, errors="coerce").fillna(0.0)
    out.loc[mask_pct] = out.loc[mask_pct] / 100.0
    return out


def _friendly_label(a: str, b: str) -> str:
    na = _norm_name(a)
    nb = _norm_name(b)
    if not na and not nb:
        return ""
    if na == nb:
        return _clean_text(a) or _clean_text(b)

    tokens_a = na.split()
    tokens_b = nb.split()

    if tokens_a and tokens_b:
        if set(tokens_a).issubset(set(tokens_b)):
            return _clean_text(a)
        if set(tokens_b).issubset(set(tokens_a)):
            return _clean_text(b)

    inter = [t for t in tokens_a if t in tokens_b]
    if inter:
        return " ".join(w.capitalize() for w in inter)

    return _clean_text(a) or _clean_text(b)


def _unique_preserve_order(values: List[str]) -> List[str]:
    seen = set()
    out = []
    for v in values:
        if v not in seen:
            out.append(v)
            seen.add(v)
    return out


def _parse_date_series(sr) -> pd.Series:
    if isinstance(sr, pd.DataFrame):
        if sr.shape[1] == 0:
            return pd.Series([pd.NaT] * len(sr), index=sr.index)
        sr = sr.iloc[:, 0]

    s = _to_text(sr)
    if s.empty:
        return pd.to_datetime(s, errors="coerce")

    parsed = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if parsed.notna().mean() >= 0.6:
        return parsed

    parsed = pd.to_datetime(s, errors="coerce", dayfirst=False)
    return parsed


def _sample_non_null(sr, n: int = 80) -> pd.Series:
    s = pd.Series(sr)
    s = s[pd.notna(s)]
    if s.empty:
        return s
    return s.head(n)


def _infer_semantic_type(col_name: str, sr: pd.Series, prefer_value: bool = False) -> str:
    col_norm = _norm_name(col_name)
    s = _sample_non_null(sr)

    if s.empty:
        return "numero" if prefer_value else "texto"

    name_has_date = any(k in col_norm for k in ["data", "dt", "emissao", "venc", "nasc", "date"])
    parsed_date = _parse_date_series(s)
    if name_has_date or parsed_date.notna().mean() >= 0.8:
        return "data"

    txt = _to_text(s)
    if txt.str.contains("%", regex=False).mean() >= 0.4:
        return "percentual"

    num = _to_number(s)
    numeric_ratio = pd.to_numeric(num, errors="coerce").notna().mean()
    if numeric_ratio >= 0.8:
        if any(k in col_norm for k in [
            "valor", "vlr", "preco", "preço", "saldo", "montante", "debito", "débito",
            "credito", "crédito", "total", "custo", "receita", "despesa", "liq", "líq",
            "bruto", "acumul", "aquisi", "deprec", "mensal"
        ]):
            return "moeda"
        if prefer_value:
            if (num.abs() >= 1000).mean() >= 0.3:
                return "moeda"
            return "numero"
        return "numero"

    if prefer_value:
        return "numero"
    return "texto"


def _infer_decimal_places_for_type(semantic_type: str, sr: pd.Series) -> int:
    if semantic_type == "percentual":
        return 2
    if semantic_type == "moeda":
        return 2
    if semantic_type in ["texto", "data"]:
        return 0

    sample_num = _to_number(_sample_non_null(sr))
    if sample_num.empty:
        return 0
    has_decimal = ((sample_num % 1).abs() > 0.0000001).any()
    return 2 if has_decimal else 0


def _excel_format_by_semantic_type(semantic_type: str, decimals: int = 2) -> str:
    if semantic_type == "moeda":
        return "R$ #,##0" + ("." + ("0" * decimals) if decimals > 0 else "")
    if semantic_type == "percentual":
        return "0" + ("." + ("0" * decimals) if decimals > 0 else "") + "%"
    if semantic_type == "numero":
        return "0" if decimals == 0 else "0." + ("0" * decimals)
    if semantic_type == "data":
        return "dd/mm/yyyy"
    return ""


def _get_col_series(df: pd.DataFrame, col_name: str):
    if col_name and col_name in df.columns:
        return df[col_name]
    return pd.Series([], dtype="object")


def _suggest_pair_semantics(df_a: pd.DataFrame, df_b: pd.DataFrame, col_a: str, col_b: str, prefer_value: bool = False):
    sr_a = _get_col_series(df_a, col_a)
    sr_b = _get_col_series(df_b, col_b)

    ta = _infer_semantic_type(col_a, sr_a, prefer_value=prefer_value) if col_a else "texto"
    tb = _infer_semantic_type(col_b, sr_b, prefer_value=prefer_value) if col_b else "texto"

    if ta == tb:
        final_type = ta
    elif "moeda" in [ta, tb]:
        final_type = "moeda"
    elif "percentual" in [ta, tb]:
        final_type = "percentual"
    elif "numero" in [ta, tb]:
        final_type = "numero"
    elif "data" in [ta, tb]:
        final_type = "data"
    else:
        final_type = "texto"

    dec_a = _infer_decimal_places_for_type(final_type, sr_a) if col_a else 0
    dec_b = _infer_decimal_places_for_type(final_type, sr_b) if col_b else 0
    decimals = max(dec_a, dec_b)

    if final_type == "moeda":
        decimals = max(2, decimals)
    if final_type == "percentual":
        decimals = max(2, decimals)

    return {
        "tipo_logico": final_type,
        "casas_decimais": decimals,
        "formato_excel": _excel_format_by_semantic_type(final_type, decimals),
    }


# ============================================================
# Preparação do matching
# ============================================================

def _prepare_base_for_matching(
    df: pd.DataFrame,
    key_pairs: List[dict],
    value_pairs: List[dict],
    side: str,
) -> pd.DataFrame:
    if side == "A":
        raw_key_cols = [kp["a"] for kp in key_pairs if kp.get("a") in df.columns]
        raw_val_cols = [vp["a"] for vp in value_pairs if vp.get("a") in df.columns]
    else:
        raw_key_cols = [kp["b"] for kp in key_pairs if kp.get("b") in df.columns]
        raw_val_cols = [vp["b"] for vp in value_pairs if vp.get("b") in df.columns]

    cols_needed = _unique_preserve_order(raw_key_cols + raw_val_cols)
    base = df[cols_needed].copy()

    for kp in key_pairs:
        src = kp["a"] if side == "A" else kp["b"]
        lbl = kp["label"]
        if src in base.columns:
            base[f"KEY::{lbl}"] = _to_key(base[src])
        else:
            base[f"KEY::{lbl}"] = ""

    for vp in value_pairs:
        src = vp["a"] if side == "A" else vp["b"]
        lbl = vp["label"]
        if src in base.columns:
            base[f"NUM::{lbl}"] = _to_number(base[src]).round(6)
        else:
            base[f"NUM::{lbl}"] = 0.0

    return base


# ============================================================
# Cache
# ============================================================

@st.cache_data(show_spinner=False)
def _read_file_cached(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    name = file_name.lower()
    if name.endswith(".csv"):
        for sep in [";", ",", None]:
            try:
                if sep is None:
                    return pd.read_csv(BytesIO(file_bytes), dtype=str, sep=None, engine="python")
                return pd.read_csv(BytesIO(file_bytes), dtype=str, sep=sep)
            except Exception:
                pass
        raise ValueError(f"Não foi possível ler o CSV: {file_name}")
    return pd.read_excel(BytesIO(file_bytes), dtype=str)


@st.cache_data(show_spinner=False)
def _get_column_list(columns: Tuple[str, ...]) -> List[str]:
    return list(columns)


# ============================================================
# Logs
# ============================================================

def _reset_logs():
    st.session_state["cm_logs"] = []


def _log_event(msg: str):
    if "cm_logs" not in st.session_state:
        st.session_state["cm_logs"] = []
    st.session_state["cm_logs"].append(f"{pd.Timestamp.now().strftime('%H:%M:%S')} - {msg}")


def _show_logs():
    logs = st.session_state.get("cm_logs", [])
    if logs:
        with st.expander("Log da execução", expanded=False):
            for item in logs:
                st.write(item)


# ============================================================
# Estado
# ============================================================

def _init_state():
    defaults = {
        "cm_analysis_name": "Nova análise",
        "cm_base1_name": "Base 1",
        "cm_base2_name": "Base 2",
        "cm_logs": [],
        "cm_key_rows": [{
            "a": "", "b": "", "label": "",
            "semantic_type": "",
            "excel_format": "",
            "type_manual": False,
            "fmt_manual": False,
            "last_signature": "",
        }],
        "cm_val_rows": [{
            "a": "", "b": "", "label": "",
            "semantic_type": "",
            "excel_format": "",
            "type_manual": False,
            "fmt_manual": False,
            "last_signature": "",
        }],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ============================================================
# Engine principal
# ============================================================

def _run_reconciliation(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    key_pairs: List[dict],
    value_pairs: List[dict],
    base1_name: str,
    base2_name: str,
) -> dict:
    labels_keys = [kp["label"] for kp in key_pairs]
    labels_vals = [vp["label"] for vp in value_pairs]

    A = _prepare_base_for_matching(df_a, key_pairs, value_pairs, side="A")
    B = _prepare_base_for_matching(df_b, key_pairs, value_pairs, side="B")

    key_cols = [f"KEY::{lbl}" for lbl in labels_keys]
    val_cols = [f"NUM::{lbl}" for lbl in labels_vals]

    if key_cols:
        A["__KEY__"] = A[key_cols].astype(str).agg("|".join, axis=1)
        B["__KEY__"] = B[key_cols].astype(str).agg("|".join, axis=1)
    else:
        A["__KEY__"] = ""
        B["__KEY__"] = ""

    dup_a = A.groupby("__KEY__", dropna=False).size().rename("QTD_DUP_A").reset_index()
    dup_b = B.groupby("__KEY__", dropna=False).size().rename("QTD_DUP_B").reset_index()

    agg_map = {c: "sum" for c in val_cols}
    for c in key_cols:
        agg_map[c] = "first"

    A1 = A.groupby("__KEY__", dropna=False, as_index=False).agg(agg_map)
    B1 = B.groupby("__KEY__", dropna=False, as_index=False).agg(agg_map)

    full = A1.merge(B1, on="__KEY__", how="outer", suffixes=(f"_{base1_name}", f"_{base2_name}"), indicator=True)
    full = full.merge(dup_a, on="__KEY__", how="left").merge(dup_b, on="__KEY__", how="left")
    full["QTD_DUP_A"] = full["QTD_DUP_A"].fillna(0).astype(int)
    full["QTD_DUP_B"] = full["QTD_DUP_B"].fillna(0).astype(int)

    for lbl in labels_keys:
        ca = f"KEY::{lbl}_{base1_name}"
        cb = f"KEY::{lbl}_{base2_name}"
        if ca not in full.columns and cb in full.columns:
            full[ca] = full[cb]
        if cb not in full.columns and ca in full.columns:
            full[cb] = full[ca]
        full[lbl] = full[ca].fillna(full[cb]).fillna("")

    for lbl in labels_vals:
        ca = f"NUM::{lbl}_{base1_name}"
        cb = f"NUM::{lbl}_{base2_name}"
        if ca not in full.columns:
            full[ca] = 0.0
        if cb not in full.columns:
            full[cb] = 0.0
        full[f"DIF::{lbl}"] = (pd.to_numeric(full[ca], errors="coerce").fillna(0.0) - pd.to_numeric(full[cb], errors="coerce").fillna(0.0)).round(6)

    def classificar_origem(x):
        if x == "left_only":
            return f"Somente {base1_name}"
        if x == "right_only":
            return f"Somente {base2_name}"
        return "Ambos"

    full["ORIGEM"] = full["_merge"].map(classificar_origem)

    diff_matrix = pd.DataFrame({lbl: full[f"DIF::{lbl}"].abs() > 0.0000001 for lbl in labels_vals}) if labels_vals else pd.DataFrame(index=full.index)
    full["TEM_DIF_VALOR"] = diff_matrix.any(axis=1) if not diff_matrix.empty else False
    full["EM_DUPLICIDADE"] = (full["QTD_DUP_A"] > 1) | (full["QTD_DUP_B"] > 1)

    full["STATUS"] = np.select(
        [
            full["EM_DUPLICIDADE"],
            full["_merge"].eq("left_only"),
            full["_merge"].eq("right_only"),
            full["TEM_DIF_VALOR"],
        ],
        [
            "Duplicidade",
            f"Chave só na {base1_name}",
            f"Chave só na {base2_name}",
            "Valor divergente",
        ],
        default="Conciliado",
    )

    rows = []
    for lbl in labels_vals:
        total_a = pd.to_numeric(full[f"NUM::{lbl}_{base1_name}"], errors="coerce").fillna(0.0).sum()
        total_b = pd.to_numeric(full[f"NUM::{lbl}_{base2_name}"], errors="coerce").fillna(0.0).sum()
        rows.append({
            "Campo confrontado": lbl,
            f"Total {base1_name}": round(total_a, 2),
            f"Total {base2_name}": round(total_b, 2),
            "Diferença total": round(total_a - total_b, 2),
        })
    resumo_global = pd.DataFrame(rows)

    return {
        "full": full,
        "resumo_global": resumo_global,
        "key_labels": labels_keys,
        "value_labels": labels_vals,
        "base1_name": base1_name,
        "base2_name": base2_name,
    }


# ============================================================
# Resumo executivo e detalhe
# ============================================================

def _build_executive_and_detail(results: dict, group_labels: List[str], total_labels: List[str], base1_name: str, base2_name: str):
    full = results["full"].copy()
    value_labels = results["value_labels"]

    if not group_labels:
        full["Agrupador"] = "TOTAL GERAL"
    else:
        full["Agrupador"] = full[group_labels].astype(str).agg(" | ".join, axis=1)
        full.loc[full["Agrupador"].str.strip().eq(""), "Agrupador"] = "TOTAL GERAL"

    comps = []
    for label in total_labels:
        if label not in value_labels:
            continue
        dif_col = f"DIF::{label}"
        tmp = full.groupby("Agrupador", dropna=False)[dif_col].sum().reset_index()
        tmp["Campo confrontado"] = label
        tmp["Componente"] = np.where(
            tmp[dif_col].abs() <= 0.0000001,
            "Sem diferença",
            np.where(tmp[dif_col] > 0, f"Valor divergente entre {base1_name} e {base2_name}", f"Valor divergente entre {base1_name} e {base2_name}"),
        )
        tmp = tmp.rename(columns={dif_col: "Valor"})
        comps.append(tmp[["Agrupador", "Campo confrontado", "Componente", "Valor"]])

    ponte = pd.concat(comps, ignore_index=True) if comps else pd.DataFrame(columns=["Agrupador", "Campo confrontado", "Componente", "Valor"])

    exclusivos_a = full[full["STATUS"].eq(f"Chave só na {base1_name}")]
    exclusivos_b = full[full["STATUS"].eq(f"Chave só na {base2_name}")]
    duplicados = full[full["STATUS"].eq("Duplicidade")]

    extras = []
    if not exclusivos_a.empty:
        x = exclusivos_a.groupby("Agrupador", dropna=False).size().reset_index(name="Valor")
        x["Campo confrontado"] = total_labels[0] if total_labels else ""
        x["Componente"] = f"Chave só na {base1_name}"
        extras.append(x[["Agrupador", "Campo confrontado", "Componente", "Valor"]])
    if not exclusivos_b.empty:
        x = exclusivos_b.groupby("Agrupador", dropna=False).size().reset_index(name="Valor")
        x["Campo confrontado"] = total_labels[0] if total_labels else ""
        x["Componente"] = f"Chave só na {base2_name}"
        extras.append(x[["Agrupador", "Campo confrontado", "Componente", "Valor"]])
    if not duplicados.empty:
        x = duplicados.groupby("Agrupador", dropna=False).size().reset_index(name="Valor")
        x["Campo confrontado"] = total_labels[0] if total_labels else ""
        x["Componente"] = "Duplicidade"
        extras.append(x[["Agrupador", "Campo confrontado", "Componente", "Valor"]])

    if extras:
        ponte = pd.concat([ponte] + extras, ignore_index=True)

    detalhe_cols = ["STATUS", "ORIGEM", "Agrupador"] + results["key_labels"]
    for lbl in value_labels:
        detalhe_cols.extend([
            f"NUM::{lbl}_{base1_name}",
            f"NUM::{lbl}_{base2_name}",
            f"DIF::{lbl}",
        ])
    detalhe = full[detalhe_cols].copy()

    rename_map = {
        "STATUS": "Status",
        "ORIGEM": "Origem",
        "Agrupador": "Agrupador",
    }
    for lbl in value_labels:
        rename_map[f"NUM::{lbl}_{base1_name}"] = f"{lbl} {base1_name}"
        rename_map[f"NUM::{lbl}_{base2_name}"] = f"{lbl} {base2_name}"
        rename_map[f"DIF::{lbl}"] = f"Diferença {lbl}"
    detalhe = detalhe.rename(columns=rename_map)

    resumo_exec = ponte.copy()
    return resumo_exec, detalhe, ponte


# ============================================================
# Metadados semânticos de saída
# ============================================================

def _build_output_semantic_maps(key_pairs: List[dict], value_pairs: List[dict], group_labels: List[str], base1_name: str, base2_name: str):
    sem_keys = {kp["label"]: {"tipo_logico": kp.get("semantic_type", "texto"), "casas_decimais": 0, "formato_excel": kp.get("excel_format", "")} for kp in key_pairs}
    sem_vals = {}
    for vp in value_pairs:
        fmt = vp.get("excel_format", "")
        tipo = vp.get("semantic_type", "numero")
        casas = 2
        if tipo == "texto":
            casas = 0
        elif fmt and "." in fmt:
            casas = max(fmt.count("0") - 1, 0)
        sem_vals[vp["label"]] = {"tipo_logico": tipo, "casas_decimais": casas, "formato_excel": fmt}

    detalhe_map = {
        "Status": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""},
        "Origem": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""},
        "Agrupador": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""},
    }
    for lbl in sem_keys:
        detalhe_map[lbl] = sem_keys[lbl]
    for lbl, meta in sem_vals.items():
        detalhe_map[f"{lbl} {base1_name}"] = meta
        detalhe_map[f"{lbl} {base2_name}"] = meta
        detalhe_map[f"Diferença {lbl}"] = meta

    resumo_global = {"Campo confrontado": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""}}
    for lbl, meta in sem_vals.items():
        resumo_global.setdefault(f"Total {base1_name}", meta)
        resumo_global.setdefault(f"Total {base2_name}", meta)
        resumo_global.setdefault("Diferença total", meta)

    resumo_exec = {
        "Agrupador": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""},
        "Campo confrontado": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""},
        "Componente": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""},
        "Valor": {"tipo_logico": "numero", "casas_decimais": 2, "formato_excel": "0.00"},
    }

    top10 = {"Chave": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""}, "Impacto": {"tipo_logico": "numero", "casas_decimais": 2, "formato_excel": "0.00"}}

    return {
        "detalhe": detalhe_map,
        "resumo_global": resumo_global,
        "resumo_exec": resumo_exec,
        "top10": top10,
    }


# ============================================================
# Formatação Excel
# ============================================================

def _make_xlsx_format(wb, semantic_type: str, decimals: int, is_header: bool = False, is_diff: bool = False, is_total: bool = False):
    base = {
        "border": 1,
        "valign": "vcenter",
    }
    if is_header:
        base.update({"bold": True, "font_color": "#000000", "bg_color": "#D9E6F2", "align": "center"})
    elif is_total:
        base.update({"bold": True, "bg_color": "#F7EFD1"})
    else:
        base.update({"align": "left" if semantic_type == "texto" else "right"})

    if is_diff:
        base.update({"font_color": "#C00000", "bold": True})

    num_fmt = None
    if semantic_type == "moeda":
        num_fmt = "R$ #,##0" + ("." + ("0" * decimals) if decimals > 0 else "")
    elif semantic_type == "percentual":
        num_fmt = "0" + ("." + ("0" * decimals) if decimals > 0 else "") + "%"
    elif semantic_type == "numero":
        num_fmt = "0" if decimals == 0 else "0." + ("0" * decimals)
    elif semantic_type == "data":
        num_fmt = "dd/mm/yyyy"
    if num_fmt:
        base["num_format"] = num_fmt
    return wb.add_format(base)


def _auto_width(df: pd.DataFrame, col_name: str, min_w: int = 12, max_w: int = 28) -> int:
    if col_name not in df.columns:
        return min_w
    max_len = max([len(str(col_name))] + [len(_clean_text(v)) for v in df[col_name].head(300).astype(str).tolist()])
    return max(min_w, min(max_len + 2, max_w))


def _build_resumo_metricas(resumo_global: pd.DataFrame, detalhe: pd.DataFrame, ponte: pd.DataFrame, full: pd.DataFrame, base1_name: str, base2_name: str) -> pd.DataFrame:
    qtd_dup = int((full["STATUS"] == "Duplicidade").sum()) if "STATUS" in full.columns else 0
    qtd_only_a = int((full["STATUS"] == f"Chave só na {base1_name}").sum()) if "STATUS" in full.columns else 0
    qtd_only_b = int((full["STATUS"] == f"Chave só na {base2_name}").sum()) if "STATUS" in full.columns else 0
    qtd_div = int((full["STATUS"] == "Valor divergente").sum()) if "STATUS" in full.columns else 0
    qtd_conc = int((full["STATUS"] == "Conciliado").sum()) if "STATUS" in full.columns else 0

    metricas = pd.DataFrame([
        {"Indicador": "Campos confrontados", "Valor": int(len(resumo_global))},
        {"Indicador": f"Registros {base1_name}", "Valor": int((full["ORIGEM"].eq(f"Somente {base1_name}") | full["ORIGEM"].eq("Ambos")).sum())},
        {"Indicador": f"Registros {base2_name}", "Valor": int((full["ORIGEM"].eq(f"Somente {base2_name}") | full["ORIGEM"].eq("Ambos")).sum())},
        {"Indicador": "Itens em divergência", "Valor": qtd_div},
        {"Indicador": "Qtd. em duplicidade", "Valor": qtd_dup},
        {"Indicador": f"Qtd. só na {base1_name}", "Valor": qtd_only_a},
        {"Indicador": f"Qtd. só na {base2_name}", "Valor": qtd_only_b},
        {"Indicador": "Qtd. conciliados", "Valor": qtd_conc},
    ])
    return metricas


def _build_metricas_semantic_map(metricas: pd.DataFrame) -> Dict[str, dict]:
    return {
        "Indicador": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""},
        "Valor": {"tipo_logico": "numero", "casas_decimais": 0, "formato_excel": "0"},
    }


def _prepare_top_pendencias(detalhe: pd.DataFrame) -> pd.DataFrame:
    diff_cols = [c for c in detalhe.columns if c.startswith("Diferença ")]
    key_cols = [c for c in detalhe.columns if c not in ["Status", "Origem", "Agrupador"] + diff_cols and not c.endswith(" Base 1") and not c.endswith(" Base 2")]
    if not diff_cols:
        return pd.DataFrame(columns=["Chave", "Impacto"])
    tmp = detalhe.copy()
    tmp["Impacto"] = tmp[diff_cols].apply(pd.to_numeric, errors="coerce").fillna(0).abs().sum(axis=1)
    tmp["Chave"] = tmp[key_cols].astype(str).agg(" | ".join, axis=1) if key_cols else tmp.index.astype(str)
    tmp = tmp.sort_values("Impacto", ascending=False).head(10)
    return tmp[["Chave", "Impacto"]]


def _write_dataframe_block(ws, wb, start_row: int, start_col: int, title: str, df: pd.DataFrame, semantic_map: Dict[str, dict]):
    title_fmt = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F3157", "font_size": 10})
    colhead_fmt = _make_xlsx_format(wb, "texto", 0, is_header=True)
    text_fmt = _make_xlsx_format(wb, "texto", 0)

    ws.write(start_row, start_col, title, title_fmt)
    for j, col in enumerate(df.columns):
        ws.write(start_row + 1, start_col + j, col, colhead_fmt)

    for i in range(len(df)):
        for j, col in enumerate(df.columns):
            val = df.iloc[i, j]
            meta = semantic_map.get(col, {"tipo_logico": "texto", "casas_decimais": 0})
            is_diff = "diferença" in _norm_name(col)
            fmt = _make_xlsx_format(wb, meta["tipo_logico"], int(meta.get("casas_decimais", 0)), is_diff=is_diff)
            if meta["tipo_logico"] in ["numero", "moeda", "percentual"]:
                ws.write_number(start_row + 2 + i, start_col + j, float(pd.to_numeric(val, errors="coerce") or 0), fmt)
            else:
                ws.write(start_row + 2 + i, start_col + j, _clean_text(val), fmt if meta["tipo_logico"] != "texto" else text_fmt)

    for j, col in enumerate(df.columns):
        ws.set_column(start_col + j, start_col + j, _auto_width(df, col))

    return start_row + 2 + len(df)


def _write_resumo_global_block(ws, wb, start_row: int, start_col: int, title: str, df: pd.DataFrame, semantic_map: Dict[str, dict]):
    return _write_dataframe_block(ws, wb, start_row, start_col, title, df, semantic_map)


def _write_metricas_block(ws, wb, start_row: int, start_col: int, title: str, df: pd.DataFrame, semantic_map: Dict[str, dict]):
    return _write_dataframe_block(ws, wb, start_row, start_col, title, df, semantic_map)


def _set_column_formats(writer, sheet_name: str, df: pd.DataFrame, semantic_map: Dict[str, dict]):
    wb = writer.book
    ws = writer.sheets[sheet_name]
    colhead_fmt = _make_xlsx_format(wb, "texto", 0, is_header=True)
    text_fmt = _make_xlsx_format(wb, "texto", 0)

    for j, col in enumerate(df.columns):
        ws.write(0, j, col, colhead_fmt)
        meta = semantic_map.get(col, {"tipo_logico": "texto", "casas_decimais": 0})
        fmt = _make_xlsx_format(wb, meta["tipo_logico"], int(meta.get("casas_decimais", 0)), is_diff=("diferença" in _norm_name(col)))
        if meta["tipo_logico"] == "texto":
            ws.set_column(j, j, _auto_width(df, col), text_fmt)
        else:
            ws.set_column(j, j, _auto_width(df, col), fmt)


def _add_total_row(writer, sheet_name: str, df: pd.DataFrame, semantic_map: Dict[str, dict], skip_when_only_total_geral: bool = False):
    wb = writer.book
    ws = writer.sheets[sheet_name]
    fmt_total_txt = _make_xlsx_format(wb, "texto", 0, is_total=True)

    if skip_when_only_total_geral and "Agrupador" in df.columns:
        vals = {_clean_text(v) for v in df["Agrupador"].dropna().astype(str).unique().tolist() if _clean_text(v)}
        if vals == {"TOTAL GERAL"}:
            return

    row = len(df) + 1
    ws.write(row, 0, "Total", fmt_total_txt)
    for j, col in enumerate(df.columns[1:], start=1):
        meta = semantic_map.get(col, {"tipo_logico": "texto", "casas_decimais": 0})
        if meta["tipo_logico"] in ["numero", "moeda", "percentual"]:
            fmt = _make_xlsx_format(wb, meta["tipo_logico"], int(meta.get("casas_decimais", 0)), is_total=True, is_diff=("diferença" in _norm_name(col)))
            ws.write_formula(row, j, f"=SUM({chr(65+j)}2:{chr(65+j)}{len(df)+1})", fmt)
        else:
            ws.write_blank(row, j, None, fmt_total_txt)


def _export_excel(
    results: dict,
    resumo_exec: pd.DataFrame,
    detalhe: pd.DataFrame,
    ponte: pd.DataFrame,
    output_semantics: Dict[str, Dict[str, dict]],
    base1_name: str,
    base2_name: str,
    analysis_name: str = "Nova análise",
) -> bytes:
    resumo_global = results["resumo_global"]
    metricas = _build_resumo_metricas(resumo_global, detalhe, ponte, results["full"], base1_name, base2_name)
    metricas_semantic = _build_metricas_semantic_map(metricas)
    top10 = _prepare_top_pendencias(detalhe)

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        wb = writer.book

        ws = wb.add_worksheet("Resumo")
        writer.sheets["Resumo"] = ws

        title_fmt = wb.add_format({"bold": True, "font_size": 18, "font_color": "#0F172A"})
        sub_fmt = wb.add_format({"italic": True, "font_color": "#5B6577"})

        ws.write(0, 0, "Central de Conciliações — Resumo Executivo", title_fmt)
        ws.write(1, 0, f"Análise: {analysis_name}", sub_fmt)
        ws.write(2, 0, pd.Timestamp.now().strftime("Gerado em %d/%m/%Y %H:%M:%S"), sub_fmt)

        row_a = _write_resumo_global_block(
            ws, wb, 4, 0, "Fechamento global dos campos confrontados", resumo_global,
            semantic_map=output_semantics.get("resumo_global", {}),
        )

        row_b = _write_metricas_block(
            ws, wb, 4, 6, "Indicadores executivos", metricas,
            semantic_map=metricas_semantic,
        )

        next_row = max(row_a, row_b) + 2

        row_exec = _write_dataframe_block(
            ws, wb, next_row, 0, "Diferença por agrupador", resumo_exec,
            semantic_map=output_semantics.get("resumo_exec", {}),
        )

        top_start = row_exec + 2
        _write_dataframe_block(
            ws, wb, top_start, 0, "Top 10 pendências mais impactantes", top10,
            semantic_map=output_semantics.get("top10", {}),
        )

        ws.freeze_panes(5, 0)

        detalhe.to_excel(writer, sheet_name="Divergencias", index=False)
        _set_column_formats(writer, "Divergencias", detalhe, semantic_map=output_semantics.get("detalhe", {}))
        _add_total_row(writer, "Divergencias", detalhe, semantic_map=output_semantics.get("detalhe", {}))

        ponte_export = ponte.copy()
        if "Agrupador" in ponte_export.columns:
            ponte_export["Agrupador"] = ponte_export["Agrupador"].astype(str)
            ponte_export.loc[ponte_export["Agrupador"].eq("TOTAL GERAL"), "Agrupador"] = " TOTAL GERAL"
        ponte_export.to_excel(writer, sheet_name="Agrupador", index=False)
        _set_column_formats(writer, "Agrupador", ponte_export, semantic_map=output_semantics.get("resumo_exec", {}))
        _add_total_row(writer, "Agrupador", ponte_export, semantic_map=output_semantics.get("resumo_exec", {}), skip_when_only_total_geral=True)

        resumo_global.to_excel(writer, sheet_name="Fechamento_Global", index=False)
        _set_column_formats(writer, "Fechamento_Global", resumo_global, semantic_map=output_semantics.get("resumo_global", {}))
        _add_total_row(writer, "Fechamento_Global", resumo_global, semantic_map=output_semantics.get("resumo_global", {}))

    bio.seek(0)
    return bio.getvalue()


# ============================================================
# UI helpers
# ============================================================

def _render_header_row(labels: List[str]):
    cols = st.columns([1.3, 1.3, 1.2, 0.9, 1.1])
    for c, lbl in zip(cols, labels):
        c.markdown(f"**{lbl}**")


def _render_pair_line(df_a: pd.DataFrame, df_b: pd.DataFrame, cols_a: List[str], cols_b: List[str], row: dict, idx: int, prefix: str, is_value: bool = False):
    cols = st.columns([1.3, 1.3, 1.2, 0.9, 1.1])

    with cols[0]:
        a_col = st.selectbox(
            f"{'Valor ' if is_value else ''}Base 1 #{idx+1}",
            [""] + cols_a,
            index=([""] + cols_a).index(row.get("a", "")) if row.get("a", "") in ([""] + cols_a) else 0,
            key=f"{prefix}_a_{idx}",
            label_visibility="collapsed",
        )

    with cols[1]:
        b_col = st.selectbox(
            f"{'Valor ' if is_value else ''}Base 2 #{idx+1}",
            [""] + cols_b,
            index=([""] + cols_b).index(row.get("b", "")) if row.get("b", "") in ([""] + cols_b) else 0,
            key=f"{prefix}_b_{idx}",
            label_visibility="collapsed",
        )

    inferred = _suggest_pair_semantics(df_a, df_b, a_col, b_col, prefer_value=is_value)
    default_label = row.get("label") or _friendly_label(a_col, b_col)

    signature = f"{a_col}||{b_col}"
    last_signature = row.get("last_signature", "")
    fmt_manual = bool(row.get("fmt_manual", False))

    suggested_by_type = {
        "texto": "",
        "data": "dd/mm/yyyy",
        "numero": "0.00",
        "moeda": "R$ #,##0.00",
        "percentual": "0.00%",
    }

    if signature != last_signature:
        current_type = inferred["tipo_logico"]
        current_format = suggested_by_type.get(current_type, "")
        fmt_manual = False
    else:
        current_type = row.get("semantic_type") or inferred["tipo_logico"]
        current_format = row.get("excel_format", "") or suggested_by_type.get(current_type, "")

    with cols[2]:
        label = st.text_input(
            f"Nome #{idx+1}",
            value=default_label,
            key=f"{prefix}_lbl_{idx}",
            label_visibility="collapsed",
        )

    with cols[3]:
        selected_type = st.selectbox(
            f"Tipo de dado #{idx+1}",
            options=SEMANTIC_TYPES,
            index=SEMANTIC_TYPES.index(current_type) if current_type in SEMANTIC_TYPES else 0,
            key=f"{prefix}_type_{idx}",
            label_visibility="collapsed",
        )

    if selected_type != current_type:
        current_type = selected_type
        if not fmt_manual:
            current_format = suggested_by_type.get(current_type, "")

    format_disabled = current_type == "texto"

    if current_type == "texto":
        current_format = ""
        fmt_manual = False

    fmt_key = f"{prefix}_fmt_{idx}_{current_type}"

    with cols[4]:
        excel_format = st.text_input(
            f"Formato no relatório #{idx+1}",
            value=current_format,
            key=fmt_key,
            label_visibility="collapsed",
            disabled=format_disabled,
        )

    if current_type == "texto":
        excel_format = ""
        fmt_manual = False
    else:
        fmt_manual = excel_format != suggested_by_type.get(current_type, "")

    return {
        "a": a_col,
        "b": b_col,
        "label": label,
        "semantic_type": current_type,
        "excel_format": excel_format,
        "fmt_manual": fmt_manual,
        "last_signature": signature,
    }


# ============================================================
# App
# ============================================================

def main():
    _init_state()

    st.title("Central de Conciliações")
    st.subheader("Análise de Consistência entre Bases")
    st.caption("Compare duas bases, identifique divergências de registros e valores e gere relatórios executivos em Excel.")
    _show_logs()

    st.subheader("1) Dados da análise")
    a1, c1, c2 = st.columns([1.2, 1, 1])

    with a1:
        st.session_state["cm_analysis_name"] = st.text_input("Nome da análise", st.session_state["cm_analysis_name"])

    with c1:
        st.session_state["cm_base1_name"] = st.text_input("Nome da Base 1", st.session_state["cm_base1_name"])
        up_a = st.file_uploader("Arquivo da Base 1", type=["xlsx", "xls", "csv"], key="cm_up_a")

    with c2:
        st.session_state["cm_base2_name"] = st.text_input("Nome da Base 2", st.session_state["cm_base2_name"])
        up_b = st.file_uploader("Arquivo da Base 2", type=["xlsx", "xls", "csv"], key="cm_up_b")

    if not up_a or not up_b:
        st.info("Carregue as duas bases para continuar.")
        return

    analysis_name = st.session_state["cm_analysis_name"]
    base1_name = st.session_state["cm_base1_name"]
    base2_name = st.session_state["cm_base2_name"]

    with st.spinner("Carregando bases..."):
        df_a = _read_file_cached(up_a.getvalue(), up_a.name)
        df_b = _read_file_cached(up_b.getvalue(), up_b.name)

    cols_a = _get_column_list(tuple(df_a.columns))
    cols_b = _get_column_list(tuple(df_b.columns))

    st.subheader("2) Chave da conciliação")
    st.caption("Selecione os campos que o sistema deve usar para identificar o mesmo registro nas duas bases. O tipo de dado é sugerido automaticamente e o formato no relatório pode ser ajustado quando necessário.")
    _render_header_row(["Campo da Base 1", "Campo da Base 2", "Nome exibido", "Tipo de dado", "Formato no relatório"])

    for i, row in enumerate(st.session_state["cm_key_rows"]):
        st.session_state["cm_key_rows"][i] = _render_pair_line(
            df_a, df_b, cols_a, cols_b, row, i, prefix="cm_key", is_value=False
        )

    if st.button("Adicionar campo à chave da conciliação"):
        st.session_state["cm_key_rows"].append({
            "a": "", "b": "", "label": "",
            "semantic_type": "",
            "excel_format": "",
            "fmt_manual": False,
            "last_signature": "",
        })
        st.rerun()

    st.subheader("3) Campos para comparação")
    st.caption("Selecione os campos cujos valores devem ser comparados entre as bases. O tipo de dado é sugerido automaticamente e o formato no relatório pode ser ajustado quando necessário.")
    _render_header_row(["Campo da Base 1", "Campo da Base 2", "Nome exibido", "Tipo de dado", "Formato no relatório"])

    for i, row in enumerate(st.session_state["cm_val_rows"]):
        st.session_state["cm_val_rows"][i] = _render_pair_line(
            df_a, df_b, cols_a, cols_b, row, i, prefix="cm_val", is_value=True
        )

    if st.button("Adicionar campo para comparação"):
        st.session_state["cm_val_rows"].append({
            "a": "", "b": "", "label": "",
            "semantic_type": "",
            "excel_format": "",
            "fmt_manual": False,
            "last_signature": "",
        })
        st.rerun()

    st.subheader("4) Configuração do resumo")
    gerar_exec = st.checkbox("Gerar resumo executivo", value=True)

    valid_keys = [r for r in st.session_state["cm_key_rows"] if r.get("a") and r.get("b")]
    valid_vals = [r for r in st.session_state["cm_val_rows"] if r.get("a") and r.get("b")]

    default_group = [r.get("label") or _friendly_label(r.get("a", ""), r.get("b", "")) for r in valid_keys]
    default_total = [r.get("label") or _friendly_label(r.get("a", ""), r.get("b", "")) for r in valid_vals]

    g1, g2 = st.columns(2)
    with g1:
        group_labels = st.multiselect(
            "Agrupar resumo por",
            options=default_group,
            default=default_group[:1] if default_group else []
        ) if gerar_exec else []
    with g2:
        total_labels = st.multiselect(
            "O que deseja totalizar/confrontar",
            options=default_total,
            default=default_total
        ) if gerar_exec else []

    st.subheader("5) Processar conciliação")
    executar = st.button("Processar conciliação")

    if executar:
        if not valid_keys:
            st.error("Informe pelo menos um campo válido na chave da conciliação.")
            return

        if not valid_vals:
            st.error("Informe pelo menos um campo válido para comparação.")
            return

        for r in valid_keys:
            if not r.get("label"):
                r["label"] = _friendly_label(r["a"], r["b"])

        for r in valid_vals:
            if not r.get("label"):
                r["label"] = _friendly_label(r["a"], r["b"])

        _reset_logs()
        progress = st.progress(0, text="Iniciando processamento...")
        with st.spinner("Processando conciliação..."):
            t0 = time.perf_counter()

            _log_event("Lendo configurações da análise")
            progress.progress(10, text="Preparando comparação...")

            results = _run_reconciliation(df_a, df_b, valid_keys, valid_vals, base1_name, base2_name)
            _log_event("Conciliação das bases concluída")
            progress.progress(40, text="Montando resumo executivo...")

            exec_df, detail_df, ponte_df = _build_executive_and_detail(
                results,
                group_labels,
                total_labels or default_total,
                base1_name,
                base2_name,
            )
            _log_event("Resumo executivo e detalhe das diferenças gerados")
            progress.progress(65, text="Preparando formatos de saída...")

            output_semantics = _build_output_semantic_maps(
                key_pairs=valid_keys,
                value_pairs=valid_vals,
                group_labels=group_labels if group_labels else [valid_keys[0]["label"]],
                base1_name=base1_name,
                base2_name=base2_name,
            )

            progress.progress(80, text="Gerando arquivo Excel...")
            excel = _export_excel(
                results,
                exec_df,
                detail_df,
                ponte_df,
                output_semantics,
                base1_name,
                base2_name,
                analysis_name=analysis_name,
            )

            elapsed = time.perf_counter() - t0
            _log_event(f"Processamento concluído em {elapsed:.2f}s")
            progress.progress(100, text="Conciliação concluída")

        st.success(f"Conciliação concluída em {elapsed:.2f}s.")
        st.caption(f"Análise: {analysis_name}")
        _show_logs()
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
            file_name=f"Central_Conciliacoes_{analysis_name.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
