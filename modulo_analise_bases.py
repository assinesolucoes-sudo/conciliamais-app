import re
import unicodedata
import time
from io import BytesIO
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import streamlit as st

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
        ("centro custo", "Centro de Custo"),
        ("ccusto", "Centro de Custo"),
        ("fornecedor", "Fornecedor"),
        ("cliente", "Cliente"),
        ("produto", "Produto"),
        ("documento", "Documento"),
        ("data", "Data"),
        ("percentual", "Percentual"),
        ("aliquota", "Alíquota"),
        ("quantidade", "Quantidade"),
    ]
    for key, label in mapping:
        if key in joined:
            return label
    if _clean_text(a) == _clean_text(b) and _clean_text(a):
        return _clean_text(a)
    return f"{_clean_text(a)} ↔ {_clean_text(b)}"


def _top_reason_from_df(df: pd.DataFrame) -> str:
    if df.empty or "MOTIVO" not in df.columns:
        return ""
    s = df["MOTIVO"].dropna().astype(str).map(_clean_text)
    s = s[s.ne("") & s.ne("Conciliado")]
    if s.empty:
        return ""
    return s.value_counts().index[0]


def _unique_preserve_order(items: List[str]) -> List[str]:
    seen = set()
    out = []
    for x in items:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def _build_hash_key_from_cols(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    if not cols:
        return pd.Series([0] * len(df), index=df.index, dtype="uint64")
    return pd.util.hash_pandas_object(df[cols], index=False).astype("uint64")


def _safe_get_series(df: pd.DataFrame, col_name: str, default_text: bool = True) -> pd.Series:
    if col_name not in df.columns:
        if default_text:
            return pd.Series([""] * len(df), index=df.index)
        return pd.Series([0.0] * len(df), index=df.index)

    obj = df[col_name]

    if isinstance(obj, pd.DataFrame):
        if obj.shape[1] == 0:
            if default_text:
                return pd.Series([""] * len(df), index=df.index)
            return pd.Series([0.0] * len(df), index=df.index)
        return obj.iloc[:, 0]

    return obj


# ============================================================
# Semântica
# ============================================================

def _sample_non_null(sr: pd.Series, n: int = 30) -> pd.Series:
    sr = _to_text(sr)
    sr = sr[sr.ne("")]
    return sr.head(n)


def _looks_like_date(sr: pd.Series) -> bool:
    sample = _sample_non_null(sr)
    if sample.empty:
        return False
    parsed = pd.to_datetime(sample, errors="coerce", dayfirst=True)
    return parsed.notna().mean() >= 0.7


def _looks_like_number(sr: pd.Series) -> bool:
    sample = _sample_non_null(sr)
    if sample.empty:
        return False
    raw = sample.astype(str)
    normalized = raw.str.replace(r"\s", "", regex=True)
    normalized = normalized.str.replace("%", "", regex=False)
    ok = normalized.str.match(r"^-?[\d\.\,]+$", na=False)
    return ok.mean() >= 0.7


def _infer_semantic_type(col_name: str, sr: pd.Series, prefer_value: bool = False) -> str:
    name = _norm_name(col_name)

    if any(k in name for k in ["percentual", "aliquota", "aliq", "perc", "taxa", "%"]):
        return "percentual"

    if any(k in name for k in ["valor", "saldo", "total", "debito", "credito", "montante", "preco", "vlr", "vl "]):
        return "moeda"

    if any(k in name for k in ["qtd", "quantidade", "qtde", "itens", "numero de", "num itens"]):
        return "numero"

    if any(k in name for k in ["data", "dt ", "dt_", "emissao", "vencimento", "baixa", "aquisicao"]):
        return "data"

    if _looks_like_date(sr):
        return "data"

    if _looks_like_number(sr):
        sample = _sample_non_null(sr)
        if not sample.empty and sample.str.contains("%", regex=False).mean() >= 0.3:
            return "percentual"
        if prefer_value:
            return "moeda"
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
# Estado
# ============================================================

def _init_state():
    defaults = {
        "cm_base1_name": "Base 1",
        "cm_base2_name": "Base 2",
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
# Motor de conciliação
# ============================================================

def _run_reconciliation(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    key_pairs: List[dict],
    value_pairs: List[dict],
    base1_name: str,
    base2_name: str,
) -> Dict[str, pd.DataFrame]:
    a = _prepare_base_for_matching(df_a, key_pairs, value_pairs, side="A")
    b = _prepare_base_for_matching(df_b, key_pairs, value_pairs, side="B")

    key_labels = [kp["label"] for kp in key_pairs]
    value_labels = [vp["label"] for vp in value_pairs]

    key_cols = [f"KEY::{lbl}" for lbl in key_labels]

    a["__MATCH_KEY__"] = _build_hash_key_from_cols(a, key_cols)
    b["__MATCH_KEY__"] = _build_hash_key_from_cols(b, key_cols)

    a["__OCC__"] = a.groupby("__MATCH_KEY__", sort=False).cumcount() + 1
    b["__OCC__"] = b.groupby("__MATCH_KEY__", sort=False).cumcount() + 1

    merged = a.merge(
        b,
        on=["__MATCH_KEY__", "__OCC__"],
        how="outer",
        suffixes=("_A", "_B"),
        indicator=True,
        sort=False,
        copy=False,
    )

    for kp in key_pairs:
        label = kp["label"]
        a_raw = f"{kp['a']}_A" if f"{kp['a']}_A" in merged.columns else kp["a"]
        b_raw = f"{kp['b']}_B" if f"{kp['b']}_B" in merged.columns else kp["b"]

        s = _to_text(_safe_get_series(merged, a_raw, default_text=True))
        sb = _to_text(_safe_get_series(merged, b_raw, default_text=True))
        s = s.where(s.ne(""), sb)

        merged[f"DIM::{label}"] = s

    for vp in value_pairs:
        lbl = vp["label"]

        a_num = f"NUM::{lbl}_A" if f"NUM::{lbl}_A" in merged.columns else f"NUM::{lbl}"
        b_num = f"NUM::{lbl}_B" if f"NUM::{lbl}_B" in merged.columns else f"NUM::{lbl}"

        aval = _to_number(_safe_get_series(merged, a_num, default_text=False)).round(6)
        bval = _to_number(_safe_get_series(merged, b_num, default_text=False)).round(6)

        merged[f"VALOR::{lbl}::{base1_name}"] = aval
        merged[f"VALOR::{lbl}::{base2_name}"] = bval
        merged[f"DIF::{lbl}"] = (aval - bval).round(6)

    merged["PRESENCA"] = merged["_merge"].map(
        {
            "both": "Em ambas",
            "left_only": f"Somente {base1_name}",
            "right_only": f"Somente {base2_name}",
        }
    )

    dup_a = a.groupby("__MATCH_KEY__", sort=False).size().rename("QTD_A")
    dup_b = b.groupby("__MATCH_KEY__", sort=False).size().rename("QTD_B")
    dup = (
        dup_a.to_frame()
        .join(dup_b, how="outer")
        .fillna(0)
        .astype(int)
        .reset_index()
    )
    dup["DUPLICIDADE"] = np.where((dup["QTD_A"] > 1) | (dup["QTD_B"] > 1), 1, 0)

    merged = merged.merge(
        dup[["__MATCH_KEY__", "QTD_A", "QTD_B", "DUPLICIDADE"]],
        on="__MATCH_KEY__",
        how="left",
        sort=False,
        copy=False,
    )

    merged["QTD_A"] = merged["QTD_A"].fillna(0).astype(int)
    merged["QTD_B"] = merged["QTD_B"].fillna(0).astype(int)
    merged["DUPLICIDADE"] = merged["DUPLICIDADE"].fillna(0).astype(int)

    any_diff = pd.Series(0.0, index=merged.index)
    for lbl in value_labels:
        any_diff = any_diff.add(merged[f"DIF::{lbl}"].abs(), fill_value=0.0)

    merged["MOTIVO"] = np.select(
        [
            merged["PRESENCA"].eq(f"Somente {base1_name}"),
            merged["PRESENCA"].eq(f"Somente {base2_name}"),
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

    resumo_global_rows = []
    for lbl in value_labels:
        total_a = float(merged[f"VALOR::{lbl}::{base1_name}"].sum())
        total_b = float(merged[f"VALOR::{lbl}::{base2_name}"].sum())
        resumo_global_rows.append(
            {
                "Campo confrontado": lbl,
                f"Total {base1_name}": total_a,
                f"Total {base2_name}": total_b,
                "Diferença total": total_a - total_b,
            }
        )

    resumo_global = pd.DataFrame(resumo_global_rows)

    return {
        "full": merged,
        "resumo_global": resumo_global,
        "value_labels": value_labels,
        "key_labels": key_labels,
    }


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

    agg_map = {}
    for lbl in value_labels:
        agg_map[f"VALOR::{lbl}::{base1_name}"] = "sum"
        agg_map[f"VALOR::{lbl}::{base2_name}"] = "sum"
        agg_map[f"DIF::{lbl}"] = "sum"

    exec_df = df.groupby(group_cols, dropna=False).agg(agg_map).reset_index()

    motive = (
        df[df["MOTIVO"].ne("Conciliado")]
        .groupby(group_cols, dropna=False)
        .apply(_top_reason_from_df)
        .reset_index(name="Motivo predominante da diferença")
    )

    exec_df = exec_df.merge(motive, on=group_cols, how="left")
    exec_df["Motivo predominante da diferença"] = exec_df["Motivo predominante da diferença"].fillna("")

    rename_map = {c: c.replace("DIM::", "") for c in group_cols}
    for lbl in value_labels:
        rename_map[f"VALOR::{lbl}::{base1_name}"] = f"{lbl} {base1_name}"
        rename_map[f"VALOR::{lbl}::{base2_name}"] = f"{lbl} {base2_name}"
        rename_map[f"DIF::{lbl}"] = f"Diferença {lbl}"

    exec_df = exec_df.rename(columns=rename_map)

    ponte_rows = []
    for lbl in value_labels:
        ponte_rows.append(
            {
                "Agrupador": "TOTAL GERAL",
                "Campo confrontado": lbl,
                "Componente": f"Chave só na {base1_name}",
                "Valor": float(df.loc[df["MOTIVO"].eq(f"Chave só na {base1_name}"), f"DIF::{lbl}"].sum()),
            }
        )
        ponte_rows.append(
            {
                "Agrupador": "TOTAL GERAL",
                "Campo confrontado": lbl,
                "Componente": f"Chave só na {base2_name}",
                "Valor": float(df.loc[df["MOTIVO"].eq(f"Chave só na {base2_name}"), f"DIF::{lbl}"].sum()),
            }
        )
        ponte_rows.append(
            {
                "Agrupador": "TOTAL GERAL",
                "Campo confrontado": lbl,
                "Componente": f"Valor divergente entre {base1_name} e {base2_name}",
                "Valor": float(df.loc[df["MOTIVO"].eq(f"Valor divergente entre {base1_name} e {base2_name}"), f"DIF::{lbl}"].sum()),
            }
        )
        ponte_rows.append(
            {
                "Agrupador": "TOTAL GERAL",
                "Campo confrontado": lbl,
                "Componente": "Duplicidade",
                "Valor": float(df.loc[df["MOTIVO"].eq("Duplicidade"), f"DIF::{lbl}"].sum()),
            }
        )

        grp = (
            df[df["MOTIVO"].ne("Conciliado")]
            .groupby(group_cols + ["MOTIVO"], dropna=False)[f"DIF::{lbl}"]
            .sum()
            .reset_index()
        )

        for _, row in grp.iterrows():
            agrupador = " | ".join([_clean_text(row[c]) for c in group_cols])
            ponte_rows.append(
                {
                    "Agrupador": agrupador,
                    "Campo confrontado": lbl,
                    "Componente": row["MOTIVO"],
                    "Valor": float(row[f"DIF::{lbl}"]),
                }
            )

    ponte_df = pd.DataFrame(ponte_rows)

    detail = df[df["MOTIVO"].ne("Conciliado")].copy()

    detail_cols = []
    for lbl in results["key_labels"]:
        col = f"DIM::{lbl}"
        if col in detail.columns:
            detail_cols.append(col)

    value_cols = []
    for lbl in value_labels:
        value_cols.extend([
            f"VALOR::{lbl}::{base1_name}",
            f"VALOR::{lbl}::{base2_name}",
            f"DIF::{lbl}",
        ])

    detail = detail[detail_cols + value_cols + ["MOTIVO"]].copy()
    detail = detail.rename(columns=rename_map)

    for lbl in value_labels:
        detail = detail.rename(
            columns={
                f"VALOR::{lbl}::{base1_name}": f"{lbl} {base1_name}",
                f"VALOR::{lbl}::{base2_name}": f"{lbl} {base2_name}",
                f"DIF::{lbl}": f"Diferença {lbl}",
            }
        )

    if value_labels:
        diff_col = f"Diferença {value_labels[0]}"
        if diff_col in exec_df.columns:
            exec_df = (
                exec_df.assign(__ABS__=exec_df[diff_col].abs())
                .sort_values("__ABS__", ascending=False)
                .drop(columns=["__ABS__"])
            )
        if diff_col in detail.columns:
            detail = (
                detail.assign(__ABS__=detail[diff_col].abs())
                .sort_values("__ABS__", ascending=False)
                .drop(columns=["__ABS__"])
            )

    return exec_df, detail, ponte_df


# ============================================================
# Semântica de saída
# ============================================================

def _get_output_semantic_for_key_pair(row: dict):
    t = row.get("semantic_type", "texto")
    decimals = 0 if t in ["texto", "data"] else 2
    return {
        "tipo_logico": t,
        "casas_decimais": decimals,
        "formato_excel": row.get("excel_format", _excel_format_by_semantic_type(t, decimals)),
    }


def _get_output_semantic_for_value_pair(row: dict):
    t = row.get("semantic_type", "moeda")
    decimals = 2 if t in ["moeda", "percentual", "numero"] else 0
    return {
        "tipo_logico": t,
        "casas_decimais": decimals,
        "formato_excel": row.get("excel_format", _excel_format_by_semantic_type(t, decimals)),
    }


def _build_output_semantic_maps(
    key_pairs: List[dict],
    value_pairs: List[dict],
    group_labels: List[str],
    base1_name: str,
    base2_name: str,
) -> Dict[str, Dict[str, dict]]:
    key_meta = {kp["label"]: _get_output_semantic_for_key_pair(kp) for kp in key_pairs}
    value_meta = {vp["label"]: _get_output_semantic_for_value_pair(vp) for vp in value_pairs}

    resumo_exec = {}
    detalhe = {}
    top10 = {}

    for lbl in group_labels:
        resumo_exec[lbl] = key_meta.get(lbl, {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""})

    for lbl, meta in value_meta.items():
        resumo_exec[f"{lbl} {base1_name}"] = meta
        resumo_exec[f"{lbl} {base2_name}"] = meta
        resumo_exec[f"Diferença {lbl}"] = meta

        detalhe[f"{lbl} {base1_name}"] = meta
        detalhe[f"{lbl} {base2_name}"] = meta
        detalhe[f"Diferença {lbl}"] = meta

    for kp in key_pairs:
        detalhe[kp["label"]] = key_meta.get(kp["label"], {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""})
        top10[kp["label"]] = key_meta.get(kp["label"], {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""})

    if value_pairs:
        first_lbl = value_pairs[0]["label"]
        top10[f"Diferença {first_lbl}"] = value_meta.get(first_lbl, {"tipo_logico": "numero", "casas_decimais": 2, "formato_excel": "0.00"})

    detalhe["MOTIVO"] = {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""}
    top10["MOTIVO"] = {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""}
    resumo_exec["Motivo predominante da diferença"] = {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""}

    if value_pairs:
        main_meta = _get_output_semantic_for_value_pair(value_pairs[0])
    else:
        main_meta = {"tipo_logico": "numero", "casas_decimais": 2, "formato_excel": "0.00"}

    ponte = {
        "Agrupador": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""},
        "Campo confrontado": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""},
        "Componente": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""},
        "Valor": main_meta,
    }

    resumo_global_line = {}
    for vp in value_pairs:
        resumo_global_line[vp["label"]] = _get_output_semantic_for_value_pair(vp)

    resumo_global = {
        "Campo confrontado": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""},
        f"Total {base1_name}": main_meta,
        f"Total {base2_name}": main_meta,
        "Diferença total": main_meta,
        "__linha__": resumo_global_line,
    }

    metricas = {
        "Indicador": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""},
        "Valor": {"tipo_logico": "numero", "casas_decimais": 0, "formato_excel": "0"},
    }

    return {
        "resumo_exec": resumo_exec,
        "detalhe": detalhe,
        "top10": top10,
        "ponte": ponte,
        "resumo_global": resumo_global,
        "metricas": metricas,
    }


# ============================================================
# Excel helpers
# ============================================================

def _make_xlsx_format(wb, semantic_type: str, decimals: int, is_header: bool = False, is_diff: bool = False, is_total: bool = False):
    base = {"border": 1}

    if is_header:
        base.update({"bold": True, "bg_color": "#D9EAF7", "align": "center", "valign": "vcenter"})
        return wb.add_format(base)

    if is_total:
        base.update({"bold": True, "bg_color": "#FFF2CC"})

    if semantic_type == "moeda":
        base["num_format"] = _excel_format_by_semantic_type("moeda", decimals)
    elif semantic_type == "percentual":
        base["num_format"] = _excel_format_by_semantic_type("percentual", decimals)
    elif semantic_type == "numero":
        base["num_format"] = _excel_format_by_semantic_type("numero", decimals)
    elif semantic_type == "data":
        base["num_format"] = _excel_format_by_semantic_type("data", decimals)

    if is_diff:
        base.update({"font_color": "#C00000", "bold": True})

    return wb.add_format(base)


def _get_semantic_for_output_col(col: str, semantic_map: Dict[str, dict] = None) -> dict:
    if semantic_map and col in semantic_map:
        return semantic_map[col]

    uc = str(col).upper()
    if "DIFERENÇA" in uc:
        return {"tipo_logico": "numero", "casas_decimais": 2, "formato_excel": "0.00"}
    if any(x in uc for x in ["VALOR", "TOTAL", "IMPACTO"]):
        return {"tipo_logico": "moeda", "casas_decimais": 2, "formato_excel": "R$ #,##0.00"}
    if uc.startswith("QTD") or uc.startswith("QTDE") or "REGISTROS" in uc or "CAMPOS CONFRONTADOS" in uc or "ITENS EM DIVERGÊNCIA" in uc:
        return {"tipo_logico": "numero", "casas_decimais": 0, "formato_excel": "0"}
    if "TAXA" in uc or "COBERTURA" in uc:
        return {"tipo_logico": "percentual", "casas_decimais": 2, "formato_excel": "0.00%"}
    return {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""}


def _complete_semantic_map_for_df(df: pd.DataFrame, semantic_map: Dict[str, dict] = None) -> Dict[str, dict]:
    semantic_map = dict(semantic_map or {})
    completed = {}

    for col in df.columns:
        if col in semantic_map:
            completed[col] = semantic_map[col]
            continue

        uc = str(col).upper()
        nc = _norm_name(str(col))

        if "diferen" in nc:
            completed[col] = {
                "tipo_logico": "numero",
                "casas_decimais": 2,
                "formato_excel": "0.00",
            }
        elif any(x in uc for x in ["VALOR", "TOTAL", "IMPACTO"]):
            completed[col] = {
                "tipo_logico": "moeda",
                "casas_decimais": 2,
                "formato_excel": "R$ #,##0.00",
            }
        elif "TAXA" in uc or "COBERTURA" in uc:
            completed[col] = {
                "tipo_logico": "percentual",
                "casas_decimais": 2,
                "formato_excel": "0.00%",
            }
        elif uc.startswith("QTD") or uc.startswith("QTDE") or "REGISTROS" in uc or "CAMPOS CONFRONTADOS" in uc or "ITENS EM DIVERGÊNCIA" in uc:
            completed[col] = {
                "tipo_logico": "numero",
                "casas_decimais": 0,
                "formato_excel": "0",
            }
        elif pd.api.types.is_numeric_dtype(df[col]):
            completed[col] = {
                "tipo_logico": "numero",
                "casas_decimais": 2,
                "formato_excel": "0.00",
            }
        else:
            completed[col] = {
                "tipo_logico": "texto",
                "casas_decimais": 0,
                "formato_excel": "",
            }

    return completed


def _set_column_formats(writer, sheet_name: str, df: pd.DataFrame, semantic_map: Dict[str, dict] = None):
    wb = writer.book
    ws = writer.sheets[sheet_name]
    fmt_head = _make_xlsx_format(wb, "texto", 0, is_header=True)

    final_map = _complete_semantic_map_for_df(df, semantic_map)

    for i, col in enumerate(df.columns):
        ws.write(0, i, col, fmt_head)
        ser = df[col]
        meta = final_map[col]
        semantic_type = meta["tipo_logico"]
        decimals = int(meta.get("casas_decimais", 0))
        is_diff = "diferen" in _norm_name(str(col))
        width = max(len(str(col)), min(60, ser.astype(str).map(len).max() if len(ser) else 10)) + 2
        fmt = _make_xlsx_format(wb, semantic_type, decimals, is_diff=is_diff)
        ws.set_column(i, i, width, fmt)

    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(df), max(0, len(df.columns) - 1))


def _add_total_row(writer, sheet_name: str, df: pd.DataFrame, semantic_map: Dict[str, dict] = None, skip_when_only_total_geral: bool = False):
    if df.empty:
        return

    if skip_when_only_total_geral and "Agrupador" in df.columns:
        vals = {_clean_text(v) for v in df["Agrupador"].dropna().astype(str).unique().tolist() if _clean_text(v)}
        if vals and vals == {"TOTAL GERAL"}:
            return

    wb = writer.book
    ws = writer.sheets[sheet_name]
    total_row = len(df) + 1
    fmt_total_txt = _make_xlsx_format(wb, "texto", 0, is_total=True)

    final_map = _complete_semantic_map_for_df(df, semantic_map)

    label = "TOTAL FILTRADO"
    if skip_when_only_total_geral and "Agrupador" in df.columns:
        tem_detalhe = any(_clean_text(v) != "TOTAL GERAL" for v in df["Agrupador"].dropna().astype(str).tolist())
        if not tem_detalhe:
            return
        label = "TOTAL FILTRADO (visão filtrada)"

    ws.write(total_row, 0, label, fmt_total_txt)

    for col_idx, col in enumerate(df.columns[1:], start=1):
        if pd.api.types.is_numeric_dtype(df[col]):
            col_letter = chr(65 + col_idx) if col_idx < 26 else None
            if col_letter:
                meta = final_map[col]
                fmt = _make_xlsx_format(
                    wb,
                    meta["tipo_logico"],
                    int(meta.get("casas_decimais", 0)),
                    is_total=True,
                    is_diff=("diferen" in _norm_name(str(col))),
                )
                ws.write_formula(total_row, col_idx, f"=SUBTOTAL(109,{col_letter}2:{col_letter}{len(df)+1})", fmt)


def _write_dataframe_block(ws, wb, start_row: int, start_col: int, title: str, df: pd.DataFrame, semantic_map: Dict[str, dict] = None):
    header_fmt = wb.add_format({"bold": True, "bg_color": "#1F2A44", "font_color": "#FFFFFF", "border": 1, "align": "left", "valign": "vcenter"})
    colhead_fmt = _make_xlsx_format(wb, "texto", 0, is_header=True)

    width = max(1, len(df.columns))
    ws.merge_range(start_row, start_col, start_row, start_col + width - 1, title, header_fmt)

    if df.empty:
        ws.write(start_row + 1, start_col, "Sem dados para exibir.")
        return start_row + 2

    final_map = _complete_semantic_map_for_df(df, semantic_map)

    for j, col in enumerate(df.columns):
        ws.write(start_row + 1, start_col + j, col, colhead_fmt)
        max_len = len(str(col))
        meta = final_map[col]
        fmt = _make_xlsx_format(
            wb,
            meta["tipo_logico"],
            int(meta.get("casas_decimais", 0)),
            is_diff=("diferen" in _norm_name(str(col))),
        )

        for i, val in enumerate(df[col].tolist(), start=0):
            cell_row = start_row + 2 + i

            if meta["tipo_logico"] in ["moeda", "numero", "percentual"] and pd.notna(val) and str(val) != "":
                try:
                    ws.write_number(cell_row, start_col + j, float(val), fmt)
                except Exception:
                    ws.write(cell_row, start_col + j, val, fmt)
            else:
                ws.write(cell_row, start_col + j, val, fmt)

            max_len = max(max_len, len(str(val)))

        ws.set_column(start_col + j, start_col + j, min(max_len + 2, 32))

    return start_row + 2 + len(df)


# ============================================================
# Indicadores executivos
# ============================================================

def _build_resumo_metricas(
    resumo_global: pd.DataFrame,
    detalhe: pd.DataFrame,
    ponte: pd.DataFrame,
    resultado_full: pd.DataFrame,
    base1_name: str,
    base2_name: str,
):
    qtd_campos = int(resumo_global["Campo confrontado"].nunique()) if "Campo confrontado" in resumo_global.columns else 0
    qtd_diverg = int(len(detalhe))

    qtd_dup = int(detalhe["MOTIVO"].eq("Duplicidade").sum()) if (not detalhe.empty and "MOTIVO" in detalhe.columns) else 0
    qtd_so_b1 = int(detalhe["MOTIVO"].eq(f"Chave só na {base1_name}").sum()) if (not detalhe.empty and "MOTIVO" in detalhe.columns) else 0
    qtd_so_b2 = int(detalhe["MOTIVO"].eq(f"Chave só na {base2_name}").sum()) if (not detalhe.empty and "MOTIVO" in detalhe.columns) else 0

    qtd_reg_b1 = int(resultado_full["PRESENCA"].isin(["Em ambas", f"Somente {base1_name}"]).sum()) if (not resultado_full.empty and "PRESENCA" in resultado_full.columns) else 0
    qtd_reg_b2 = int(resultado_full["PRESENCA"].isin(["Em ambas", f"Somente {base2_name}"]).sum()) if (not resultado_full.empty and "PRESENCA" in resultado_full.columns) else 0
    qtd_conc = int(resultado_full["MOTIVO"].eq("Conciliado").sum()) if (not resultado_full.empty and "MOTIVO" in resultado_full.columns) else 0

    base_ref = max(qtd_reg_b1, qtd_reg_b2, 1)
    taxa_conc = qtd_conc / base_ref
    taxa_div = qtd_diverg / base_ref

    metricas = pd.DataFrame(
        [
            {"Indicador": "Campos confrontados", "Valor": qtd_campos},
            {"Indicador": f"Registros {base1_name}", "Valor": qtd_reg_b1},
            {"Indicador": f"Registros {base2_name}", "Valor": qtd_reg_b2},
            {"Indicador": "Itens em divergência", "Valor": qtd_diverg},
            {"Indicador": "Qtd. em duplicidade", "Valor": qtd_dup},
            {"Indicador": f"Qtd. só na {base1_name}", "Valor": qtd_so_b1},
            {"Indicador": f"Qtd. só na {base2_name}", "Valor": qtd_so_b2},
            {"Indicador": "Qtd. conciliados", "Valor": qtd_conc},
            {"Indicador": "Taxa de conciliação", "Valor": taxa_conc},
            {"Indicador": "Taxa de divergência", "Valor": taxa_div},
        ]
    )

    return metricas

def _build_metricas_semantic_map(metricas: pd.DataFrame) -> Dict[str, dict]:
    out = {}
    for _, row in metricas.iterrows():
        ind = str(row["Indicador"])
        if ind in [
            "Campos confrontados",
            "Itens em divergência",
            "Qtd. em duplicidade",
            "Qtd. conciliados",
        ] or ind.startswith("Registros ") or ind.startswith("Qtd. só na "):
            out[ind] = {"tipo_logico": "numero", "casas_decimais": 0, "formato_excel": "0"}
        elif ind in ["Taxa de conciliação", "Taxa de divergência"]:
            out[ind] = {"tipo_logico": "percentual", "casas_decimais": 2, "formato_excel": "0.00%"}
        else:
            out[ind] = {"tipo_logico": "numero", "casas_decimais": 0, "formato_excel": "0"}

    return {
        "Indicador": {"tipo_logico": "texto", "casas_decimais": 0, "formato_excel": ""},
        "Valor": {"tipo_logico": "numero", "casas_decimais": 0, "formato_excel": "0"},
        "__linha__": out,
    }

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


# ============================================================
# Export Excel
# ============================================================

def _write_resumo_global_block(ws, wb, start_row: int, start_col: int, title: str, df: pd.DataFrame, semantic_map: Dict[str, dict]):
    header_fmt = wb.add_format({"bold": True, "bg_color": "#1F2A44", "font_color": "#FFFFFF", "border": 1, "align": "left", "valign": "vcenter"})
    colhead_fmt = _make_xlsx_format(wb, "texto", 0, is_header=True)
    text_fmt = _make_xlsx_format(wb, "texto", 0)

    width = max(1, len(df.columns))
    ws.merge_range(start_row, start_col, start_row, start_col + width - 1, title, header_fmt)

    if df.empty:
        ws.write(start_row + 1, start_col, "Sem dados para exibir.")
        return start_row + 2

    linha_map = semantic_map.get("__linha__", {})

    for j, col in enumerate(df.columns):
        ws.write(start_row + 1, start_col + j, col, colhead_fmt)

    for i, row in df.reset_index(drop=True).iterrows():
        cell_row = start_row + 2 + i
        campo = str(row.get("Campo confrontado", ""))
        meta = linha_map.get(campo, {"tipo_logico": "numero", "casas_decimais": 2})

        for j, col in enumerate(df.columns):
            val = row[col]
            if j == 0:
                ws.write(cell_row, start_col + j, val, text_fmt)
                continue

            fmt = _make_xlsx_format(
                wb,
                meta.get("tipo_logico", "numero"),
                int(meta.get("casas_decimais", 2)),
                is_diff=("diferen" in _norm_name(str(col))),
            )

            if meta.get("tipo_logico") in ["moeda", "numero", "percentual"] and pd.notna(val) and str(val) != "":
                try:
                    ws.write_number(cell_row, start_col + j, float(val), fmt)
                except Exception:
                    ws.write(cell_row, start_col + j, val, fmt)
            else:
                ws.write(cell_row, start_col + j, val, fmt)

    ws.set_column(start_col, start_col, 28)
    ws.set_column(start_col + 1, start_col + 3, 18)
    return start_row + 2 + len(df)


def _write_metricas_block(ws, wb, start_row: int, start_col: int, title: str, df: pd.DataFrame, semantic_map: Dict[str, dict]):
    header_fmt = wb.add_format({"bold": True, "bg_color": "#1F2A44", "font_color": "#FFFFFF", "border": 1, "align": "left", "valign": "vcenter"})
    colhead_fmt = _make_xlsx_format(wb, "texto", 0, is_header=True)
    text_fmt = _make_xlsx_format(wb, "texto", 0)

    width = max(1, len(df.columns))
    ws.merge_range(start_row, start_col, start_row, start_col + width - 1, title, header_fmt)

    if df.empty:
        ws.write(start_row + 1, start_col, "Sem dados para exibir.")
        return start_row + 2

    linha_map = semantic_map.get("__linha__", {})

    for j, col in enumerate(df.columns):
        ws.write(start_row + 1, start_col + j, col, colhead_fmt)

    for i, row in df.reset_index(drop=True).iterrows():
        cell_row = start_row + 2 + i
        indicador = str(row["Indicador"])
        meta = linha_map.get(indicador, {"tipo_logico": "texto", "casas_decimais": 0})

        ws.write(cell_row, start_col, indicador, text_fmt)

        fmt = _make_xlsx_format(wb, meta["tipo_logico"], int(meta.get("casas_decimais", 0)))
        val = row["Valor"]

        if meta["tipo_logico"] in ["moeda", "numero", "percentual"] and pd.notna(val):
            try:
                ws.write_number(cell_row, start_col + 1, float(val), fmt)
            except Exception:
                ws.write(cell_row, start_col + 1, val, fmt)
        else:
            ws.write(cell_row, start_col + 1, val, fmt)

    ws.set_column(start_col, start_col, 34)
    ws.set_column(start_col + 1, start_col + 1, 18)

    return start_row + 2 + len(df)


def _export_excel(
    results: Dict[str, pd.DataFrame],
    resumo_exec: pd.DataFrame,
    detalhe: pd.DataFrame,
    ponte: pd.DataFrame,
    output_semantics: Dict[str, Dict[str, dict]],
    base1_name: str,
    base2_name: str,
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

        ws.write(0, 0, "Central de Análises — Resumo Executivo", title_fmt)
        ws.write(1, 0, pd.Timestamp.now().strftime("Gerado em %d/%m/%Y %H:%M:%S"), sub_fmt)

        row_a = _write_resumo_global_block(
            ws, wb, 3, 0, "Fechamento global dos campos confrontados", resumo_global,
            semantic_map=output_semantics.get("resumo_global", {}),
        )

        row_b = _write_metricas_block(
            ws, wb, 3, 6, "Indicadores executivos", metricas,
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

        ws.freeze_panes(4, 0)

        detalhe.to_excel(writer, sheet_name="Divergencias", index=False)
        _set_column_formats(writer, "Divergencias", detalhe, semantic_map=output_semantics.get("detalhe", {}))
        _add_total_row(writer, "Divergencias", detalhe, semantic_map=output_semantics.get("detalhe", {}))

        ponte_export = ponte.copy()
        if "Agrupador" in ponte_export.columns:
            ponte_export["Agrupador"] = ponte_export["Agrupador"].astype(str)
            ponte_export.loc[ponte_export["Agrupador"].eq("TOTAL GERAL"), "Agrupador"] = " TOTAL GERAL"

        ponte_export.to_excel(writer, sheet_name="Ponte_Conciliacao", index=False)
        _set_column_formats(writer, "Ponte_Conciliacao", ponte_export, semantic_map=output_semantics.get("ponte", {}))
        _add_total_row(writer, "Ponte_Conciliacao", ponte_export, semantic_map=output_semantics.get("ponte", {}), skip_when_only_total_geral=True)

    return bio.getvalue()


# ============================================================
# UI helpers
# ============================================================

def _render_header_row(labels: List[str]):
    cols = st.columns([1.15, 1.15, 0.9, 0.7, 0.9])
    for col, label in zip(cols, labels):
        with col:
            st.markdown(f"**{label}**")


def _render_pair_line(df_a, df_b, cols_a, cols_b, row, idx, prefix: str, is_value: bool = False):
    cols = st.columns([1.15, 1.15, 0.9, 0.7, 0.9])

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
            f"Tipo lógico #{idx+1}",
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
            f"Formato Excel #{idx+1}",
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

    st.title("Central de Análises")
    st.subheader("Análise de Consistência entre Bases")
    st.caption("Confronta valores entre duas bases, fecha totais, destaca divergências e evidencia onde estão as diferenças.")

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

    with st.spinner("Carregando bases..."):
        df_a = _read_file_cached(up_a.getvalue(), up_a.name)
        df_b = _read_file_cached(up_b.getvalue(), up_b.name)

    cols_a = _get_column_list(tuple(df_a.columns))
    cols_b = _get_column_list(tuple(df_b.columns))

    st.subheader("2) Campos que identificam o mesmo registro nas duas bases")
    st.caption("Nesta etapa você define as chaves e já registra o tipo lógico e o formato esperado no Excel.")
    _render_header_row(["Base 1", "Base 2", "Nome da dimensão", "Tipo lógico", "Formato Excel"])

    for i, row in enumerate(st.session_state["cm_key_rows"]):
        st.session_state["cm_key_rows"][i] = _render_pair_line(
            df_a, df_b, cols_a, cols_b, row, i, prefix="cm_key", is_value=False
        )

    if st.button("Adicionar par-chave"):
        st.session_state["cm_key_rows"].append({
            "a": "", "b": "", "label": "",
            "semantic_type": "",
            "excel_format": "",
            "fmt_manual": False,
            "last_signature": "",
        })
        st.rerun()

    st.subheader("3) Quais campos deseja confrontar para validar")
    st.caption("Aqui você escolhe os valores que serão confrontados e já define o tipo lógico e o formato de saída.")
    _render_header_row(["Valor Base 1", "Valor Base 2", "Nome do valor", "Tipo lógico", "Formato Excel"])

    for i, row in enumerate(st.session_state["cm_val_rows"]):
        st.session_state["cm_val_rows"][i] = _render_pair_line(
            df_a, df_b, cols_a, cols_b, row, i, prefix="cm_val", is_value=True
        )

    if st.button("Adicionar campo de valor"):
        st.session_state["cm_val_rows"].append({
            "a": "", "b": "", "label": "",
            "semantic_type": "",
            "excel_format": "",
            "fmt_manual": False,
            "last_signature": "",
        })
        st.rerun()

    st.subheader("4) Como deseja receber o resultado?")
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

    st.subheader("5) Processar análise")
    executar = st.button("Executar análise")

    if executar:
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

        with st.spinner("Processando análise..."):
            t0 = time.perf_counter()

            results = _run_reconciliation(df_a, df_b, valid_keys, valid_vals, base1_name, base2_name)
            exec_df, detail_df, ponte_df = _build_executive_and_detail(
                results,
                group_labels,
                total_labels or default_total,
                base1_name,
                base2_name,
            )

            output_semantics = _build_output_semantic_maps(
                key_pairs=valid_keys,
                value_pairs=valid_vals,
                group_labels=group_labels if group_labels else [valid_keys[0]["label"]],
                base1_name=base1_name,
                base2_name=base2_name,
            )

            excel = _export_excel(
                results,
                exec_df,
                detail_df,
                ponte_df,
                output_semantics,
                base1_name,
                base2_name,
            )

            elapsed = time.perf_counter() - t0

        st.success(f"Análise concluída em {elapsed:.2f}s.")
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
            file_name="Central_Analises_V31.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
