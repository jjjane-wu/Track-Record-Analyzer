"""
mapper.py — Semantic column matching and per-GP config save/load.

Three-layer detection:
  1. Structural: exact / near-exact name match (after normalisation)
  2. Value-pattern: infer field from data distribution
  3. Semantic: sentence-transformer cosine similarity (offline, ~90 MB model)

Configs are saved as JSON in ../configs/<gp_slug>_config.json.
"""

from __future__ import annotations

import json
import os
import re
import unicodedata
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

# ── Template field catalogue ───────────────────────────────────────────────────
# Each entry describes one target field: its canonical name, aliases for rule-
# based matching, and a value-pattern hint.
TEMPLATE_FIELDS: list[dict] = [
    {"id": "fund",            "label": "Fund",            "aliases": ["fund", "fund name", "vehicle"]},
    {"id": "company",         "label": "Company",         "aliases": ["company", "portfolio company", "company name", "portco", "target company"]},
    {"id": "region",          "label": "Geography",       "aliases": ["region", "geography", "country", "location", "hq", "domicile", "headquarters", "hq country", "country of domicile"]},
    {"id": "sector",          "label": "Sector",          "aliases": ["sector", "industry", "gics", "classification"]},
    {"id": "entry_date",      "label": "Entry Date",      "aliases": ["entry date", "investment date", "acquisition date", "inv date", "entry", "investitionsdatum", "datum der akquisition"],       "value_hint": "date"},
    {"id": "exit_date",       "label": "Exit Date",       "aliases": ["exit date", "exit / valuation date", "exit valuation date", "valuation date", "exit"],                                        "value_hint": "date"},
    {"id": "status",          "label": "Status",          "aliases": ["status", "deal status", "realization status"]},
    {"id": "role",            "label": "GP Role",         "aliases": ["role", "gp role", "fund role", "lead", "co-invest"]},
    {"id": "transaction_type","label": "Transaction Type","aliases": ["source", "transaction type", "deal type", "type", "deal source"]},
    {"id": "competition",     "label": "Process Type",    "aliases": ["competition", "process type", "process", "auction type", "sourcing process"]},
    {"id": "sourcing_partner","label": "Sourcing Partner","aliases": ["sourcing partner", "advisory partner", "responsible partner", "partner"]},
    {"id": "exit_type",       "label": "Exit Type",       "aliases": ["exit type", "type of exit", "exit route"]},
    {"id": "holding_period",  "label": "Hold Period",     "aliases": ["holding period", "hold period", "years held", "duration"],                                                                     "value_hint": "decimal_years"},
    {"id": "ic_initial",  "label": "Initial Fund Equity Invested (m)", "aliases": ["initial fund equity invested", "initial equity invested", "initial invested", "equity invested initial", "initial equity", "initial fund equity"]},
    {"id": "ic_total",    "label": "Total Fund Equity Invested (m)",   "aliases": ["total fund equity invested", "total invested", "total equity invested", "invested capital", "investment cost", "cost of investment", "current investment cost", "final investment cost"]},
    {"id": "fund_currency", "label": "Deal Currency",                    "aliases": ["fund currency", "currency", "ccy", "fund ccy", "local currency", "reporting currency", "deal currency"]},
    {"id": "realized",    "label": "Realized Value (m)",                "aliases": ["realized value", "realised value", "realization proceeds", "exit proceeds", "total realized", "realization"]},
    {"id": "unrealized",  "label": "Unrealized Value (m)",              "aliases": ["unrealized value", "unrealised value", "current value", "fair value", "nav", "current equity value"]},
    {"id": "total_value", "label": "Total Value (m)",                   "aliases": ["total value", "total fund value", "gross value"]},
    {"id": "gross_moic",      "label": "Gross MOIC",      "aliases": ["gross moic", "gross mom", "gross multiple", "moic", "mom", "gross mv/ic", "gross mv ic", "total moic", "multiple of cost", "money multiple", "return multiple", "cost multiple", "gross return multiple"],  "value_hint": "moic"},
    {"id": "gross_irr",       "label": "Gross IRR",       "aliases": ["gross irr", "irr", "gross return"],                                                                                           "value_hint": "percent"},
    {"id": "entry_rev",   "label": "Entry LTM Revenue (m)",  "aliases": ["ltm sales", "revenue", "entry revenue", "entry ltm revenue", "entry ltm sales", "sales at acquisition", "ltm sales k"]},
    {"id": "entry_ebitda","label": "Entry LTM EBITDA (m)",   "aliases": ["ltm ebitda", "ebitda", "entry ebitda", "entry ltm ebitda", "ebitda at acquisition", "ltm ebitda k"]},
    {"id": "entry_net_debt","label":"Entry Net Debt (m)",     "aliases": ["net debt", "entry net debt", "net debt at acquisition", "net debt i"]},
    {"id": "entry_ev",    "label": "Entry Enterprise Value (m)", "aliases": ["enterprise value", "ev", "entry ev", "entry enterprise value", "ev at acquisition", "tev", "enterprise value j"]},
    {"id": "exit_rev",    "label": "Exit LTM Revenue (m)",   "aliases": ["exit ltm sales", "exit revenue", "exit ltm revenue", "current ltm sales", "ltm sales exit"]},
    {"id": "exit_ebitda", "label": "Exit LTM EBITDA (m)",    "aliases": ["exit ltm ebitda", "exit ebitda", "current ltm ebitda", "ltm ebitda exit"]},
    {"id": "exit_net_debt","label":"Exit Net Debt (m)",       "aliases": ["exit net debt", "current net debt", "net debt exit", "current valuation net debt"]},
    {"id": "exit_ev",     "label": "Exit Enterprise Value (m)", "aliases": ["exit enterprise value", "exit ev", "current enterprise value", "current ev", "enterprise value r"]},
    {"id": "valuation_method","label": "Valuation Method",           "aliases": ["valuation method", "unrealized valuation methodology", "methodology", "valuation approach"]},
    {"id": "fund_ownership",  "label": "Fund Ownership %",           "aliases": ["fund ownership", "ownership", "ownership pct", "initial fund ownership"]},
    {"id": "no_of_seats",     "label": "Board Seats",                "aliases": ["no of seats", "board seats", "seats"]},
    {"id": "coi_deal",        "label": "COI Deal",                    "aliases": ["coi", "coi deal", "co-investment", "coinvestment", "carry over", "carried over"]},
]

FIELD_IDS = [f["id"] for f in TEMPLATE_FIELDS]

# Fields commonly absent from GP files and not required for a valid output
# template (the template leaves them blank or auto-computes them). When one of
# these is unmapped, it is expected — the UI should not flag it for review.
# Everything NOT listed here is treated as a core/required field.
OPTIONAL_FIELD_IDS: set[str] = {
    "ic_initial", "total_value", "holding_period", "gross_irr",
    "valuation_method", "no_of_seats", "coi_deal", "role", "competition",
    "sourcing_partner", "transaction_type", "exit_type", "fund_ownership",
    "entry_rev", "entry_ebitda", "entry_net_debt", "entry_ev",
    "exit_rev", "exit_ebitda", "exit_net_debt", "exit_ev",
}


def is_optional(field_id: str) -> bool:
    """True when an unmapped field is expected/acceptable (not required)."""
    return field_id in OPTIONAL_FIELD_IDS


def _slug(name: str) -> str:
    """Make a filesystem-safe slug from a GP name."""
    s = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")
    return s or "unknown_gp"


def config_path(gp_name: str, configs_dir: str | Path = "../configs") -> Path:
    return Path(configs_dir) / f"{_slug(gp_name)}_config.json"


def load_config(gp_name: str, configs_dir: str | Path = "../configs") -> dict | None:
    p = config_path(gp_name, configs_dir)
    if p.exists():
        try:
            return json.loads(p.read_text())
        except Exception:
            return None
    return None


def save_config(gp_name: str, mapping: dict, meta: dict | None = None,
                configs_dir: str | Path = "../configs") -> None:
    p = config_path(gp_name, configs_dir)
    p.parent.mkdir(parents=True, exist_ok=True)
    payload = {"mapping": mapping, "meta": meta or {}}
    p.write_text(json.dumps(payload, indent=2, default=str))


# ── Rule-based matching ────────────────────────────────────────────────────────

def _normalise(text: str) -> str:
    s = str(text).lower()
    s = re.sub(r"\([a-z0-9]+\)", "", s)   # remove footnotes
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def rule_match(col_name: str) -> str | None:
    """
    Return field_id if col_name matches any known alias, else None.
    Longest alias match wins (avoids 'net debt' matching before 'entry net debt').
    """
    norm = _normalise(col_name)
    best_id, best_len = None, 0
    for field in TEMPLATE_FIELDS:
        for alias in field["aliases"]:
            a = _normalise(alias)
            if a in norm or norm in a:
                if len(a) > best_len:
                    best_id = field["id"]
                    best_len = len(a)
    return best_id


# ── Value-pattern matching ─────────────────────────────────────────────────────

def _sample_values(series: pd.Series, n: int = 50) -> list:
    return series.dropna().head(n).tolist()


def value_pattern_match(series: pd.Series) -> str | None:
    """Infer field type from the values in a column."""
    from datetime import datetime, date
    samples = _sample_values(series)
    if not samples:
        return None

    n = len(samples)

    # Date pattern
    date_count = 0
    for v in samples:
        if isinstance(v, (datetime, date)):
            date_count += 1
        elif isinstance(v, str) and re.match(r"\d{4}-\d{2}-\d{2}", v.strip()):
            date_count += 1
    if date_count / n > 0.6:
        return "entry_date"  # ambiguous; name-matching will disambiguate

    # MOIC pattern: values like "2.5x" or float 1.0–20.0
    moic_count = 0
    for v in samples:
        sv = str(v).strip()
        if re.match(r"^\d+\.?\d*x?$", sv):
            try:
                f = float(sv.rstrip("x"))
                if 0.0 < f < 25.0:
                    moic_count += 1
            except (ValueError, TypeError):
                pass
    if moic_count / n > 0.7:
        return "gross_moic"

    # IRR pattern: small percentages (0–100)
    pct_count = 0
    for v in samples:
        try:
            f = float(str(v).strip().rstrip("%"))
            if -0.5 < f < 1.5:  # raw fraction form
                pct_count += 1
        except (ValueError, TypeError):
            pass
    if pct_count / n > 0.7:
        return "gross_irr"

    return None


# ── Semantic matching with sentence-transformers ───────────────────────────────

_model = None  # lazy-loaded


def _get_model():
    global _model
    if _model is None:
        from sentence_transformers import SentenceTransformer
        _model = SentenceTransformer("paraphrase-multilingual-MiniLM-L12-v2")
    return _model


def semantic_match(col_names: list[str], threshold_auto: float = 0.85,
                   threshold_review: float = 0.50) -> dict[str, tuple[str | None, float]]:
    """
    Return {col_name: (field_id | None, score)} for each input column.
    Scores are cosine similarities in [0, 1].
    """
    from sklearn.metrics.pairwise import cosine_similarity

    model = _get_model()

    field_labels = [f["label"] + " " + " ".join(f["aliases"]) for f in TEMPLATE_FIELDS]
    field_vecs = model.encode(field_labels, normalize_embeddings=True)
    col_vecs = model.encode(col_names, normalize_embeddings=True)

    sims = cosine_similarity(col_vecs, field_vecs)  # (n_cols, n_fields)

    result = {}
    for i, col in enumerate(col_names):
        best_j = int(np.argmax(sims[i]))
        score = float(sims[i][best_j])
        if score >= threshold_review:
            result[col] = (FIELD_IDS[best_j], score)
        else:
            result[col] = (None, score)
    return result


# ── Main matcher ───────────────────────────────────────────────────────────────

def build_mapping(
    df: pd.DataFrame,
    gp_name: str = "",
    use_semantic: bool = True,
    configs_dir=None,   # kept for backwards-compat, no longer used
) -> dict[str, Any]:
    """
    Return a mapping dict:
      {
        "field_to_col": {field_id: col_name_or_None},
        "col_confidence": {col_name: float},
        "source": "rules" | "semantic" | "mixed",
      }

    Always runs fresh three-layer detection — no config loading or saving.
    """

    col_names = [c for c in df.columns if str(c).strip()]
    field_to_col: dict[str, str | None] = {fid: None for fid in FIELD_IDS}
    col_confidence: dict[str, float] = {}

    # ── Layer 1: rule-based ────────────────────────────────────────────
    for col in col_names:
        fid = rule_match(col)
        if fid and field_to_col[fid] is None:
            field_to_col[fid] = col
            col_confidence[col] = 0.95

    # ── Layer 2: value pattern (only for still-unmapped fields) ────────
    unmapped_cols = [c for c in col_names if c not in col_confidence]
    for col in unmapped_cols:
        fid = value_pattern_match(df[col])
        if fid and field_to_col[fid] is None:
            field_to_col[fid] = col
            col_confidence[col] = 0.75

    # ── Layer 3: semantic ──────────────────────────────────────────────
    if use_semantic:
        still_unmapped = [c for c in col_names if c not in col_confidence]
        if still_unmapped:
            try:
                sem = semantic_match(still_unmapped)
                for col, (fid, score) in sem.items():
                    if fid and field_to_col[fid] is None:
                        field_to_col[fid] = col
                        col_confidence[col] = score
            except Exception:
                pass  # semantic model unavailable; rule-only is fine

    return {
        "field_to_col": field_to_col,
        "col_confidence": col_confidence,
        "source": "semantic" if use_semantic else "rules",
    }


def confidence_tier(score: float) -> str:
    """Map a confidence score to a display tier."""
    if score >= 0.85:
        return "green"
    if score >= 0.50:
        return "yellow"
    return "red"


def mapping_from_col_indices(df: pd.DataFrame, index_map: dict[str, int]) -> dict[str, str | None]:
    """Convert {field_id: col_index} to {field_id: col_name}."""
    cols = list(df.columns)
    result = {}
    for fid, idx in index_map.items():
        if idx is not None and 0 <= idx < len(cols):
            result[fid] = cols[idx]
        else:
            result[fid] = None
    return result
