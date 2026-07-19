"""
inferencer.py — Stage 3: Schema Inference

Maps raw DataFrame columns to the standardised PE deal schema using
multiple independent signals, combined transparently.

Signal priority (highest first):
  1. Alias exact/substring match   — definitive name evidence    (0.85–0.95)
  2. Regex pattern matching        — structural name patterns    (0.70–0.90)
  3. Fuzzy token overlap           — partial name similarity     (0.55–0.85)
  4. Column profile heuristics     — distribution-based boost    (+0–0.20)
  5. Semantic embedding similarity — broad similarity fallback   (0.50–0.80)
  6. LLM fallback                  — for genuinely ambiguous cols (optional)

Design principle: signals are INDEPENDENT. If two signals disagree, the
higher-confidence one wins; they never blindly reinforce each other.
Signals from different tiers are not summed — the final confidence is the
maximum of any tier, with a small additive boost from corroborating tiers.

If no field reaches CONFIDENCE_REVIEW (0.50), the column is left unmapped.
The parser never silently forces a mapping it is not confident about.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from typing import Any, Optional, TYPE_CHECKING

import numpy as np
import pandas as pd

from extractor import ExtractedTable, ColumnProfile
from mapper import TEMPLATE_FIELDS, FIELD_IDS     # reuse field catalogue

if TYPE_CHECKING:
    from llm_interface import LLMInterface

# ═══════════════════════════════════════════════════════════════════════════════
# Confidence thresholds
# ═══════════════════════════════════════════════════════════════════════════════

CONFIDENCE_AUTO    = 0.85   # green — auto-confirm, skip review
CONFIDENCE_REVIEW  = 0.50   # yellow — present to user for confirmation
# below CONFIDENCE_REVIEW → unmapped (red)


# ═══════════════════════════════════════════════════════════════════════════════
# Data Structures
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class FieldInference:
    """Result of inferring one standardised field."""
    field_id: str
    field_label: str
    source_col: Optional[str]   # DataFrame column name, or None if unmapped
    confidence: float           # 0.0 – 1.0
    signals: list[str]          # human-readable explanation of why


@dataclass
class SchemaInference:
    """
    Output of Stage 3.

    Contains a complete mapping from standardised field IDs to DataFrame
    column names, with confidence scores and explanations for every field.
    Unmapped columns and fields are explicitly listed.
    """
    # Core mapping
    field_to_col: dict[str, Optional[str]]   # {field_id: col_name | None}
    col_to_field: dict[str, Optional[str]]   # {col_name: field_id | None}

    # Confidence
    confidences: dict[str, float]            # {field_id: 0.0–1.0}

    # Explanations (for the review screen)
    explanations: dict[str, list[str]]       # {field_id: [signal1, signal2, …]}

    # Status buckets
    auto_confirmed: list[str]                # field_ids with confidence ≥ 0.85
    needs_review: list[str]                  # 0.50 ≤ confidence < 0.85
    unmapped_fields: list[str]               # field_ids with confidence < 0.50
    unmapped_cols: list[str]                 # column names not mapped to any field

    # Compatibility: expose col_confidence keyed by column name (like mapper.py)
    @property
    def col_confidence(self) -> dict[str, float]:
        return {
            col: self.confidences.get(fid, 0.0)
            for fid, col in self.field_to_col.items()
            if col is not None
        }


# ═══════════════════════════════════════════════════════════════════════════════
# Normalisation helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _norm(text: str) -> str:
    """Lowercase, strip footnote markers, collapse whitespace/underscores."""
    s = unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode()
    s = re.sub(r"\([a-z]+\)", " ", s, flags=re.IGNORECASE)   # footnote (a)(b)
    s = re.sub(r"\(\d+\)", " ", s)                           # footnote (1)(2)
    s = re.sub(r"[\n\r|]+", " ", s)                          # newlines, pipe seps
    s = re.sub(r"[^a-z0-9 ]+", " ", s.lower())
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _tokens(text: str) -> set[str]:
    return set(re.findall(r"[a-z0-9]+", _norm(text)))


def _jaccard(a: str, b: str) -> float:
    ta, tb = _tokens(a), _tokens(b)
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / len(ta | tb)


# ═══════════════════════════════════════════════════════════════════════════════
# Signal 1: Alias matching
# ═══════════════════════════════════════════════════════════════════════════════

def _alias_signal(col_name: str, field: dict) -> tuple[float, str]:
    """
    Exact and near-exact match against the field's alias list.

    Returns (confidence, explanation).
    Scores:
      - normalised col == normalised alias → 0.95
      - normalised col ⊆ normalised alias or vice versa → 0.85
    """
    nc = _norm(col_name)
    for alias in field.get("aliases", []):
        na = _norm(alias)
        if nc == na:
            return 0.95, f"exact alias match: '{alias}'"
        if na in nc or nc in na:
            overlap = len(na) / max(len(nc), len(na))
            if overlap >= 0.7:
                conf = 0.80 + 0.10 * overlap
                return conf, f"alias substring match: '{alias}'"
    return 0.0, ""


# ═══════════════════════════════════════════════════════════════════════════════
# Signal 2: Regex pattern matching
# ═══════════════════════════════════════════════════════════════════════════════

# Field-specific regex patterns.  Each entry: (pattern, confidence).
# Patterns are applied to the normalised column name.
_FIELD_PATTERNS: dict[str, list[tuple[str, float]]] = {
    "entry_date": [
        (r"\b(entry|investment|inv|acquisition|acq|initial)\s+(date|year|dt)\b", 0.90),
        (r"\bdate\s+of\s+(inv|invest|acquisition|entry)\b", 0.90),
        (r"\binvestition\b", 0.80),   # German "Investitionsdatum"
    ],
    "exit_date": [
        (r"\b(exit|disposal|divestment|reali[sz]ation|sale|divest)\s+(date|year|dt)\b", 0.90),
        (r"\bexit\s*/?\s*valuation\s+date\b", 0.85),
        (r"\bvaluation\s+date\b", 0.70),   # weaker — could also be exit
    ],
    "company": [
        (r"\b(portfolio\s+company|portco|target\s+company|investment\s+name)\b", 0.92),
        (r"\bcompany\s*(name)?\b", 0.88),
        (r"\b(asset|holding)\s*(name)?\b", 0.72),
        (r"\bposition\b", 0.65),           # ILPA template uses "Position"
    ],
    "fund": [
        (r"\bfund\s*(name|vehicle)?\b", 0.90),
        (r"\bvehicle\b", 0.75),
    ],
    "sector": [
        (r"\b(sector|industry|gics|classification|vertical)\b", 0.88),
        (r"\bbusiness\s+segment\b", 0.72),
    ],
    "region": [
        (r"\b(geography|geographic|region|country|domicile|hq\s+location)\b", 0.88),
        (r"\bheadquarters?\b", 0.82),
        (r"\blocation\b", 0.65),
    ],
    "status": [
        (r"\b(deal|investment|reali[sz]ation|position)\s+status\b", 0.92),
        (r"\bstatus\b", 0.75),
        (r"\b(asset|position)\s+type\b", 0.60),   # ILPA "Position Status"
    ],
    "gross_moic": [
        (r"\b(gross|blended|total)?\s*m[ou]ic\b", 0.93),
        (r"\b(gross|blended)?\s*mo[nm]\b", 0.88),   # MOM / MON
        (r"\b(gross|blended|total)?\s*multiple\s+(on|of)\s+(invested|money|cost)\b", 0.90),
        (r"\bmultiple\s+of\s+(cost|money|capital)\b", 0.88),
        (r"\b(money|return|cost|investment)\s+multiple\b", 0.85),
        (r"\b(total|investment)\s+multiple\b", 0.82),
        (r"\bgross\s+(mv|value)\s*/\s*ic\b", 0.85),
    ],
    "gross_irr": [
        (r"\b(gross|project)\s+irr\b", 0.93),
        (r"\bgross\s+(return|yield)\b", 0.78),
    ],
    "net_irr": [
        (r"\bnet\s+irr\b", 0.93),
        (r"\bnet\s+(return|yield)\b", 0.78),
    ],
    "net_moic": [
        (r"\bnet\s+m[ou]ic\b", 0.93),
        (r"\bnet\s+multiple\b", 0.85),
        (r"\btvpi\b", 0.88),   # TVPI ≈ net MOIC in limited partnership context
    ],
    "ic_total": [
        (r"\b(total|fund)\s+(equity\s+)?(invested|investment)\s+(capital|equity|cost)\b", 0.90),
        (r"\binvested\s+capital\b", 0.90),
        (r"\b(total|current|final)\s+(investment|fund)\s+cost\b(?!\s+saving)", 0.88),
        (r"\b(acquisition|purchase|entry)\s+(cost|price)\b", 0.80),
        (r"\binvestment\s+cost\b(?!\s+saving)", 0.85),
        (r"\bic\b", 0.70),
    ],
    "ic_initial": [
        (r"\b(initial|original)\s+(equity\s+)?(invested|investment|capital)\b", 0.90),
        (r"\binitial\s+fund\s+equity\b", 0.92),
    ],
    "realized": [
        (r"\breali[sz]ed\s+(value|proceeds|return|equity)\b", 0.92),
        (r"\b(exit|sale|disposal)\s+proceeds\b", 0.88),
        (r"\bcumulative\s+reali[sz]ation\b", 0.85),
    ],
    "unrealized": [
        (r"\bunreali[sz]ed\s+(value|nav|equity|market\s+value)\b", 0.92),
        (r"\b(current|fair|residual)\s+(value|nav|equity)\b", 0.80),
        (r"\bremaining\s+(value|equity)\b", 0.80),
    ],
    "total_value": [
        (r"\btotal\s+(fund\s+)?(value|return|proceeds)\b", 0.88),
        (r"\bgross\s+(value|return)\b", 0.80),
    ],
    "holding_period": [
        (r"\b(holding|hold)\s+period\b", 0.92),
        (r"\b(years?\s+held|duration|years?\s+owned)\b", 0.85),
    ],
    "entry_rev": [
        (r"\b(entry|initial|acq|acquisition|at\s+entry)\s+(ltm\s+)?(revenue|sales|turnover)\b", 0.90),
        (r"\bltm\s+(revenue|sales)\s+at\s+(entry|acquisition|investment)\b", 0.90),
        (r"\b(revenue|sales)\s+at\s+(entry|acquisition)\b", 0.85),
        # Order-agnostic ("Revenue | Acquisition"); excludes growth/margin/multiple columns.
        (r"(?=.*\b(revenue|sales|turnover)\b)(?=.*\b(acq|acquisition|entry|initial|investment)\b)(?!.*\b(growth|change|cagr|margin|multiple|per\s+share)\b)", 0.88),
    ],
    "exit_rev": [
        (r"\b(exit|current|latest|final)\s+(ltm\s+)?(revenue|sales|turnover)\b", 0.90),
        (r"\bltm\s+(revenue|sales)\s+at\s+(exit|current)\b", 0.90),
        (r"(?=.*\b(revenue|sales|turnover)\b)(?=.*\b(exit|current|latest|final)\b)(?!.*\b(growth|change|cagr|margin|multiple|per\s+share)\b)", 0.88),
    ],
    "entry_ebitda": [
        (r"\b(entry|initial|acq|acquisition|at\s+entry)\s+(ltm\s+)?ebitda\b", 0.92),
        (r"\bltm\s+ebitda\s+at\s+(entry|acquisition|investment)\b", 0.92),
    ],
    "exit_ebitda": [
        (r"\b(exit|current|latest|final)\s+(ltm\s+)?ebitda\b", 0.92),
        (r"\bltm\s+ebitda\s+at\s+(exit|current)\b", 0.92),
    ],
    "entry_ev": [
        (r"\b(entry|initial|acq|acquisition|at\s+entry)\s+(enterprise\s+value|ev|tev)\b", 0.92),
        (r"\b(enterprise\s+value|ev|tev)\s+at\s+(entry|acquisition|investment)\b", 0.90),
        # Order-agnostic ("Total Enterprise Valuation | Acquisition").
        (r"(?=.*\b(enterprise\s+valuation|enterprise\s+value|tev|ev)\b)(?=.*\b(acq|acquisition|entry|initial|investment)\b)(?!.*\b(growth|change|multiple|per\s+share|equity)\b)", 0.88),
    ],
    "exit_ev": [
        (r"\b(exit|current|latest|final)\s+(enterprise\s+value|ev|tev)\b", 0.92),
        (r"\b(enterprise\s+value|ev|tev)\s+at\s+(exit|current)\b", 0.90),
        (r"\bcurrent\s+(enterprise\s+value|ev|tev)\b", 0.88),
        (r"(?=.*\b(enterprise\s+valuation|enterprise\s+value|tev|ev)\b)(?=.*\b(exit|current|latest|final)\b)(?!.*\b(growth|change|multiple|per\s+share|equity)\b)", 0.88),
    ],
    "entry_net_debt": [
        (r"\b(entry|initial|acq|acquisition|at\s+entry)\s+net\s+debt\b", 0.92),
        (r"\bnet\s+debt\s+at\s+(entry|acquisition|investment)\b", 0.90),
    ],
    "exit_net_debt": [
        (r"\b(exit|current|latest|final)\s+net\s+debt\b", 0.92),
        (r"\bnet\s+debt\s+at\s+(exit|current|valuation)\b", 0.90),
    ],
    "valuation_method": [
        (r"\b(valuation|unreali[sz]ed\s+valuation)\s+method(ology)?\b", 0.92),
        (r"\bvaluation\s+approach\b", 0.85),
    ],
    "fund_ownership": [
        (r"\b(fund\s+)?ownership\s+(%|pct|percent|stake)\b", 0.90),
        (r"\bfully[- ]diluted\s+ownership\b", 0.88),
        (r"\b(equity\s+)?ownership\b", 0.75),
    ],
    "transaction_type": [
        (r"\b(transaction|deal|investment)\s+type\b", 0.88),
        (r"\bstructure\s+type\b", 0.75),
    ],
    "competition": [
        (r"\b(process|auction)\s+type\b", 0.88),
        (r"\bcompetition\b", 0.85),
        (r"\bsourcing\s+process\b", 0.82),
        # "Deal Source" / "Source" describes how the deal was sourced
        # (Proprietary / Auction / …) — a process-type signal, not a deal type.
        (r"\b(deal\s+)?source\b", 0.80),
    ],
    "role": [
        (r"\b(gp|fund|investment)\s+role\b", 0.90),
        (r"\b(lead|co-?invest)\b", 0.65),
    ],
    "exit_type": [
        (r"\b(exit|divestment)\s+type\b", 0.90),
        (r"\b(exit|divestment)\s+route\b", 0.88),
        (r"\btype\s+of\s+exit\b", 0.90),
    ],
    "sourcing_partner": [
        (r"\b(sourcing|advisory|responsible)\s+partner\b", 0.90),
        (r"\bpartner\b", 0.55),
    ],
}

# Pre-compile
_COMPILED_PATTERNS: dict[str, list[tuple[re.Pattern, float]]] = {
    fid: [(re.compile(pat, re.IGNORECASE), conf) for pat, conf in pats]
    for fid, pats in _FIELD_PATTERNS.items()
}


def _regex_signal(col_name: str, field_id: str) -> tuple[float, str]:
    """Match normalised col_name against field-specific regex patterns."""
    nc = _norm(col_name)
    for pattern, conf in _COMPILED_PATTERNS.get(field_id, []):
        if pattern.search(nc):
            return conf, f"regex: {pattern.pattern!r}"
    return 0.0, ""


# ═══════════════════════════════════════════════════════════════════════════════
# Signal 3: Fuzzy token overlap
# ═══════════════════════════════════════════════════════════════════════════════

def _fuzzy_signal(col_name: str, field: dict) -> tuple[float, str]:
    """
    Jaccard token similarity against the field label and each alias.
    Only contributes when similarity is meaningful (≥ 0.40).
    """
    best_score = 0.0
    best_alias = ""
    targets = [field.get("label", "")] + field.get("aliases", [])
    nc = col_name
    for target in targets:
        j = _jaccard(nc, target)
        if j > best_score:
            best_score = j
            best_alias = target
    if best_score < 0.40:
        return 0.0, ""
    conf = 0.50 + (best_score - 0.40) * 3.5   # 0.50 at 0.40 → 0.85 at 0.60
    return min(conf, 0.85), f"fuzzy token match: '{best_alias}' (Jaccard={best_score:.2f})"


# ═══════════════════════════════════════════════════════════════════════════════
# Signal 4: Column profile heuristics (additive boost)
# ═══════════════════════════════════════════════════════════════════════════════

# Which profile attributes boost which fields, and by how much
_PROFILE_BOOSTS: dict[str, list[tuple[str, float]]] = {
    "gross_moic":        [("looks_like_moic", 0.20)],
    "net_moic":          [("looks_like_moic", 0.15)],
    "gross_irr":         [("looks_like_irr",  0.20)],
    "net_irr":           [("looks_like_irr",  0.15)],
    "entry_date":        [("looks_like_date",  0.20)],
    "exit_date":         [("looks_like_date",  0.20)],
    "ic_total":      [("looks_like_currency", 0.10)],
    "ic_initial":    [("looks_like_currency", 0.10)],
    "realized":      [("looks_like_currency", 0.10)],
    "unrealized":    [("looks_like_currency", 0.10)],
    "total_value":   [("looks_like_currency", 0.10)],
    "company":           [("looks_like_identifier", 0.10)],
    "fund":              [("looks_like_identifier", 0.08)],
    "region":            [("looks_like_identifier", 0.05)],
    "sector":            [("looks_like_identifier", 0.05)],
}


def _profile_signal(col_prof: ColumnProfile, field_id: str) -> tuple[float, str]:
    """
    Additive confidence boost from column distribution characteristics.
    This is a supporting signal only — it never drives a mapping on its own.
    Returns (boost, explanation).
    """
    total = 0.0
    notes = []
    for attr, boost in _PROFILE_BOOSTS.get(field_id, []):
        if getattr(col_prof, attr, False):
            total += boost
            notes.append(attr.replace("looks_like_", ""))
    if total == 0:
        return 0.0, ""
    return total, f"profile signal: {'+'.join(notes)} (+{total:.2f})"


# ═══════════════════════════════════════════════════════════════════════════════
# Signal 5: Semantic embedding similarity (fallback)
# ═══════════════════════════════════════════════════════════════════════════════

_sem_model = None   # lazy-loaded


def _get_sem_model():
    global _sem_model
    if _sem_model is None:
        from sentence_transformers import SentenceTransformer
        _sem_model = SentenceTransformer("paraphrase-multilingual-MiniLM-L12-v2")
    return _sem_model


def _semantic_signals_bulk(
    col_names: list[str],
) -> dict[str, dict[str, tuple[float, str]]]:
    """
    For each (col_name, field_id) pair, return the semantic cosine similarity.

    Only called for columns that scored < CONFIDENCE_REVIEW from deterministic
    signals, to avoid expensive model inference when not needed.

    Returns {col_name: {field_id: (score, explanation)}}.
    """
    from sklearn.metrics.pairwise import cosine_similarity

    model = _get_sem_model()
    field_labels = [
        f["label"] + " " + " ".join(f["aliases"])
        for f in TEMPLATE_FIELDS
    ]
    field_vecs = model.encode(field_labels, normalize_embeddings=True)
    col_vecs   = model.encode(col_names,    normalize_embeddings=True)
    sims = cosine_similarity(col_vecs, field_vecs)

    result: dict[str, dict[str, tuple[float, str]]] = {}
    for i, col in enumerate(col_names):
        result[col] = {}
        for j, field in enumerate(TEMPLATE_FIELDS):
            score = float(sims[i, j])
            if score >= 0.45:
                result[col][field["id"]] = (score * 0.85, f"semantic similarity={score:.2f}")
    return result




# ═══════════════════════════════════════════════════════════════════════════════
# Company value plausibility
# ═══════════════════════════════════════════════════════════════════════════════
# Header text alone mis-identifies company columns ("Company Website" holds
# URLs; the real names may sit under a junk header like "Currency in $M").
# These checks look at the VALUES.

_URL_RE = re.compile(r"https?://|www\.|\.(com|net|org|io|co)\b", re.IGNORECASE)


def _str_samples(cp: "ColumnProfile") -> list[str]:
    return [str(v).strip() for v in cp.sample_values
            if v is not None and str(v).strip()]


def _company_value_multiplier(cp: "ColumnProfile", row_count: int) -> tuple[float, str]:
    """Multiplier (0–1] applied to name-based company scores, with a reason."""
    if cp.data_type in ("numeric", "percent", "date"):
        return 0.1, "values are numeric/dates, not names"
    if cp.fill_rate < 0.3:
        return 0.3, f"mostly empty ({cp.fill_rate:.0%} filled)"
    samples = _str_samples(cp)
    if not samples:
        return 0.3, "no text values"
    url_frac = sum(1 for s in samples if _URL_RE.search(s)) / len(samples)
    if url_frac > 0.3:
        return 0.15, "values look like URLs"
    avg_len = sum(len(s) for s in samples) / len(samples)
    avg_words = sum(len(s.split()) for s in samples) / len(samples)
    if avg_len > 60 or avg_words > 7:
        return 0.3, "values look like descriptions, not names"
    filled = max(1.0, row_count * cp.fill_rate)
    if cp.unique_count / filled < 0.35:
        return 0.4, "values repeat too much for company names"
    return 1.0, ""


def _company_likeness(cp: "ColumnProfile", row_count: int) -> float:
    """0–1 score: do this column's VALUES look like company names?"""
    if cp.data_type in ("numeric", "percent", "date", "empty"):
        return 0.0
    samples = _str_samples(cp)
    if not samples or cp.fill_rate < 0.5:
        return 0.0
    if sum(1 for s in samples if _URL_RE.search(s)) / len(samples) > 0.2:
        return 0.0
    avg_len = sum(len(s) for s in samples) / len(samples)
    avg_words = sum(len(s.split()) for s in samples) / len(samples)
    if not (2 <= avg_len <= 45 and avg_words <= 6):
        return 0.0
    filled = max(1.0, row_count * cp.fill_rate)
    unique_ratio = min(1.0, cp.unique_count / filled)
    if unique_ratio < 0.5:
        return 0.0
    score = 0.40 + 0.25 * unique_ratio + 0.15 * cp.fill_rate
    if cp.col_index <= 2:          # company is usually among the first columns
        score += 0.08
    return min(score, 0.9)


def _rescue_company(
    resolved: dict, scores: dict, col_profiles: dict, row_count: int,
) -> None:
    """When no header-based company mapping is trustworthy, find the column
    whose VALUES most look like company names (handles junk headers)."""
    col, conf, _sigs = resolved.get("company", (None, 0.0, []))
    if col is not None and conf >= CONFIDENCE_REVIEW:
        return
    taken = {c for fid, (c, cf, _s) in resolved.items()
             if c is not None and fid != "company"}
    best_col, best_score = None, 0.0
    for name, cp in col_profiles.items():
        if name in taken:
            continue
        s = _company_likeness(cp, row_count)
        if s > best_score:
            best_col, best_score = name, s
    if best_col and best_score >= 0.55:
        resolved["company"] = (
            best_col, min(0.65, best_score),
            [f"value-based company detection (header '{best_col}' gave no name signal)"])


# ═══════════════════════════════════════════════════════════════════════════════
# Conflict resolution
# ═══════════════════════════════════════════════════════════════════════════════

def _resolve_conflicts(
    scores: dict[str, dict[str, tuple[float, list[str]]]],
) -> dict[str, tuple[Optional[str], float, list[str]]]:
    """
    Given scores[field_id][col_name] = (confidence, signals),
    resolve conflicts so that each field maps to at most one column and
    each column is used at most once.

    Strategy: greedy highest-confidence assignment, then re-resolve ties.
    Returns {field_id: (col_name_or_None, confidence, signals)}.
    """
    # Flatten: list of (confidence, field_id, col_name, signals)
    entries = []
    for fid, cols in scores.items():
        for col, (conf, sigs) in cols.items():
            if conf >= CONFIDENCE_REVIEW:
                entries.append((conf, fid, col, sigs))

    entries.sort(key=lambda x: x[0], reverse=True)

    assigned_cols: set[str]    = set()
    assigned_fields: set[str]  = set()
    result: dict[str, tuple[Optional[str], float, list[str]]] = {}

    for conf, fid, col, sigs in entries:
        if fid in assigned_fields or col in assigned_cols:
            continue
        result[fid] = (col, conf, sigs)
        assigned_fields.add(fid)
        assigned_cols.add(col)

    # Fields with no confident mapping
    for field in TEMPLATE_FIELDS:
        if field["id"] not in result:
            result[field["id"]] = (None, 0.0, [])

    return result


# ═══════════════════════════════════════════════════════════════════════════════
# Public API
# ═══════════════════════════════════════════════════════════════════════════════

def infer_schema(
    table: ExtractedTable,
    llm: Optional["LLMInterface"] = None,
    use_semantic: bool = True,
) -> SchemaInference:
    """
    Infer which DataFrame column maps to which standardised field.

    Parameters
    ----------
    table        : ExtractedTable   Stage 2 output.
    llm          : LLMInterface     Optional LLM fallback for ambiguous columns.
    use_semantic : bool             Enable sentence-transformer fallback.

    Returns
    -------
    SchemaInference  Complete field mapping with confidence scores and
                     human-readable signal explanations.
    """
    col_profiles = {cp.name: cp for cp in table.col_profiles}
    col_names    = list(col_profiles.keys())

    # ── Compute deterministic scores (Signals 1–4) ────────────────────
    # scores[field_id][col_name] = (confidence, [signal_strings])
    scores: dict[str, dict[str, tuple[float, list[str]]]] = {
        f["id"]: {} for f in TEMPLATE_FIELDS
    }

    for field in TEMPLATE_FIELDS:
        fid = field["id"]
        for col in col_names:
            cp   = col_profiles[col]
            sigs = []

            # Signal 1: alias
            a_conf, a_sig = _alias_signal(col, field)
            # Signal 2: regex
            r_conf, r_sig = _regex_signal(col, fid)
            # Signal 3: fuzzy
            f_conf, f_sig = _fuzzy_signal(col, field)
            # Signal 4: profile boost (additive, capped)
            p_conf, p_sig = _profile_signal(cp, fid)

            # Primary confidence = max of name-based signals
            primary = max(a_conf, r_conf, f_conf)
            if a_conf > 0 and a_sig: sigs.append(a_sig)
            if r_conf > 0 and r_sig: sigs.append(r_sig)
            if f_conf > 0 and f_sig: sigs.append(f_sig)

            # Profile boost: only when there is already some primary signal
            # (prevents pure-distribution guesses)
            boost = 0.0
            if primary >= 0.30 and p_conf > 0:
                boost = p_conf
                sigs.append(p_sig)

            total = min(1.0, primary + boost)

            if total > 0:
                scores[fid][col] = (total, sigs)

    # ── Semantic fallback (Signal 5) ──────────────────────────────────
    # Only run for columns that didn't reach CONFIDENCE_REVIEW yet
    if use_semantic:
        low_conf_cols = [
            col for col in col_names
            if max(
                (scores[fid].get(col, (0.0, []))[0] for fid in scores),
                default=0.0,
            ) < CONFIDENCE_REVIEW
        ]

        if low_conf_cols:
            try:
                sem_results = _semantic_signals_bulk(low_conf_cols)
                for col, field_scores in sem_results.items():
                    for fid, (sem_conf, sem_sig) in field_scores.items():
                        existing_conf = scores[fid].get(col, (0.0, []))[0]
                        if sem_conf > existing_conf:
                            existing_sigs = scores[fid].get(col, (0.0, []))[1]
                            scores[fid][col] = (sem_conf, existing_sigs + [sem_sig])
            except Exception:
                pass   # semantic model unavailable — deterministic signals only

    # ── LLM fallback (Signal 6) ───────────────────────────────────────
    if llm is not None:
        still_low = [
            col for col in col_names
            if max(
                (scores[fid].get(col, (0.0, []))[0] for fid in scores),
                default=0.0,
            ) < CONFIDENCE_REVIEW
        ]

        if still_low:
            try:
                already_mapped = {
                    fid: scores[fid].get(col, (0.0, []))[0]
                    for fid in scores
                    for col in [max(scores[fid], key=lambda c: scores[fid][c][0], default=None)]
                    if col
                }
                candidate_fields = [
                    {"id": f["id"], "label": f["label"],
                     "aliases": f["aliases"], "description": ""}
                    for f in TEMPLATE_FIELDS
                ]
                ambiguous = [
                    {
                        "name": col,
                        "sample_values": col_profiles[col].sample_values[:5],
                        "data_type": col_profiles[col].data_type,
                        "neighboring_cols": _neighboring_cols(col, col_names, n=3),
                    }
                    for col in still_low
                ]
                llm_results = llm.infer_fields(ambiguous, candidate_fields,
                                               context={"already_mapped": already_mapped})
                for inf in llm_results:
                    if inf.source_col and inf.confidence >= CONFIDENCE_REVIEW:
                        existing = scores[inf.field_id].get(inf.source_col, (0.0, []))
                        if inf.confidence > existing[0]:
                            scores[inf.field_id][inf.source_col] = (
                                inf.confidence,
                                existing[1] + [f"llm: {'; '.join(inf.signals)}"],
                            )
            except Exception:
                pass   # LLM unavailable or returned unusable output

    # ── Company value plausibility (penalise URL/description/code columns) ─
    for col in list(scores.get("company", {})):
        mult, why = _company_value_multiplier(col_profiles[col], table.row_count)
        if mult < 1.0:
            conf0, sigs0 = scores["company"][col]
            scores["company"][col] = (conf0 * mult, sigs0 + [f"penalised: {why}"])

    # ── Resolve conflicts ─────────────────────────────────────────────
    resolved = _resolve_conflicts(scores)

    # ── Company rescue: value-driven detection for junk headers ───────
    _rescue_company(resolved, scores, col_profiles, table.row_count)

    # ── Build SchemaInference ─────────────────────────────────────────
    field_to_col:  dict[str, Optional[str]]   = {}
    col_to_field:  dict[str, Optional[str]]   = {c: None for c in col_names}
    confidences:   dict[str, float]           = {}
    explanations:  dict[str, list[str]]       = {}
    auto_conf:     list[str] = []
    needs_rev:     list[str] = []
    unmapped_flds: list[str] = []

    for field in TEMPLATE_FIELDS:
        fid = field["id"]
        col, conf, sigs = resolved[fid]
        field_to_col[fid] = col
        confidences[fid]  = conf
        explanations[fid] = sigs

        if col is not None:
            col_to_field[col] = fid

        if conf >= CONFIDENCE_AUTO:
            auto_conf.append(fid)
        elif conf >= CONFIDENCE_REVIEW:
            needs_rev.append(fid)
        else:
            unmapped_flds.append(fid)

    unmapped_cols = [c for c in col_names if col_to_field.get(c) is None]

    return SchemaInference(
        field_to_col=field_to_col,
        col_to_field=col_to_field,
        confidences=confidences,
        explanations=explanations,
        auto_confirmed=auto_conf,
        needs_review=needs_rev,
        unmapped_fields=unmapped_flds,
        unmapped_cols=unmapped_cols,
    )


def _neighboring_cols(col: str, all_cols: list[str], n: int = 3) -> list[str]:
    """Return up to n column names adjacent to col in the list."""
    try:
        idx = all_cols.index(col)
    except ValueError:
        return []
    start = max(0, idx - n)
    end   = min(len(all_cols), idx + n + 1)
    return [c for c in all_cols[start:end] if c != col]
