"""
extractor.py — Stage 2: Table Extraction

Takes a WorkbookProfile (Stage 1) and the raw file bytes, then extracts
the primary (or user-specified) table into an ExtractedTable.

Responsibilities:
  - Open the workbook at the profiler-identified table location
  - Build composite column names from multi-row headers
  - Profile each column: data type, fill rate, value distribution, units
  - Handle merged cells by propagating their values

The output is the canonical "raw data layer" consumed by Stage 3 (inference).
No field mapping or standardisation happens here.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any, Optional

import pandas as pd

from profiler import WorkbookProfile, TableCandidate
from parser import parse_gp_file


# ═══════════════════════════════════════════════════════════════════════════════
# Data Structures
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class ColumnProfile:
    """
    Structural and statistical description of one extracted column.

    Used by Stage 3 (inferencer) to combine distribution-based evidence
    with name-based signals.
    """
    name: str                           # raw composite header (may contain " | ")
    col_index: int                      # 0-based position in the DataFrame

    # Type characterisation
    data_type: str                      # "date" | "numeric" | "percent" | "string" | "mixed" | "empty"
    fill_rate: float                    # fraction of non-empty rows

    # Value samples
    sample_values: list[Any]            # first ≤10 non-empty values

    # Numeric statistics (None when data_type is "string" or "empty")
    unit: Optional[str]                 # detected unit: "EUR m", "USD m", "%", "x", etc.
    value_min: Optional[float]
    value_max: Optional[float]
    value_mean: Optional[float]

    # Distribution heuristics — used as supplementary inference signals
    unique_count: int
    looks_like_moic: bool               # values cluster in 0.5–25 × range
    looks_like_irr: bool                # values look like IRR fractions or pcts
    looks_like_date: bool               # majority of values are date/datetime
    looks_like_currency: bool           # large positive integers/floats (≥ 0.1)
    looks_like_identifier: bool         # short strings, high uniqueness


@dataclass
class ExtractedTable:
    """
    Output of Stage 2.

    A clean, profile-annotated representation of the detected deal table.
    The DataFrame uses composite headers when the source has multi-row headers.
    No field mapping has been applied yet.
    """
    source_file: str
    sheet_name: str
    table_candidate: TableCandidate     # from profiler
    df: pd.DataFrame                    # raw data, columns = composite headers
    col_profiles: list[ColumnProfile]   # one per DataFrame column
    header_rows: list[int]              # 1-based
    first_data_row: int                 # 1-based
    row_count: int
    col_count: int
    warnings: list[str]

    def column_by_name(self, name: str) -> Optional[ColumnProfile]:
        """Return the ColumnProfile for the given column name, or None."""
        for cp in self.col_profiles:
            if cp.name == name:
                return cp
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# Unit detection
# ═══════════════════════════════════════════════════════════════════════════════

_UNIT_PATTERNS: list[tuple[str, str]] = [
    (r"\(EUR\s*[mM]\b", "EUR m"),
    (r"\(USD\s*[mM]\b", "USD m"),
    (r"\(\$\s*[mM]\b",  "USD m"),
    (r"\(£\s*[mM]\b",   "GBP m"),
    (r"\(€\s*[mM]\b",   "EUR m"),
    (r"\([Mm]illions?\)", "m"),
    (r"\([Mm]\)",         "m"),
    (r"\b[Mm]ln\b",       "m"),
    (r"\([Kk]000s?\)",    "k"),
    (r"\([Kk]\)",         "k"),
    (r"\bthousands?\b",   "k"),
    (r"['’]000s?\b", "k"),
    (r"\(\s*000s?\s*\)",  "k"),
    (r"\bin\s+000s?\b",   "k"),
    (r"\b(?:USD|EUR|GBP)\s*[Kk]\b", "k"),
    (r"\(\$\s*[Kk]\b",    "k"),
    (r"\(%\)|\bpct\b",    "%"),
    (r"\(x\)|\bx\b",      "x"),
]


def _detect_unit(col_name: str) -> Optional[str]:
    for pat, unit in _UNIT_PATTERNS:
        if re.search(pat, col_name, re.IGNORECASE):
            return unit
    return None


# ═══════════════════════════════════════════════════════════════════════════════
# Column profiling
# ═══════════════════════════════════════════════════════════════════════════════

def _profile_column(series: pd.Series, name: str, col_idx: int) -> ColumnProfile:
    """Compute ColumnProfile for one DataFrame column."""

    # Clean: drop null and blank/dash placeholders
    clean = series.dropna()
    clean = clean[~clean.astype(str).str.strip().isin(("", "-", "–", "—", "N/A", "n/a", "#N/A", "#VALUE!"))]

    n_total = max(len(series), 1)
    n_clean = len(clean)
    fill_rate = n_clean / n_total

    if n_clean == 0:
        return ColumnProfile(
            name=name, col_index=col_idx,
            data_type="empty", fill_rate=0.0,
            sample_values=[], unit=_detect_unit(name),
            value_min=None, value_max=None, value_mean=None,
            unique_count=0,
            looks_like_moic=False, looks_like_irr=False,
            looks_like_date=False, looks_like_currency=False,
            looks_like_identifier=False,
        )

    # ── Date detection ─────────────────────────────────────────────────
    date_count = sum(
        1 for v in clean
        if isinstance(v, (date, datetime))
        or (isinstance(v, str) and re.match(r"\d{4}-\d{2}-\d{2}", v.strip()))
    )
    date_fraction = date_count / n_clean

    # ── Numeric detection ──────────────────────────────────────────────
    nums = pd.to_numeric(clean, errors="coerce").dropna()
    n_numeric = len(nums)
    num_fraction = n_numeric / n_clean

    # ── String detection ───────────────────────────────────────────────
    # "string" means non-date, non-numeric textual content
    n_string = sum(1 for v in clean if isinstance(v, str) and not re.match(r"^-?\d", str(v)))

    # ── Determine primary type ─────────────────────────────────────────
    if date_fraction >= 0.5:
        data_type = "date"
    elif num_fraction >= 0.75:
        # Distinguish plain numeric from percentage
        pct_like = (
            (nums.abs() <= 2.0).sum() / len(nums) >= 0.8
            or clean.astype(str).str.contains("%").mean() >= 0.3
        )
        data_type = "percent" if pct_like else "numeric"
    elif num_fraction >= 0.3:
        data_type = "mixed"
    else:
        data_type = "string"

    # ── Numeric statistics ─────────────────────────────────────────────
    if n_numeric >= 3:
        v_min   = float(nums.min())
        v_max   = float(nums.max())
        v_mean  = float(nums.mean())
    else:
        v_min = v_max = v_mean = None

    # ── Distribution heuristics ────────────────────────────────────────
    looks_moic = False
    looks_irr  = False
    looks_ccy  = False

    if n_numeric >= 3:
        # MOIC: positive, most values 0.5–25 ×, mean 1–8
        in_moic_range = ((nums > 0) & (nums < 30)).mean()
        in_moic_sweet = ((nums >= 0.5) & (nums <= 15)).mean()
        has_x_suffix  = clean.astype(str).str.rstrip().str.endswith("x").mean()
        looks_moic = bool(in_moic_range >= 0.85 and in_moic_sweet >= 0.6 and v_mean is not None and 0.5 < v_mean < 20)  # type: ignore[operator]

        # IRR: raw fractions (-0.9 to 3.0) OR percentages (-90 to 300 if /100)
        frac_in_irr = ((nums > -0.90) & (nums < 3.0)).mean()
        pct_in_irr  = ((nums > -90) & (nums < 300)).mean()
        looks_irr   = bool(
            (frac_in_irr >= 0.9 and data_type in ("numeric", "mixed") and (v_mean or 0) < 0.5)
            or (pct_in_irr >= 0.9 and data_type == "percent")
        )

        # Currency: large positive numbers (EUR/USD millions)
        large_pos = (nums >= 0.1).mean()
        looks_ccy = bool(large_pos >= 0.8 and (v_max or 0) >= 1.0)

    # Identifier: short unique strings (e.g., company names, fund names)
    avg_len = clean.astype(str).str.len().mean() if n_clean > 0 else 0
    looks_id = bool(
        data_type == "string"
        and fill_rate >= 0.5
        and clean.nunique() >= min(3, n_clean * 0.3)
        and avg_len < 60
    )

    return ColumnProfile(
        name=name, col_index=col_idx,
        data_type=data_type, fill_rate=fill_rate,
        sample_values=clean.head(10).tolist(),
        unit=_detect_unit(name),
        value_min=v_min, value_max=v_max, value_mean=v_mean,
        unique_count=int(clean.nunique()),
        looks_like_moic=looks_moic,
        looks_like_irr=looks_irr,
        looks_like_date=date_fraction >= 0.5,
        looks_like_currency=looks_ccy,
        looks_like_identifier=looks_id,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# Public API
# ═══════════════════════════════════════════════════════════════════════════════

# Column injected into the combined frame when concatenating per-fund tabs.
# Double-underscore marks it internal so the UI can filter it from "unmapped".
FUND_FROM_TAB_COL = "__fund_tab__"

# Separators used between a fund label and a view descriptor in a tab name,
# e.g. "Fund XI - Acquisition vs Exit", "SGF I | Summary".
_TAB_SEPARATOR = re.compile(r"\s*[-–—|:]\s*")


def _split_tab_segments(name: str) -> list[str]:
    return [s for s in _TAB_SEPARATOR.split(str(name).strip()) if s]


def _derive_fund_labels(sheet_names: list[str]) -> dict[str, str]:
    """
    Recover a clean fund label for each per-fund tab by stripping the *common
    trailing view descriptor* shared across the group — no hardcoded vocabulary.

    ["Fund XI - Deal Data", "Fund XII - Deal Data"] → {..: "Fund XI", ..: "Fund XII"}
    ["Fund XI", "Fund XII"]                          → unchanged (no common suffix)

    This generalises across GP naming schemes: whatever view label the tabs have
    in common (e.g. "Acquisition vs Exit", "Summary", "Deal Data") is removed,
    leaving the part that actually varies — the fund.
    """
    seg_lists = {n: _split_tab_segments(n) for n in sheet_names}

    # Longest common trailing run of segments across all tabs.
    common = 0
    if len(sheet_names) >= 2:
        shortest = min(len(v) for v in seg_lists.values())
        for k in range(1, shortest):          # keep ≥1 segment as the fund
            tail_sets = {tuple(v[-k:]) for v in seg_lists.values()}
            if len(tail_sets) == 1:
                common = k
            else:
                break

    labels: dict[str, str] = {}
    for name, segs in seg_lists.items():
        kept = segs[:-common] if common and len(segs) > common else segs
        labels[name] = " - ".join(kept).strip() or str(name).strip()
    return labels


# Trailing footnote markers on a header: "(5)", "*", "†", superscript digits.
# Only pure-digit parentheticals are footnotes — "(mm/dd/yyyy)", "(mlns)" and
# other lettered parentheticals are real content and must survive.
_FOOTNOTE_TAIL_RE = re.compile(r"(?:\s*\(\d{1,2}\)|\s*[*†‡¹²³⁴⁵⁶⁷⁸⁹⁰])+\s*$")


def _footnote_key(name: Any) -> str:
    """Case/whitespace-insensitive header identity with footnote tails removed."""
    s = re.sub(r"\s+", " ", str(name)).strip()
    return _FOOTNOTE_TAIL_RE.sub("", s).strip().lower()


def _align_footnote_variant_columns(
    frames: list[pd.DataFrame], warnings: list[str]
) -> None:
    """
    Rename columns in-place so tabs whose headers differ only by a footnote
    marker ("Initial Investment Date (5)" vs "Initial Investment Date") align
    into ONE column when concatenated, instead of fragmenting into several
    sparsely-filled ones.

    The canonical display name for each group is the marker-free variant when
    one exists, else the first variant seen. A key that maps to two different
    columns WITHIN a single tab is left untouched — those are genuinely
    distinct columns and merging them would destroy data.
    """
    variants: dict[str, list[str]] = {}          # key -> distinct raw names, in order
    unsafe: set[str] = set()
    for df_i in frames:
        seen_in_frame: dict[str, str] = {}
        for col in df_i.columns:
            key = _footnote_key(col)
            if key in seen_in_frame and seen_in_frame[key] != str(col):
                unsafe.add(key)                  # two same-key columns in one tab
            seen_in_frame.setdefault(key, str(col))
            names = variants.setdefault(key, [])
            if str(col) not in names:
                names.append(str(col))

    renames: dict[str, str] = {}
    for key, names in variants.items():
        if key in unsafe or len(names) < 2:
            continue
        clean = [n for n in names if _footnote_key(n) == re.sub(r"\s+", " ", n).strip().lower()]
        canonical = clean[0] if clean else names[0]
        for n in names:
            if n != canonical:
                renames[n] = canonical

    if not renames:
        return
    for df_i in frames:
        df_i.rename(columns=renames, inplace=True)
    merged = ", ".join(f"'{a}' → '{b}'" for a, b in sorted(renames.items()))
    warnings.append(f"Aligned footnote-variant headers across tabs: {merged}")


def _extract_one_df(file_bytes: bytes, table: TableCandidate) -> pd.DataFrame:
    """Parse a single TableCandidate into a DataFrame using profiler hints."""
    table_hint = {
        "sheet_name":     table.sheet_name,
        "header_rows":    table.header_rows,
        "first_data_row": table.first_data_row,
    }
    df, _meta = parse_gp_file(file_bytes, table_hint=table_hint)
    return df


def extract_tables(
    file_bytes: bytes,
    profile: WorkbookProfile,
    tables: list[TableCandidate],
) -> ExtractedTable:
    """
    Extract and vertically concatenate several same-schema tables (the per-fund
    layout) into one ExtractedTable.

    Each source tab contributes its rows; a `__fund_tab__` column carries the
    fund identity derived from the tab name, used downstream only when the data
    itself has no fund column. Columns are unioned across tabs so partially
    differing layouts still align. De-duplication happens later in the pipeline,
    once inference has identified the company/fund/entry-date key columns.

    Falls back to single-table behaviour when `tables` has one entry.
    """
    if len(tables) <= 1:
        return extract_table(file_bytes, profile,
                             target_table=tables[0] if tables else None)

    warnings: list[str] = []
    fund_labels = _derive_fund_labels([t.sheet_name for t in tables])
    frames: list[pd.DataFrame] = []
    for t in tables:
        try:
            df_i = _extract_one_df(file_bytes, t)
        except Exception as exc:
            warnings.append(f"Skipped tab '{t.sheet_name}': {exc}")
            continue
        if df_i.empty:
            continue
        df_i = df_i.copy()
        df_i[FUND_FROM_TAB_COL] = fund_labels.get(t.sheet_name, t.sheet_name)
        frames.append(df_i)

    if not frames:
        warnings.append("No rows extracted from any per-fund tab.")
        return ExtractedTable(
            source_file=profile.filename, sheet_name="",
            table_candidate=tables[0], df=pd.DataFrame(), col_profiles=[],
            header_rows=tables[0].header_rows, first_data_row=tables[0].first_data_row,
            row_count=0, col_count=0, warnings=warnings,
        )

    _align_footnote_variant_columns(frames, warnings)
    combined = pd.concat(frames, ignore_index=True, sort=False)

    col_profiles: list[ColumnProfile] = []
    for idx, col in enumerate(combined.columns):
        try:
            col_profiles.append(_profile_column(combined[col], str(col), idx))
        except Exception as exc:
            warnings.append(f"Column profiling failed for '{col}': {exc}")

    sheet_label = " + ".join(t.sheet_name for t in tables)
    warnings.append(
        f"Combined {len(frames)} per-fund tabs into {len(combined)} rows: {sheet_label}"
    )

    return ExtractedTable(
        source_file=profile.filename,
        sheet_name=sheet_label,
        table_candidate=tables[0],
        df=combined,
        col_profiles=col_profiles,
        header_rows=tables[0].header_rows,
        first_data_row=tables[0].first_data_row,
        row_count=len(combined),
        col_count=len(combined.columns),
        warnings=warnings,
    )


def extract_table(
    file_bytes: bytes,
    profile: WorkbookProfile,
    target_table: Optional[TableCandidate] = None,
) -> ExtractedTable:
    """
    Extract a table from the workbook using the profiler's structural knowledge.

    Parameters
    ----------
    file_bytes    : bytes           Raw Excel file bytes.
    profile       : WorkbookProfile Stage 1 output.
    target_table  : TableCandidate  If None, uses profile.primary_table.

    Returns
    -------
    ExtractedTable  Stage 2 output — raw DataFrame + column profiles.
    """
    warnings: list[str] = []
    table = target_table or profile.primary_table

    if table is None:
        # No table detected — return empty result
        warnings.append("No primary table found in workbook. Returning empty extraction.")
        return ExtractedTable(
            source_file=profile.filename, sheet_name="",
            table_candidate=None,  # type: ignore[arg-type]
            df=pd.DataFrame(), col_profiles=[],
            header_rows=[], first_data_row=0,
            row_count=0, col_count=0,
            warnings=warnings,
        )

    # ── Use parse_gp_file with profiler hints for robust extraction ────
    table_hint = {
        "sheet_name":     table.sheet_name,
        "header_rows":    table.header_rows,
        "first_data_row": table.first_data_row,
    }

    try:
        df, meta = parse_gp_file(file_bytes, table_hint=table_hint)
    except Exception as exc:
        warnings.append(f"Extraction error: {exc}")
        df   = pd.DataFrame()
        meta = {}

    # ── Profile every column ───────────────────────────────────────────
    col_profiles: list[ColumnProfile] = []
    for idx, col in enumerate(df.columns):
        try:
            cp = _profile_column(df[col], str(col), idx)
        except Exception as exc:
            warnings.append(f"Column profiling failed for '{col}': {exc}")
            cp = ColumnProfile(
                name=str(col), col_index=idx,
                data_type="mixed", fill_rate=0.0,
                sample_values=[], unit=None,
                value_min=None, value_max=None, value_mean=None,
                unique_count=0,
                looks_like_moic=False, looks_like_irr=False,
                looks_like_date=False, looks_like_currency=False,
                looks_like_identifier=False,
            )
        col_profiles.append(cp)

    return ExtractedTable(
        source_file=profile.filename,
        sheet_name=table.sheet_name,
        table_candidate=table,
        df=df,
        col_profiles=col_profiles,
        header_rows=table.header_rows,
        first_data_row=table.first_data_row,
        row_count=len(df),
        col_count=len(df.columns),
        warnings=warnings,
    )
