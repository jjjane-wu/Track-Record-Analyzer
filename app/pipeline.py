"""
pipeline.py — Full 6-Stage GP Track Record Pipeline

Orchestrates all stages in the correct order and bundles their outputs into
a single PipelineResult.

Usage (default — no LLM)
    from pipeline import GPParserPipeline
    result = GPParserPipeline().run(file_bytes, "fund_data.xlsx")

Usage (with LLM fallback)
    from pipeline import GPParserPipeline
    from llm_interface import ClaudeAPIInterface
    llm    = ClaudeAPIInterface(api_key="sk-...")
    result = GPParserPipeline(llm=llm).run(file_bytes, "fund_data.xlsx")

Stage map
    Stage 1  profiler.profile_workbook     → WorkbookProfile
    Stage 2  extractor.extract_table       → ExtractedTable
    Stage 3  inferencer.infer_schema       → SchemaInference
    Stage 4  (normalisation happens inside writer/transformer — see below)
    Stage 5  validator.validate            → list[ValidationIssue]
    Stage 6  reviewer.generate_report      → ReviewReport

Stage 4 note
    Full data normalisation (date parsing, numeric cleaning, status
    standardisation) is handled by transformer.py during template population.
    The pipeline exposes the raw extracted DataFrame and the mapping; the
    existing transform_row() / populate_template() code consumes both and
    handles normalisation in one pass to avoid double-parsing.  A dedicated
    normaliser.py can be added later as a pre-pass without breaking the
    pipeline contract.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional, TYPE_CHECKING

import pandas as pd

from profiler   import (
    WorkbookProfile, TableCandidate, TableGroup,
    profile_workbook, group_candidate_tables,
)
from extractor  import (
    ExtractedTable, extract_table, extract_tables, FUND_FROM_TAB_COL,
)
from inferencer import SchemaInference, infer_schema
from validator  import ValidationIssue, validate
from reviewer   import ReviewReport, generate_report

if TYPE_CHECKING:
    from llm_interface import LLMInterface


# ═══════════════════════════════════════════════════════════════════════════════
# Result dataclass
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class PipelineResult:
    """
    The complete output of a successful pipeline run.

    All five stages are represented. Callers can use any subset:
      - UI screens only need .report and .profile.
      - Template writers need .table.df and .schema.field_to_col.
      - Validators need .validation_issues.
    """
    # Stage 1
    profile: WorkbookProfile

    # Stage 2
    table: ExtractedTable

    # Stage 3
    schema: SchemaInference

    # Stage 5
    validation_issues: list[ValidationIssue]

    # Stage 6
    report: ReviewReport

    # Pipeline-level metadata
    filename: str
    warnings: list[str] = field(default_factory=list)

    # Layout detection (Stage 1.5) — which table(s) were read and how
    table_group: Optional[TableGroup] = None

    # ── Convenience accessors (backwards-compatible with mapper.py API) ──

    @property
    def df(self) -> pd.DataFrame:
        """Raw extracted DataFrame (original column names)."""
        return self.table.df

    @property
    def field_to_col(self) -> dict[str, Optional[str]]:
        """Mapping from standardised field ID to DataFrame column name."""
        return self.schema.field_to_col

    @property
    def col_confidence(self) -> dict[str, float]:
        """Confidence keyed by DataFrame column name (backwards compat)."""
        return self.schema.col_confidence


# ═══════════════════════════════════════════════════════════════════════════════
# Pipeline
# ═══════════════════════════════════════════════════════════════════════════════

class GPParserPipeline:
    """
    The main entry point for parsing a GP track record Excel file.

    Parameters
    ----------
    llm          : LLMInterface | None
        Optional LLM for Stage 3 fallback.  When None (default), the pipeline
        uses only deterministic signals (alias, regex, fuzzy, distribution).

    use_semantic : bool
        Enable sentence-transformer semantic similarity as Stage 3 fallback.
        Requires the paraphrase-multilingual-MiniLM-L12-v2 model (~90 MB).
        Default: True.  Disable for faster/offline operation.

    target_table : TableCandidate | None
        Override the profiler's auto-detected primary table.  Useful when
        the user selects a specific sheet in the UI.
    """

    def __init__(
        self,
        llm: Optional["LLMInterface"] = None,
        use_semantic: bool = True,
    ):
        self.llm          = llm
        self.use_semantic = use_semantic

    def run(
        self,
        file_bytes: bytes,
        filename: str = "upload.xlsx",
        target_table: Optional[TableCandidate] = None,
        combine_tables: Optional[list[TableCandidate]] = None,
    ) -> PipelineResult:
        """
        Execute all pipeline stages and return a PipelineResult.

        Parameters
        ----------
        file_bytes     : bytes             Raw Excel file bytes.
        filename       : str               Display name (used for error messages).
        target_table   : TableCandidate    Force a single sheet (consolidated override).
        combine_tables : list[TableCandidate]
            Force a specific set of per-fund tabs to concatenate. Overrides
            auto-detection. When None and no target_table is given, the pipeline
            auto-detects consolidated vs per-fund layout.

        Returns
        -------
        PipelineResult
        """
        all_warnings: list[str] = []

        # ── Stage 1: Workbook Profiling ────────────────────────────────────────
        profile = profile_workbook(file_bytes, filename)
        all_warnings.extend(profile.warnings)

        # ── Stage 1.5: Layout decision (consolidated vs per-fund) ──────────────
        # Precedence: explicit combine set > explicit single sheet > auto-detect.
        if combine_tables:
            group = TableGroup(
                tables=combine_tables,
                layout="per_fund" if len(combine_tables) > 1 else "consolidated",
            )
        elif target_table is not None:
            group = TableGroup(tables=[target_table], layout="consolidated")
        else:
            group = group_candidate_tables(profile)

        # ── Stage 2: Table Extraction ──────────────────────────────────────────
        if group and group.layout == "per_fund" and len(group.tables) > 1:
            table = extract_tables(file_bytes, profile, group.tables)
        else:
            single = group.tables[0] if (group and group.tables) else target_table
            table = extract_table(file_bytes, profile, target_table=single)
        all_warnings.extend(table.warnings)

        # ── Stage 3: Schema Inference ──────────────────────────────────────────
        schema = infer_schema(
            table,
            llm=self.llm,
            use_semantic=self.use_semantic,
        )

        # ── Stage 3.4: Fund-section inheritance ────────────────────────────────
        # Some GPs carry no fund column: the deal table is divided into
        # sections, each headed by a row holding the fund's name in the
        # company column and nothing else. Deals inherit the banner above.
        _inherit_fund_sections(table, schema, all_warnings)

        # ── Stage 3.5: Combined-frame cleanup (per-fund only) ──────────────────
        # Supply a fund column from the tab name when the data lacks one, then
        # drop deals that appear in more than one tab (safety against a file that
        # mixes a consolidated sheet with per-fund shards).
        if FUND_FROM_TAB_COL in table.df.columns:
            _apply_fund_fallback_and_dedup(table, schema, all_warnings)

        # ── Stage 3.6: Drop non-deal rows (footnotes / totals / blanks) ────────
        _drop_non_deal_rows(table, schema, all_warnings)

        # ── Stage 5: Validation ────────────────────────────────────────────────
        issues: list[ValidationIssue] = []
        if not table.df.empty:
            try:
                issues = validate(table.df, schema)
            except Exception as exc:
                all_warnings.append(f"Validation skipped: {exc}")

        # ── Stage 6: Review Report ─────────────────────────────────────────────
        report = generate_report(schema, issues)

        return PipelineResult(
            profile=profile,
            table=table,
            schema=schema,
            validation_issues=issues,
            report=report,
            filename=filename,
            warnings=all_warnings,
            table_group=group,
        )


# Aggregate-view section headings ("All Investments Since 2012",
# "Software & Services (2001 - Present)") share the banner shape but name a
# summary cut, not a fund.
_AGGREGATE_HEADING_RE = re.compile(
    r"^\s*all\b"
    r"|\(\s*\d{4}\s*[-–—]\s*(\d{4}|present|current|today)\s*\)",
    re.IGNORECASE,
)


def _looks_like_note_header(v) -> bool:
    """Rows that share the banner shape but aren't fund names: bare 'Notes'
    headings, parenthesised footnote references like '(1) …', prose, and
    aggregate-view headings."""
    s = str(v or "").strip()
    return bool(re.match(r"^\s*(notes?|footnotes?)\s*:?\s*$", s, re.IGNORECASE)
                or s.startswith("(")
                or len(s) > 80 or len(s.split()) > 10    # prose, not a name
                or _AGGREGATE_HEADING_RE.search(s))


FUND_SECTION_COL = "__fund_section__"

# Trailing "(vintage …)" on a fund banner is metadata, not part of the name.
_VINTAGE_TAIL_RE = re.compile(r"\s*\(\s*vintage[^)]*\)\s*$", re.IGNORECASE)


def _inherit_fund_sections(
    table: ExtractedTable,
    schema: SchemaInference,
    warnings: list[str],
) -> None:
    """
    Stage 3.4 — fund names as section banners inside the company column.

    Structural signal only: a banner row has text in the company column and
    every other cell blank (and doesn't look like a footnote or total).
    Deals below a banner inherit it as their fund, until the next banner.
    Engages only when inference found no usable fund column — the per-tab
    synthetic fallback doesn't count; a real mapped fund column always wins.
    Banner rows themselves are removed from the deal set. Tabs (or leading
    rows) without banners keep the tab-name fallback where one exists.
    """
    df = table.df
    comp = schema.field_to_col.get("company")
    if df.empty or not comp or comp not in df.columns:
        return
    fund_col = schema.field_to_col.get("fund")
    if (
        fund_col
        and fund_col in df.columns
        and fund_col != FUND_FROM_TAB_COL
        and df[fund_col].map(lambda v: not _is_blank(v)).mean() > 0.5
    ):
        return                                     # real fund data exists

    other_cols = [c for c in df.columns if c not in (comp, FUND_FROM_TAB_COL)]
    if not other_cols:
        return

    comp_vals = df[comp]
    others_blank = df[other_cols].apply(lambda col: col.map(_is_blank)).all(axis=1)
    is_banner = (
        comp_vals.map(lambda v: not _is_blank(v))
        & others_blank
        & ~comp_vals.map(_looks_like_annotation)
        & ~comp_vals.map(_looks_like_note_header)
    )
    if not is_banner.any():
        return

    section = comp_vals.where(is_banner).map(
        lambda v: v if _is_blank(v) else _VINTAGE_TAIL_RE.sub("", str(v).strip()))
    if FUND_FROM_TAB_COL in df.columns:
        section = section.groupby(df[FUND_FROM_TAB_COL], sort=False).ffill()
    else:
        section = section.ffill()
    if not (section.notna() & ~is_banner).any():
        return                                     # banners head no deals

    df = df.copy()
    if FUND_FROM_TAB_COL in df.columns:
        section = section.where(section.notna(), df[FUND_FROM_TAB_COL])
    df[FUND_SECTION_COL] = section
    n_banners = int(is_banner.sum())
    table.df = df.loc[~is_banner].reset_index(drop=True)
    table.row_count = len(table.df)
    schema.field_to_col["fund"] = FUND_SECTION_COL
    schema.confidences["fund"] = max(schema.confidences.get("fund", 0.0), 0.92)
    schema.explanations["fund"] = [
        f"inherited from {n_banners} fund section banner(s) in the "
        f"'{comp}' column"
    ]
    if FUND_SECTION_COL in schema.unmapped_cols:
        schema.unmapped_cols.remove(FUND_SECTION_COL)
    warnings.append(
        f"Fund identity: {n_banners} section banner row(s) in the '{comp}' "
        f"column were promoted to fund names for the deals beneath them."
    )


def _apply_fund_fallback_and_dedup(
    table: ExtractedTable,
    schema: SchemaInference,
    warnings: list[str],
) -> None:
    """
    For a concatenated per-fund frame:
      1. If inference found no usable `fund` column, adopt the tab-derived one.
      2. De-duplicate rows on (company, fund, entry_date) so a deal listed in
         more than one tab is counted once.
    Mutates `table` and `schema` in place.
    """
    df = table.df
    fund_col = schema.field_to_col.get("fund")
    fund_missing = (
        fund_col is None
        or fund_col not in df.columns
        or df[fund_col].isna().mean() > 0.5
    )
    if fund_missing and FUND_FROM_TAB_COL in df.columns:
        schema.field_to_col["fund"] = FUND_FROM_TAB_COL
        schema.confidences["fund"] = max(schema.confidences.get("fund", 0.0), 0.90)
        schema.explanations["fund"] = ["derived from per-fund tab name"]

    # The tab-derived column is internal plumbing — never show it as "unmapped".
    if FUND_FROM_TAB_COL in schema.unmapped_cols:
        schema.unmapped_cols.remove(FUND_FROM_TAB_COL)

    key_cols = [
        schema.field_to_col.get(k)
        for k in ("company", "fund", "entry_date")
    ]
    key_cols = [c for c in key_cols if c and c in df.columns]
    if key_cols:
        before = len(df)
        deduped = df.drop_duplicates(subset=key_cols, keep="first").reset_index(drop=True)
        removed = before - len(deduped)
        if removed > 0:
            table.df = deduped
            table.row_count = len(deduped)
            warnings.append(
                f"Removed {removed} duplicate deal row(s) across tabs "
                f"(matched on {', '.join(key_cols)})."
            )


# Text in a fund/company cell that marks a row as documentation rather than a
# deal (footnotes, disclaimers, source notes). Generalised — no GP-specific text.
_ANNOTATION_RE = re.compile(
    r"^\s*[-–—•·▪◦*†‡]\s+\S"                          # leading bullet: "- …", "* …", "• …"
    r"|\b(footnote|refer\s+to|data\s*room|please\s+(refer|see|note)"
    r"|disclaimer|source\s*:|represents?\s+schedules?|important\s+note"
    r"|data\s+template|for\s+illustrative|past\s+performance"
    r"|not\s+indicative|the\s+calculation\s+reflects|column\s+reference)\b",
    re.IGNORECASE,
)
# Aggregate/summary rows that sit at the edge of a table.
_TOTAL_RE = re.compile(
    r"^\s*(grand\s+total|sub\s*total|total|portfolio\s+total|aggregate|sum)\s*$",
    re.IGNORECASE,
)

# Financial fields — a numeric value in any of these marks a row as a real deal.
_FINANCIAL_FIELDS = (
    "ic_total", "ic_initial", "realized", "unrealized", "total_value",
    "gross_moic", "gross_irr", "entry_rev", "entry_ebitda", "entry_net_debt",
    "entry_ev", "exit_rev", "exit_ebitda", "exit_net_debt", "exit_ev",
)


def _is_blank(v) -> bool:
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return str(v).strip() in ("", "-", "n/a", "N/A", "#N/A", "nan", "NaT")


def _is_number(v) -> bool:
    if v is None or isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return v == v                            # not NaN
    s = str(v).strip()
    if s in ("", "-", "n/a", "N/A", "#N/A"):
        return False
    try:
        float(s.replace(",", "").replace("%", "").replace("x", "").replace("€", "").replace("$", ""))
        return True
    except (ValueError, TypeError):
        return False


def _looks_like_annotation(v) -> bool:
    # High-precision: only explicit footnote/disclaimer/total markers. A length
    # heuristic was tried but misfires on description-type company columns.
    s = str(v or "").strip()
    if not s:
        return False
    return bool(_ANNOTATION_RE.search(s) or _TOTAL_RE.match(s))


# A company cell holding ONLY a documentation heading ("Notes", "Source:") or
# opening with a numbered footnote marker ("(3) Includes …"). Both only mark a
# row as junk together with the no-date/no-financials guard in stage 3.6 — a
# real deal named like this would carry data and survive.
_NOTE_WORD_RE = re.compile(
    r"^\s*(?:(?:foot)?notes?|sources?|disclaimers?|definitions?|abbreviations?)\s*:?\s*$",
    re.IGNORECASE,
)
_FOOTNOTE_MARKER_RE = re.compile(r"^\s*\(\d{1,2}\)\s+\S")


def _drop_non_deal_rows(
    table: ExtractedTable,
    schema: SchemaInference,
    warnings: list[str],
) -> None:
    """
    Remove rows that are not real deals: footnote/annotation rows, total
    rows, and rows with no identifying information at all (no company, no
    fund, no entry date). Generalised — driven by cell content, not GP name.
    Mutates `table` in place.
    """
    df = table.df
    if df.empty:
        return

    def _mapped(fid):
        c = schema.field_to_col.get(fid)
        return c if (c and c in df.columns) else None

    comp_col = _mapped("company")
    date_col = _mapped("entry_date")
    fund_col = _mapped("fund")
    fin_cols = [c for c in (_mapped(f) for f in _FINANCIAL_FIELDS) if c]

    # A row with none of company / entry date / any financial is not a deal —
    # it's a note or spacer. This is only trustworthy when ≥2 real signal
    # columns are mapped; otherwise an unmapped column would look empty
    # everywhere and wipe real deals.
    signal_cols = [c for c in (comp_col, date_col) if c] + fin_cols
    use_signal = len(signal_cols) >= 2

    def _has_deal_data(row) -> bool:
        if comp_col and not _is_blank(row.get(comp_col)):
            return True
        if date_col and not _is_blank(row.get(date_col)):
            return True
        return any(_is_number(row.get(c)) for c in fin_cols)

    # Fund-name labels appearing in the company column mark section-header /
    # subtotal rows (e.g. a "Fund VIII" banner above that fund's deals).
    fund_values: set[str] = set()
    if fund_col:
        fund_values = {str(v).strip() for v in df[fund_col].dropna()
                       if str(v).strip()}

    drop_idx = []
    for idx, row in df.iterrows():
        # 1) Explicit footnote / disclaimer / total text in an identifier cell.
        if _looks_like_annotation(row.get(fund_col) if fund_col else None) \
           or _looks_like_annotation(row.get(comp_col) if comp_col else None):
            drop_idx.append(idx)
            continue
        # 1b) Company cell holds a fund name and the row has no inv-date →
        #     section header / fund subtotal, not a deal.
        if comp_col and fund_values:
            comp_v = str(row.get(comp_col) or "").strip()
            if comp_v in fund_values and _is_blank(row.get(date_col) if date_col else None):
                drop_idx.append(idx)
                continue
        # 1c) Sentence-like text in the company cell on a row with no entry
        #     date and no financial values → narrative footnote row (e.g. a
        #     per-company "Notes:" block under the table). Deals with
        #     description-style company columns carry dates/financials, which
        #     keeps this away from them (the bare length heuristic misfired).
        if comp_col and (date_col or fin_cols):
            comp_v = str(row.get(comp_col) or "").strip()
            if (comp_v
                    and (len(comp_v) > 60 or comp_v.endswith(":")
                         or len(comp_v.split()) >= 10)
                    and _is_blank(row.get(date_col) if date_col else None)
                    and not any(_is_number(row.get(c)) for c in fin_cols)):
                drop_idx.append(idx)
                continue
        # 1d) Documentation heading ("Notes") or numbered footnote marker
        #     ("(3) Includes …") in the company cell, on a row with no entry
        #     date and no financial values → footnote block row.
        if comp_col and (date_col or fin_cols):
            comp_v = str(row.get(comp_col) or "").strip()
            if ((_NOTE_WORD_RE.match(comp_v) or _FOOTNOTE_MARKER_RE.match(comp_v))
                    and _is_blank(row.get(date_col) if date_col else None)
                    and not any(_is_number(row.get(c)) for c in fin_cols)):
                drop_idx.append(idx)
                continue
        # 1e) Neither a company nor an entry date, and at most ONE mapped
        #     financial value → per-fund subtotal / spacer row (its lone
        #     number is a section sum). A row with several financial values
        #     but no identity is kept: GPs (or manual anonymisation) blank
        #     out names of confidential deals while leaving their data.
        if use_signal and comp_col and date_col \
           and _is_blank(row.get(comp_col)) and _is_blank(row.get(date_col)) \
           and sum(1 for c in fin_cols if _is_number(row.get(c))) <= 1:
            drop_idx.append(idx)
            continue
        # 2) No company, no entry date, no financials → note / spacer row.
        if use_signal and not _has_deal_data(row):
            drop_idx.append(idx)

    if drop_idx:
        table.df = df.drop(index=drop_idx).reset_index(drop=True)
        table.row_count = len(table.df)
        warnings.append(
            f"Dropped {len(drop_idx)} non-deal row(s) (footnotes / totals / empty rows)."
        )
