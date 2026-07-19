"""
validator.py — Stage 5: Data Validation

Validates the standardised DataFrame produced by Stage 4.
Rules are grouped by kind:

  Field-level   — checks on a single column in isolation (MOIC > 0, IRR range)
  Row-level      — checks comparing two columns in the same row (entry_date < exit_date)
  Table-level    — statistical checks across all rows (extreme outliers, empty required fields)

No GP-specific knowledge is encoded here.  Rules express general PE data quality
invariants that hold regardless of which firm produced the data.

Adding a new rule: append a RuleSpec to _FIELD_RULES or _ROW_RULES.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Callable, Optional

import pandas as pd

from inferencer import SchemaInference


# ═══════════════════════════════════════════════════════════════════════════════
# Data structures
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class ValidationIssue:
    """One data-quality problem found in the standardised DataFrame."""
    severity: str                 # "error" | "warning"
    kind: str                     # "field" | "row" | "table"
    field_id: str                 # standardised field name (e.g. "gross_moic")
    row_index: Optional[int]      # DataFrame iloc, or None for table-level issues
    value: Any                    # the offending value
    message: str                  # human-readable description

    def label(self) -> str:
        """Short one-line summary."""
        location = f"row {self.row_index}" if self.row_index is not None else "table"
        return f"[{self.severity.upper()}] {self.field_id} @ {location}: {self.message}"


# ═══════════════════════════════════════════════════════════════════════════════
# Rule definitions
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class _FieldRule:
    field_id: str
    check: Callable[[Any], bool]    # returns True if value is VALID
    severity: str
    message: str                    # shown when check returns False
    requires_numeric: bool = True   # skip non-numeric values silently


_FIELD_RULES: list[_FieldRule] = [
    # ── MOIC ──────────────────────────────────────────────────────────────────
    _FieldRule("gross_moic", lambda v: v > 0,   "error",   "MOIC must be positive"),
    _FieldRule("gross_moic", lambda v: v < 100, "warning", "MOIC > 100× is highly unusual — verify data"),

    # ── IRR ───────────────────────────────────────────────────────────────────
    _FieldRule("gross_irr", lambda v: -1.0 < v < 20.0, "warning",
               "IRR outside plausible range (−100% to 2000%)"),

    # ── Invested capital ──────────────────────────────────────────────────────
    _FieldRule("ic_initial", lambda v: v >= 0, "error",   "Invested capital cannot be negative"),
    _FieldRule("ic_total",   lambda v: v >= 0, "error",   "Invested capital cannot be negative"),

    # ── Value fields ──────────────────────────────────────────────────────────
    _FieldRule("realized",    lambda v: v >= 0, "warning", "Realized value should be non-negative"),
    _FieldRule("unrealized",  lambda v: v >= 0, "warning", "Unrealized value should be non-negative"),
    _FieldRule("total_value", lambda v: v >= 0, "warning", "Total value should be non-negative"),

    # ── Hold period ───────────────────────────────────────────────────────────
    _FieldRule("holding_period", lambda v: 0 < v < 40, "warning",
               "Hold period outside 0–40 year range"),

    # ── Fund ownership ────────────────────────────────────────────────────────
    _FieldRule("fund_ownership", lambda v: 0 < v <= 100, "warning",
               "Ownership percentage should be 0–100"),
]


@dataclass
class _RowRule:
    field_a: str
    field_b: str
    check: Callable[[Any, Any], bool]    # returns True if the pair is VALID
    severity: str
    message: str


_ROW_RULES: list[_RowRule] = [
    # ── Date ordering ─────────────────────────────────────────────────────────
    _RowRule("entry_date", "exit_date",
             lambda a, b: _to_date(a) is None or _to_date(b) is None or _to_date(a) <= _to_date(b),
             "error",
             "Exit Date precedes Entry Date"),

    # ── IC vs realized: extreme scale mismatch suggests unit error ────────────
    # Only fires when IC is a meaningful positive number, so we don't false-
    # positive on tiny rounding values or early-stage €0 investments.
    _RowRule("ic_total", "realized",
             lambda ic, rv: (ic is None or rv is None
                             or not (isinstance(ic, (int, float)) and isinstance(rv, (int, float)))
                             or ic <= 0.5                             # ignore near-zero IC
                             or rv <= ic * 200),                      # 200× is implausible
             "warning",
             "Realized value is > 200× invested capital — possible unit mismatch"),
]


# ═══════════════════════════════════════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _to_date(v: Any) -> Optional[date]:
    if v is None:
        return None
    # pd.NaT is a datetime subclass but .date() raises; treat it as null
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, datetime):
        try:
            return v.date()
        except Exception:
            return None
    if isinstance(v, date):
        return v
    try:
        return pd.to_datetime(v).date()
    except Exception:
        return None


def _to_float(v: Any) -> Optional[float]:
    if v is None:
        return None
    try:
        f = float(v)
        return None if math.isnan(f) or math.isinf(f) else f
    except (TypeError, ValueError):
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# Public API
# ═══════════════════════════════════════════════════════════════════════════════

def validate(
    df: pd.DataFrame,
    schema: SchemaInference,
) -> list[ValidationIssue]:
    """
    Validate the standardised DataFrame against all registered rules.

    Parameters
    ----------
    df     : pd.DataFrame   The raw extracted DataFrame (original column names).
    schema : SchemaInference  Stage 3 output — maps field_id → col_name.

    Returns
    -------
    list[ValidationIssue]  All detected problems.  Empty list = clean data.
    """
    issues: list[ValidationIssue] = []

    # ── Field-level checks ─────────────────────────────────────────────────────
    for rule in _FIELD_RULES:
        col = schema.field_to_col.get(rule.field_id)
        if col is None or col not in df.columns:
            continue
        for iloc, raw_val in enumerate(df[col]):
            val = _to_float(raw_val)
            if val is None:
                if rule.requires_numeric:
                    continue
                val = raw_val
            try:
                if not rule.check(val):
                    issues.append(ValidationIssue(
                        severity=rule.severity,
                        kind="field",
                        field_id=rule.field_id,
                        row_index=iloc,
                        value=raw_val,
                        message=rule.message,
                    ))
            except Exception:
                pass

    # ── Row-level checks ───────────────────────────────────────────────────────
    for rule in _ROW_RULES:
        col_a = schema.field_to_col.get(rule.field_a)
        col_b = schema.field_to_col.get(rule.field_b)
        if col_a is None or col_b is None:
            continue
        if col_a not in df.columns or col_b not in df.columns:
            continue
        for iloc, (raw_a, raw_b) in enumerate(zip(df[col_a], df[col_b])):
            # Normalise: convert NaN → None so lambdas can use `is None` checks
            val_a = None if (isinstance(raw_a, float) and math.isnan(raw_a)) else raw_a
            val_b = None if (isinstance(raw_b, float) and math.isnan(raw_b)) else raw_b
            try:
                if not rule.check(val_a, val_b):
                    issues.append(ValidationIssue(
                        severity=rule.severity,
                        kind="row",
                        field_id=f"{rule.field_a} / {rule.field_b}",
                        row_index=iloc,
                        value=(raw_a, raw_b),
                        message=rule.message,
                    ))
            except Exception:
                pass

    # ── Table-level checks ─────────────────────────────────────────────────────
    issues.extend(_table_checks(df, schema))

    return issues


def _table_checks(
    df: pd.DataFrame,
    schema: SchemaInference,
) -> list[ValidationIssue]:
    """Checks that apply to the table as a whole, not individual rows."""
    issues = []

    # Required fields must have a mapping
    REQUIRED = ["company", "entry_date", "ic_total", "gross_moic"]
    for fid in REQUIRED:
        col = schema.field_to_col.get(fid)
        if col is None:
            issues.append(ValidationIssue(
                severity="warning",
                kind="table",
                field_id=fid,
                row_index=None,
                value=None,
                message=f"Required field '{fid}' could not be mapped to any column",
            ))
        elif col in df.columns:
            fill = df[col].notna().mean()
            if fill < 0.5:
                issues.append(ValidationIssue(
                    severity="warning",
                    kind="table",
                    field_id=fid,
                    row_index=None,
                    value=f"{fill:.0%} filled",
                    message=f"Required field '{fid}' is only {fill:.0%} populated — verify column mapping",
                ))

    # MOIC and IRR must be in compatible units
    # If most gross_irr values are > 5, they are likely percentages × 100 stored as integers
    irr_col = schema.field_to_col.get("gross_irr")
    if irr_col and irr_col in df.columns:
        nums = pd.to_numeric(df[irr_col], errors="coerce").dropna()
        if len(nums) >= 3:
            large_fraction = (nums.abs() > 3.0).mean()
            if large_fraction > 0.5:
                issues.append(ValidationIssue(
                    severity="warning",
                    kind="table",
                    field_id="gross_irr",
                    row_index=None,
                    value=f"mean={nums.mean():.2f}",
                    message="IRR values appear to be in percentage points (e.g. 25 instead of 0.25) — verify unit",
                ))

    return issues
