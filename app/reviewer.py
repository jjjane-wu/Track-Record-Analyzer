"""
reviewer.py — Stage 6: Human Review Report

Aggregates Stage 3 (schema inference) and Stage 5 (validation) outputs into
a single structured ReviewReport that drives the review UI in app.py.

The reviewer never modifies mappings.  Its sole job is to categorise findings
so the user can act on them with a minimum of cognitive overhead.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from inferencer import SchemaInference, FieldInference, CONFIDENCE_AUTO, CONFIDENCE_REVIEW
from validator import ValidationIssue
from mapper import TEMPLATE_FIELDS


# ═══════════════════════════════════════════════════════════════════════════════
# Data structures
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class MappingRow:
    """One row in the mapping review table."""
    field_id: str
    field_label: str
    source_col: Optional[str]
    confidence: float
    signals: list[str]
    status: str          # "confirmed" | "review" | "unmapped"


@dataclass
class ReviewReport:
    """
    Output of Stage 6.

    Provides three views:
      confirmed_mappings  — high-confidence, no action needed
      review_mappings     — medium-confidence, user should verify
      unmapped_fields     — no mapping found, user must provide or accept blank
      unmapped_cols       — source columns not mapped to any field
      validation_issues   — data quality problems (error > warning)
    """
    confirmed_mappings: list[MappingRow]     # confidence ≥ 0.85
    review_mappings: list[MappingRow]        # 0.50 ≤ confidence < 0.85
    unmapped_fields: list[MappingRow]        # confidence < 0.50
    unmapped_cols: list[str]                 # source columns with no match

    validation_errors: list[ValidationIssue]    # severity == "error"
    validation_warnings: list[ValidationIssue]  # severity == "warning"

    # Convenience counts
    @property
    def n_confirmed(self) -> int:
        return len(self.confirmed_mappings)

    @property
    def n_review(self) -> int:
        return len(self.review_mappings)

    @property
    def n_unmapped_fields(self) -> int:
        return len(self.unmapped_fields)

    @property
    def n_issues(self) -> int:
        return len(self.validation_errors) + len(self.validation_warnings)

    @property
    def is_clean(self) -> bool:
        return (
            self.n_review == 0
            and self.n_unmapped_fields == 0
            and len(self.validation_errors) == 0
        )

    def summary(self) -> str:
        parts = [f"{self.n_confirmed} field(s) auto-confirmed"]
        if self.n_review:
            parts.append(f"{self.n_review} need(s) review")
        if self.n_unmapped_fields:
            parts.append(f"{self.n_unmapped_fields} unmapped")
        if len(self.validation_errors):
            parts.append(f"{len(self.validation_errors)} validation error(s)")
        if len(self.validation_warnings):
            parts.append(f"{len(self.validation_warnings)} warning(s)")
        return "; ".join(parts)


# ═══════════════════════════════════════════════════════════════════════════════
# Public API
# ═══════════════════════════════════════════════════════════════════════════════

_LABEL_LOOKUP = {f["id"]: f["label"] for f in TEMPLATE_FIELDS}


def generate_report(
    schema: SchemaInference,
    validation_issues: list[ValidationIssue],
) -> ReviewReport:
    """
    Build a ReviewReport from Stage 3 + Stage 5 outputs.

    Parameters
    ----------
    schema            : SchemaInference       Field mappings + confidences.
    validation_issues : list[ValidationIssue] Data validation findings.

    Returns
    -------
    ReviewReport
    """
    confirmed, review, unmapped = [], [], []

    for field in TEMPLATE_FIELDS:
        fid   = field["id"]
        label = field["label"]
        col   = schema.field_to_col.get(fid)
        conf  = schema.confidences.get(fid, 0.0)
        sigs  = schema.explanations.get(fid, [])

        if conf >= CONFIDENCE_AUTO:
            status = "confirmed"
            confirmed.append(MappingRow(fid, label, col, conf, sigs, status))
        elif conf >= CONFIDENCE_REVIEW:
            status = "review"
            review.append(MappingRow(fid, label, col, conf, sigs, status))
        else:
            status = "unmapped"
            unmapped.append(MappingRow(fid, label, None, 0.0, [], status))

    errors   = [i for i in validation_issues if i.severity == "error"]
    warnings = [i for i in validation_issues if i.severity == "warning"]

    return ReviewReport(
        confirmed_mappings=confirmed,
        review_mappings=review,
        unmapped_fields=unmapped,
        unmapped_cols=schema.unmapped_cols,
        validation_errors=errors,
        validation_warnings=warnings,
    )
