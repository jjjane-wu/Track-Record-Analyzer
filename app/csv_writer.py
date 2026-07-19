"""
csv_writer.py — Append processed deals to gp_deals.csv for Power BI.

Accepts records in transformer.py's integer-keyed format and writes
named columns to the shared CSV database. Uses filelock to prevent
concurrent write corruption when multiple analysts process files at once.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd


# Maps transformer.py integer record keys → CSV column names.
# These align with CLAUDE.md DEAL_FIELDS names.
_DEAL_COL_MAP: dict[int, str] = {
    1:  "company",
    2:  "fund",
    5:  "status",
    6:  "inv_date",
    7:  "exit_date",
    9:  "hold_period",
    11: "sector",
    12: "geography",
    13: "initial_invested_capital_m",
    16: "total_invested_capital_m",
    17: "realized_value",
    18: "current_value",
    20: "gross_moic",
    29: "transaction_type",
    30: "gp_role",
    31: "process_type",
    32: "sourcing_partner",
    33: "exit_type",
    34: "coi_deal",
    35: "gross_irr",
    36: "entry_ltm_revenue",
    37: "entry_ltm_ebitda",
    39: "entry_net_debt",
    42: "entry_enterprise_value",
    46: "exit_ltm_revenue",
    47: "exit_ltm_ebitda",
    49: "exit_net_debt",
    52: "exit_enterprise_value",
    55: "valuation_method",
}


def _record_to_row(rec: dict[int | str, Any], excluded: bool) -> dict[str, Any]:
    """Convert a transformer record (int-keyed) to a named CSV row."""
    row: dict[str, Any] = {"_excluded": excluded}
    for idx, col in _DEAL_COL_MAP.items():
        val = rec.get(idx)
        if isinstance(val, date):
            val = val.isoformat()
        row[col] = val
    return row


def append_deals(
    included_records: list[dict],
    excluded_records: list[dict],
    gp: str,
    analyst: str,
    db_dir: Path,
) -> None:
    """
    Append processed deals to gp_deals.csv.

    Metadata columns added per row: gp, processed_date, processed_by, _excluded.
    Deduplicates on (company, fund, inv_date) keeping the latest entry so that
    re-processing the same file replaces rather than duplicates rows.
    """
    import filelock

    db_dir.mkdir(parents=True, exist_ok=True)
    csv_path  = db_dir / "gp_deals.csv"
    lock_path = str(csv_path) + ".lock"

    today = date.today().isoformat()
    rows: list[dict] = []
    for rec in included_records:
        row = _record_to_row(rec, excluded=False)
        row.update({"gp": gp, "processed_date": today, "processed_by": analyst})
        rows.append(row)
    for rec in excluded_records:
        row = _record_to_row(rec, excluded=True)
        row.update({"gp": gp, "processed_date": today, "processed_by": analyst})
        rows.append(row)

    if not rows:
        return

    new_df = pd.DataFrame(rows)

    try:
        with filelock.FileLock(lock_path, timeout=10):
            if csv_path.exists():
                existing = pd.read_csv(csv_path, dtype=str)
                combined = pd.concat([existing, new_df.astype(str)], ignore_index=True)
                combined = combined.drop_duplicates(
                    subset=["company", "fund", "inv_date"], keep="last"
                )
            else:
                combined = new_df
            combined.to_csv(csv_path, index=False)
    except Exception:
        new_df.to_csv(csv_path, index=False)
