#!/usr/bin/env python3
"""Headless parse: raw GP file -> inputs-only workbook (no UI).

The command-line half of the VBA pipeline (vba/pipeline.bat):

    python app/headless.py "<raw GP file>.xlsx" -o inputs.xlsx

Runs the full parsing pipeline with the AUTO-ACCEPTED column mapping —
i.e. exactly what Screen 2 would show, without the human review — and
writes a workbook containing only the standard "Deal Level Inputs"
sheet, ready for vba/build.vbs to inject into TR-Analyzer.xlsm.

The mapping summary is printed so the analyst still sees what was
auto-mapped; fields flagged NEEDS-REVIEW or UNMAPPED are the cue to do
this GP in the Streamlit app instead (where the mapping can be fixed).
Exit codes: 0 ok, 1 parse failed, 2 usage.
"""
from __future__ import annotations

import argparse
import pathlib
import re
import sys
import warnings
from datetime import date

warnings.filterwarnings("ignore")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))

import openpyxl                                          # noqa: E402
from pipeline import GPParserPipeline                    # noqa: E402
from transformer import (transform_row, compute_fund_vintages,  # noqa: E402
                         flag_excluded_deals, detect_monetary_scale)
import build_output as bo                                # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("raw", help="raw GP track record .xlsx/.xls")
    ap.add_argument("-o", "--out", default=None,
                    help="output workbook (default: <raw> - Inputs.xlsx)")
    ap.add_argument("--gp", default=None, help="GP name (default: from filename)")
    args = ap.parse_args()

    raw = pathlib.Path(args.raw)
    if not raw.is_file():
        print(f"ERROR: file not found: {raw}")
        return 2
    gp = args.gp or (re.split(r"[ _]Data|\.xls", raw.name)[0].strip() or raw.stem)
    out = pathlib.Path(args.out) if args.out else raw.with_name(f"{raw.stem} - Inputs.xlsx")

    try:
        res = GPParserPipeline().run(raw.read_bytes(), raw.name)
    except Exception as e:
        print(f"ERROR: parse failed: {type(e).__name__}: {e}")
        return 1

    df = res.table.df
    f2c = {f: c for f, c in res.schema.field_to_col.items() if c}
    fund_col = f2c.get("fund")
    date_col = f2c.get("entry_date")
    fvm = {}
    if fund_col and date_col and date_col in df.columns:
        try:
            fvm = compute_fund_vintages(df, fund_col, date_col)
        except Exception:
            pass
    scale_map = detect_monetary_scale(
        df, f2c, file_unit_hint=getattr(res.profile, "unit_banner", None))

    records, row_errs = [], 0
    for _, row in df.iterrows():
        try:
            records.append(transform_row(raw=row.to_dict(), field_map=f2c,
                                         fund_vintage_map=fvm, scale_map=scale_map))
        except Exception:
            row_errs += 1
    included, _ = flag_excluded_deals(records)
    if not included:
        print("ERROR: no deals survived parsing — use the Streamlit app to inspect.")
        return 1

    out.write_bytes(bo.build_inputs_workbook(
        included, gp, currency="USD",
        track_record_date=getattr(res.profile, "report_date", None) or date.today()))

    # ── mapping summary (the analyst's visibility into the auto decisions) ──
    print(f"OK: {out}")
    print(f"deals={len(included)} row_errors={row_errs} "
          f"unit_rescaled_cols={len(scale_map)}")
    sch = res.schema
    if sch.needs_review or sch.unmapped_fields:
        print("NOTE — auto-mapping was not fully confident; consider the app UI:")
        for fid in sch.needs_review:
            print(f"  NEEDS-REVIEW  {fid:<18} -> {sch.field_to_col.get(fid)!r} "
                  f"({sch.confidences.get(fid, 0):.0%})")
        for fid in sch.unmapped_fields:
            print(f"  UNMAPPED      {fid}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
