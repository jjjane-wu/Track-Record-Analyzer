"""
db_publish.py — Publish a *verified* Deal Level Input workbook to the database.

The database is deliberately simple: a folder of CSV snapshots, one file per
GP per as-of date ("GP_2 - 2025-09-30.csv"). Point the folder at a
OneDrive-synced SharePoint library and every published file uploads itself;
Power BI reads the whole folder with its SharePoint/Folder connector and
combines the files into one long table.

Publishing is a separate, human-triggered step — never automatic. The app's
parsed output may contain mapping or data errors, so the analyst first
downloads the Deal Level Input workbook, corrects and verifies it in Excel,
and only then publishes that file here. The workbook is the single source of
truth: the same verified file feeds both TR-Analyzer.xlsm and the database.

Re-publishing the same GP + as-of date overwrites its snapshot (idempotent —
a correction replaces the old rows, nothing duplicates). A new as-of date
creates a new snapshot, preserving the history of the track record over time.

CLI (for scripted use):
    python app/db_publish.py "path/to/[12-Aug-26 - GP_2] - Gross Deal Level Input.xlsx"
    python app/db_publish.py input.xlsx --dir "~/OneDrive/TR Database" --by "Jane"
"""

from __future__ import annotations

import io
import json
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_output import INPUT_COLS   # single source of truth for the schema

ROOT_DIR    = Path(__file__).resolve().parent.parent
DB_DIR      = ROOT_DIR / "database"
CONFIG_PATH = DB_DIR / "db_config.json"

# Record keys (transformer numbering, via INPUT_COLS) by value type.
_NUMERIC_KEYS = {16, 17, 18, 20, 35, 36, 37, 39, 42, 46, 47, 49, 52,
                 13, 95, 96, 98, 99}
_DATE_KEYS    = {6, 7, 91}

_ALLOWED_STATUS = {"realized", "unrealized"}


# Old column names that later releases renamed — normalised on read so
# previously downloaded input files keep working everywhere.
_LEGACY_HEADERS = {"fund currency": "Financials Currency",
                   "deal currency": "Financials Currency",
                   "inv. date": "Investment Date"}


def _clean_header(h: Any) -> str:
    """Collapse newlines/extra spaces: 'Realized\\nValue' -> 'Realized Value'."""
    out = " ".join(str(h).split()) if h is not None else ""
    return _LEGACY_HEADERS.get(out.lower(), out)


# Cleaned expected headers, in schema order, with their record keys.
EXPECTED_HEADERS: list[tuple[str, int]] = [
    (_clean_header(h), k) for h, k in INPUT_COLS
]
_HEADER_TO_KEY = {h.lower(): k for h, k in EXPECTED_HEADERS}


# ═══════════════════════════════════════════════════════════════════════════
# Database snapshot schema — "GP TR Database Uploader Input Sheet - v1"
# ═══════════════════════════════════════════════════════════════════════════
# Each GP's snapshot is one styled workbook whose "Deal Level Inputs" sheet
# matches the team's uploader sheet cell for cell. Per column:
# (header, source, number format, alignment, header block).
# source: int = transformer record key; "gp"/"tr_date"/"fund_ccy" = meta;
# "follow_on" = the template's live table formula Total - Initial.
# block: "id" black header, "entry" grey-blue, "exit" light green.
_F_MONEY_BIG   = '###,###,###.0;\\(###,###,###.0\\);"-"'
_F_MONEY_SMALL = '###,###.0;\\(###,###.0\\);"-"'
_F_DATE        = 'dd\\-mmm\\-yy'
DB_SCHEMA: list[tuple[str, Any, str, str, str]] = [
    ("Track Record\nDate",       "tr_date",   "d-mmm-yy",     "center", "id"),
    ("GP",                        "gp",        "General",      "left",   "id"),
    ("Fund",                      2,           "General",      "center", "id"),
    ("Fund\nCurrency",           "fund_ccy",  "General",      "center", "id"),
    ("Company",                   1,           "General",      "left",   "id"),
    ("Status",                    5,           "General",      "center", "id"),
    ("Investment\nDate",         6,           _F_DATE,        "center", "id"),
    ("Exit\nDate",               7,           _F_DATE,        "center", "id"),
    ("Sector",                    11,          "General",      "left",   "id"),
    ("Geography",                 12,          "General",      "center", "id"),
    ("Initial Invested Capital (mlns)",   13,  _F_MONEY_BIG,   "center", "id"),
    ("Follow-On Invested Capital (mlns)", "follow_on", _F_MONEY_BIG, "center", "id"),
    ("Total Invested Capital (mlns)",     16,  _F_MONEY_BIG,   "center", "id"),
    ("Realized\nValue",          17,          _F_MONEY_BIG,   "center", "id"),
    ("Current\nValue",           18,          _F_MONEY_BIG,   "center", "id"),
    ("Transaction Type",          29,          "General",      "center", "id"),
    ("GP Role",                   30,          "General",      "center", "id"),
    ("Process Type",              31,          "General",      "center", "id"),
    ("Sourcing Partner",          32,          "General",      "center", "id"),
    ("Exit Type",                 33,          "General",      "center", "id"),
    ("COI Deal (Yes/No)",         34,          "General",      "center", "id"),
    ("Gross TVPI",                20,          "0.00\\x",    "center", "id"),
    ("Gross\nIRR",               35,          "0.0%",         "center", "id"),
    ("Valuation Method",          55,          "General",      "center", "id"),
    ("Signing Date",              91,          _F_DATE,        "center", "id"),
    ("Seller",                    92,          "General",      "center", "id"),
    ("Seller Type",               93,          "General",      "center", "id"),
    ("Buyer",                     94,          "General",      "center", "id"),
    ("Fund Ownership %",          95,          "0.0%",         "center", "id"),
    ("GP & Affiliates\nOwnership % (incl. COI)", 96, "0.0%",  "center", "id"),
    ("Company Currency",          97,          "General",      "center", "id"),
    ("COI Amount (mlns)",         98,          _F_MONEY_BIG,   "center", "id"),
    ("# COI LPs",                 99,          "#,##0",        "center", "id"),
    ("Financials Currency",       90,          "General",      "center", "entry"),
    ("Entry LTM\nRevenue",       36,          _F_MONEY_SMALL, "center", "entry"),
    ("Entry LTM\nEBITDA",        37,          _F_MONEY_SMALL, "center", "entry"),
    ("Entry\nNet Debt",          39,          _F_MONEY_SMALL, "center", "entry"),
    ("Entry Enterprise\nValue",  42,          _F_MONEY_SMALL, "center", "entry"),
    ("Entry Multiple Basis",      100,         "General",      "center", "entry"),
    ("Exit LTM\nRevenue",        46,          _F_MONEY_SMALL, "center", "exit"),
    ("Exit LTM\nEBITDA",         47,          _F_MONEY_SMALL, "center", "exit"),
    ("Exit\nNet Debt",           49,          _F_MONEY_SMALL, "center", "exit"),
    ("Exit Enterprise Value",     52,          _F_MONEY_SMALL, "center", "exit"),
    ("Exit Multiple Basis",       101,         "General",      "center", "exit"),
]
# Column widths B..AS, plus the narrow spacer column A, from the template.
_DB_COL_WIDTHS = [19.6, 32.3, 15.9, 13.0, 32.6, 16.9, 18.1, 13.0, 28.4, 27.7,
                  18.1, 13.0, 13.0, 13.0, 13.9, 18.4, 13.0, 13.0, 13.0, 26.1,
                  16.0, 13.0, 13.0, 16.1, 13.0, 20.0, 15.0, 20.0, 14.0, 18.0,
                  13.0, 13.0, 15.0, 11.0, 13.0, 17.4, 17.3, 13.4, 17.4, 15.0,
                  16.9, 20.4, 18.1, 17.4]
# (The template's note row is guidance for the writer, not part of the
# database file: cells with no value are written truly blank.)



# ═══════════════════════════════════════════════════════════════════════════
# Reading the Deal Level Input workbook (our own fixed format)
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class ParsedInput:
    gp: str = ""
    as_of: date | None = None
    currency: str = ""
    headers: list[str] = field(default_factory=list)   # cleaned, as found
    rows: list[dict[str, Any]] = field(default_factory=list)
    issues: list[str] = field(default_factory=list)    # structural warnings
    sheet_name: str = ""
    source_name: str = ""
    header_row: int = 0                                # 1-based Excel row


def _as_date(v: Any) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str) and v.strip():
        try:
            return pd.to_datetime(v.strip(), dayfirst=False).date()
        except Exception:
            return None
    return None


def read_input_workbook(source: bytes | str | Path,
                        source_name: str = "") -> ParsedInput:
    """Parse a Deal Level Input workbook (path or raw bytes) into ParsedInput.

    The layout is our own generated format (meta labels in column B, header
    row starting at 'Company' in column B, data rows below) — but the file may
    have been hand-corrected in Excel, so positions are located by label
    rather than assumed.
    """
    if isinstance(source, (str, Path)):
        p = Path(source)
        data = p.read_bytes()
        source_name = source_name or p.name
    else:
        data = source

    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    try:
        ws = wb["Deal Level Inputs"] if "Deal Level Inputs" in wb.sheetnames \
             else wb[wb.sheetnames[0]]
        grid = [list(r) for r in ws.iter_rows(values_only=True)]
        parsed = ParsedInput(sheet_name=ws.title, source_name=source_name)
    finally:
        wb.close()

    def cell(r: int, c: int) -> Any:          # 0-based, blank-safe
        if 0 <= r < len(grid) and 0 <= c < len(grid[r]):
            return grid[r][c]
        return None

    # ── Meta block: labels in column B, values in column C ──────────────
    for r in range(min(15, len(grid))):
        label = _clean_header(cell(r, 1)).lower()
        v = cell(r, 2)
        if label == "gp name" and v is not None:
            parsed.gp = str(v).strip()
        elif label == "track record date":
            parsed.as_of = _as_date(v)
        elif label == "currency" and v is not None:
            parsed.currency = str(v).strip()

    # ── Header row: 'Company' in column B ───────────────────────────────
    hdr_r = next((r for r in range(min(25, len(grid)))
                  if _clean_header(cell(r, 1)).lower() == "company"), None)
    if hdr_r is None:
        parsed.issues.append(
            "Could not find the header row (no 'Company' cell in column B) — "
            "is this really a Deal Level Input file?")
        return parsed
    parsed.header_row = hdr_r + 1

    headers: list[str] = []
    c = 1
    while True:
        h = _clean_header(cell(hdr_r, c))
        if not h:
            break
        headers.append(h)
        c += 1
    parsed.headers = headers

    expected = [h for h, _ in EXPECTED_HEADERS]
    missing = [h for h in expected if h.lower() not in {x.lower() for x in headers}]
    unknown = [h for h in headers if h.lower() not in _HEADER_TO_KEY]
    if missing:
        parsed.issues.append(
            f"{len(missing)} expected column(s) not found (left blank in the "
            f"database): {', '.join(missing)}")
    if unknown:
        parsed.issues.append(
            f"{len(unknown)} unrecognised column(s) ignored: {', '.join(unknown)}")

    # ── Data rows: below the header until the first fully blank row ─────
    span = len(headers)
    for r in range(hdr_r + 1, len(grid)):
        vals = [cell(r, 1 + j) for j in range(span)]
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in vals):
            break
        row: dict[str, Any] = {}
        for h, v in zip(headers, vals):
            if isinstance(v, str):
                v = v.strip()
            row[h] = v
        parsed.rows.append(row)

    return parsed


# ═══════════════════════════════════════════════════════════════════════════
# Validation — the publish gate
# ═══════════════════════════════════════════════════════════════════════════

def validate(parsed: ParsedInput) -> tuple[list[str], list[str]]:
    """Return (errors, warnings). Errors block publishing; warnings don't."""
    errors: list[str] = []
    warnings: list[str] = list(parsed.issues)

    if not parsed.gp:
        errors.append("GP Name is missing (cell C4 of the input file).")
    if parsed.as_of is None:
        errors.append("Track Record Date is missing or not a date (cell C5).")
    if not parsed.rows:
        errors.append("No deal rows found below the header row.")
    if not parsed.currency:
        warnings.append("Currency is missing — Financials Currency blanks "
                        "cannot be back-filled.")

    def col_of(key: int) -> str | None:
        for h, k in EXPECTED_HEADERS:
            if k == key:
                return next((x for x in parsed.headers
                             if x.lower() == h.lower()), None)
        return None

    numeric_cols = [c for k in _NUMERIC_KEYS if (c := col_of(k))]
    date_cols    = [c for k in _DATE_KEYS if (c := col_of(k))]
    company_col, fund_col = col_of(1), col_of(2)
    status_col, tvpi_col, irr_col = col_of(5), col_of(20), col_of(35)
    ic_col = col_of(16)

    bad_numeric: list[str] = []
    bad_dates: list[str] = []
    seen: dict[tuple[str, str], int] = {}
    n_blank_company = n_blank_fund = n_bad_status = 0
    n_bad_tvpi = n_big_irr = n_no_ic = 0

    for i, row in enumerate(parsed.rows):
        xl_row = parsed.header_row + 1 + i

        company = str(row.get(company_col) or "").strip() if company_col else ""
        if not company:
            n_blank_company += 1
        fund = str(row.get(fund_col) or "").strip() if fund_col else ""
        if not fund:
            n_blank_fund += 1
        if company and fund:
            key = (company.lower(), fund.lower())
            if key in seen:
                warnings.append(
                    f"Row {xl_row}: duplicate deal '{company}' in '{fund}' "
                    f"(first seen row {seen[key]}).")
            else:
                seen[key] = xl_row

        for c in numeric_cols:
            v = row.get(c)
            if v is None or v == "" or isinstance(v, (int, float)):
                continue
            bad_numeric.append(f"{c} row {xl_row}: '{v}'")
        for c in date_cols:
            v = row.get(c)
            if v in (None, "") or _as_date(v) is not None:
                continue
            bad_dates.append(f"{c} row {xl_row}: '{v}'")

        if status_col:
            s = str(row.get(status_col) or "").strip().lower()
            if s and s not in _ALLOWED_STATUS:
                n_bad_status += 1
        if tvpi_col and isinstance(row.get(tvpi_col), (int, float)) \
                and row[tvpi_col] < 0:
            n_bad_tvpi += 1
        if irr_col and isinstance(row.get(irr_col), (int, float)) \
                and abs(row[irr_col]) > 3:
            n_big_irr += 1
        if ic_col and row.get(ic_col) in (None, "", 0):
            n_no_ic += 1

    if n_blank_company:
        errors.append(f"{n_blank_company} row(s) have no Company name.")
    if bad_numeric:
        shown = "; ".join(bad_numeric[:8])
        more = f" … and {len(bad_numeric) - 8} more" if len(bad_numeric) > 8 else ""
        errors.append(f"Non-numeric text in numeric column(s) — fix in Excel "
                      f"before publishing: {shown}{more}")
    if bad_dates:
        shown = "; ".join(bad_dates[:8])
        more = f" … and {len(bad_dates) - 8} more" if len(bad_dates) > 8 else ""
        errors.append(f"Unreadable date(s): {shown}{more}")

    if n_blank_fund:
        warnings.append(f"{n_blank_fund} row(s) have no Fund.")
    if n_bad_status:
        warnings.append(f"{n_bad_status} row(s) have a Status other than "
                        "Realized/Unrealized.")
    if n_bad_tvpi:
        warnings.append(f"{n_bad_tvpi} row(s) have a negative Gross TVPI.")
    if n_big_irr:
        warnings.append(f"{n_big_irr} row(s) have |Gross IRR| > 300% — check "
                        "whether IRR was entered in percentage points "
                        "(25 instead of 0.25).")
    if n_no_ic:
        warnings.append(f"{n_no_ic} row(s) have no Total Invested Capital.")

    return errors, warnings


# ═══════════════════════════════════════════════════════════════════════════
# Long table + publishing
# ═══════════════════════════════════════════════════════════════════════════

def to_long_table(parsed: ParsedInput,
                  published_by: str = "",
                  published_at: datetime | None = None) -> pd.DataFrame:
    """Preview table: one row per deal in DB_SCHEMA column order (headers
    flattened to one line). The published workbook is written by
    build_snapshot_workbook; provenance lives in its hidden sheet."""
    key_to_src = _key_to_src(parsed)
    out_rows: list[dict[str, Any]] = []
    for row in parsed.rows:
        rec: dict[str, Any] = {}
        for col, src_key, _fmt, _al, _blk in DB_SCHEMA:
            name = " ".join(col.split())
            v = _db_value(parsed, row, src_key, key_to_src)
            if isinstance(v, (date, datetime)):
                v = (v.date() if isinstance(v, datetime) else v).isoformat()
            rec[name] = "" if v is None else v
        out_rows.append(rec)
    return pd.DataFrame(out_rows,
                        columns=[" ".join(c.split()) for c, *_ in DB_SCHEMA])


def _key_to_src(parsed: ParsedInput) -> dict[int, str | None]:
    """Input-file column name (cleaned, as read) for each record key."""
    return {key: next((x for x in parsed.headers if x.lower() == h.lower()), None)
            for h, key in EXPECTED_HEADERS}


def _db_value(parsed: ParsedInput, row: dict, src_key,
              key_to_src: dict) -> Any:
    if src_key == "gp":
        return parsed.gp
    if src_key == "tr_date":
        return parsed.as_of
    if src_key == "fund_ccy":
        return parsed.currency or ""
    if src_key == "follow_on":            # preview only; the workbook carries
        t = _safe_num(row.get(key_to_src.get(16) or ""))   # the live formula
        i = _safe_num(row.get(key_to_src.get(13) or ""))
        return t - i if (t is not None and i is not None) else None
    v = row.get(key_to_src.get(src_key) or "", None)
    if src_key in _DATE_KEYS:
        return _as_date(v)
    if src_key == 90 and (v is None or v == "") and parsed.currency:
        return parsed.currency            # back-fill from the meta block
    return v


def _safe_num(v: Any) -> float | None:
    try:
        return float(str(v).replace(",", "")) if v not in (None, "") else None
    except (TypeError, ValueError):
        return None


def build_snapshot_workbook(parsed: ParsedInput, published_by: str = "",
                            published_at: datetime | None = None) -> bytes:
    """The published file: a workbook whose "Deal Level Inputs" sheet matches
    the team's uploader sheet cell for cell (note row, black/entry/exit
    header blocks, blue-on-light-blue centred data cells with hair borders,
    the live Follow-On formula, the GrossDealLevelInput table). Provenance
    goes on a hidden "_Publish" sheet, keeping the visible sheet identical
    to the template."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Color
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter

    published_at = published_at or datetime.now()
    key_to_src = _key_to_src(parsed)

    hdr_black_fill = PatternFill("solid", fgColor=Color(theme=1))
    hdr_entry_fill = PatternFill("solid", fgColor=Color(theme=3, tint=0.6))
    hdr_exit_fill  = PatternFill("solid", fgColor=Color(theme=9, tint=0.8))
    hdr_white_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    hdr_dark_font  = Font(name="Arial", size=10, bold=True)
    data_font      = Font(name="Arial", size=10, color="0000FF")
    data_fill      = PatternFill("solid", fgColor=Color(theme=4, tint=0.8))
    hair           = Border(left=Side(style="hair"), right=Side(style="hair"),
                            top=Side(style="hair"), bottom=Side(style="hair"))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Deal Level Inputs"
    ws.sheet_view.showGridLines = False

    # Pure flat table (per the team): header row IS row 1, data starts in
    # column A — no title, no note row, no spacer column.
    ws.row_dimensions[1].height = 35.45
    for j, (hdr, _src, _fmt, al, blk) in enumerate(DB_SCHEMA):
        cell = ws.cell(row=1, column=1 + j, value=hdr)
        cell.fill = {"id": hdr_black_fill, "entry": hdr_entry_fill,
                     "exit": hdr_exit_fill}[blk]
        cell.font = hdr_white_font if blk == "id" else hdr_dark_font
        cell.alignment = Alignment(horizontal="left" if al == "left" else "center",
                                   vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(1 + j)].width = _DB_COL_WIDTHS[j]

    fo_formula = ('=GrossDealLevelInput[[#This Row],'
                  '[Total Invested Capital (mlns)]]-GrossDealLevelInput'
                  '[[#This Row],[Initial Invested Capital (mlns)]]')
    for i, row in enumerate(parsed.rows):
        r = 2 + i
        for j, (_hdr, src_key, fmt, al, _blk) in enumerate(DB_SCHEMA):
            if src_key == "follow_on":
                cell = ws.cell(row=r, column=1 + j, value=fo_formula)
            else:
                v = _db_value(parsed, row, src_key, key_to_src)
                cell = ws.cell(row=r, column=1 + j,
                               value=None if v in (None, "") else v)
            if fmt != "General":
                cell.number_format = fmt
            cell.font = data_font
            cell.fill = data_fill
            cell.border = hair
            cell.alignment = Alignment(horizontal=al)

    n = max(len(parsed.rows), 1)
    last = get_column_letter(len(DB_SCHEMA))
    tbl = Table(displayName="GrossDealLevelInput", ref=f"A1:{last}{1 + n}")
    tbl.tableStyleInfo = TableStyleInfo(showRowStripes=True)
    ws.add_table(tbl)
    wb.calculation.fullCalcOnLoad = True

    meta = wb.create_sheet("_Publish")
    meta["A1"] = "Source File";  meta["B1"] = parsed.source_name
    meta["A2"] = "Published By"; meta["B2"] = published_by
    meta["A3"] = "Published At"; meta["B3"] = published_at.strftime("%Y-%m-%d %H:%M:%S")
    meta.sheet_state = "hidden"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def snapshot_filename(parsed: ParsedInput) -> str:
    gp = re.sub(r"[\\/:*?\"<>|]", "-", parsed.gp).strip() or "GP"
    as_of = parsed.as_of.isoformat() if parsed.as_of else "undated"
    return f"{gp} - {as_of}.xlsx"


def publish(parsed: ParsedInput, db_dir: str | Path,
            published_by: str = "") -> tuple[Path, bool]:
    """Write the snapshot CSV into db_dir. Returns (path, replaced_existing).

    Same GP + as-of overwrites its previous snapshot (a correction replaces,
    never duplicates). The write goes through a temp file + atomic rename so
    a sync client never sees a half-written file.
    """
    db_dir = Path(db_dir).expanduser()
    db_dir.mkdir(parents=True, exist_ok=True)
    target = db_dir / snapshot_filename(parsed)
    replaced = target.exists()

    payload = build_snapshot_workbook(parsed, published_by=published_by)
    tmp = target.with_suffix(".xlsx.tmp")
    tmp.write_bytes(payload)
    os.replace(tmp, target)
    return target, replaced


def list_snapshots(db_dir: str | Path) -> pd.DataFrame:
    """Inventory of the database folder: one row per snapshot workbook
    (legacy CSV snapshots are still listed)."""
    db_dir = Path(db_dir).expanduser()
    rows = []
    if db_dir.is_dir():
        for p in sorted(db_dir.glob("*.xlsx")):
            if p.name.startswith("~$"):
                continue
            try:
                wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
                ws = wb["Deal Level Inputs"]
                n = 0
                as_of = gp = ""
                for r in ws.iter_rows(min_row=2, min_col=1, max_col=5,
                                      values_only=True):
                    if all(v in (None, "") for v in r):
                        break
                    if n == 0:
                        as_of = r[0].date().isoformat() if isinstance(r[0], datetime) \
                            else (r[0].isoformat() if isinstance(r[0], date) else str(r[0] or ""))
                        gp = str(r[1] or "")
                    n += 1
                by = at = ""
                if "_Publish" in wb.sheetnames:
                    m = wb["_Publish"]
                    by = str(m["B2"].value or "")
                    at = str(m["B3"].value or "")
                wb.close()
                rows.append({"File": p.name, "GP": gp, "As of": as_of,
                             "Deals": n, "Published": at, "By": by})
            except Exception as e:
                rows.append({"File": p.name, "GP": f"(unreadable: {e})",
                             "As of": "", "Deals": 0, "Published": "", "By": ""})
        for p in sorted(db_dir.glob("*.csv")):
            try:
                df = pd.read_csv(p, dtype=str, keep_default_na=False)
                first = df.iloc[0] if len(df) else {}
                rows.append({
                    "File": p.name,
                    "GP": first.get("GP Name", first.get("GP", "")),
                    "As of": first.get("TR Date",
                                       first.get("Track Record Date", "")),
                    "Deals": len(df),
                    "Published": first.get("Published At", ""),
                    "By": first.get("Published By", ""),
                })
            except Exception as e:
                rows.append({"File": p.name, "GP": f"(unreadable: {e})",
                             "As of": "", "Deals": 0, "Published": "", "By": ""})
    return pd.DataFrame(rows,
                        columns=["File", "GP", "As of", "Deals", "Published", "By"])


# ═══════════════════════════════════════════════════════════════════════════
# Config — where the database folder lives
# ═══════════════════════════════════════════════════════════════════════════

def default_deals_dir() -> Path:
    return DB_DIR / "deals"


def load_db_config() -> dict:
    try:
        cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        cfg = {}
    cfg.setdefault("deals_dir", str(default_deals_dir()))
    return cfg


def save_db_config(deals_dir: str | Path) -> None:
    DB_DIR.mkdir(parents=True, exist_ok=True)
    cfg = load_db_config()
    cfg["deals_dir"] = str(Path(deals_dir).expanduser())
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


# ═══════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════

def _main(argv: list[str]) -> int:
    import argparse
    ap = argparse.ArgumentParser(
        description="Publish a verified Deal Level Input workbook to the "
                    "database folder (one CSV snapshot per GP per as-of date).")
    ap.add_argument("input", help="Path to the verified Deal Level Input .xlsx")
    ap.add_argument("--dir", default=None,
                    help="Database folder (default: the configured folder)")
    ap.add_argument("--by", default="", help="Analyst name for provenance")
    ap.add_argument("--allow-warnings", action="store_true",
                    help="Publish even when validation warnings exist "
                         "(errors always block)")
    args = ap.parse_args(argv)

    parsed = read_input_workbook(args.input)
    errors, warns = validate(parsed)
    for w in warns:
        print(f"WARNING: {w}")
    for e in errors:
        print(f"ERROR:   {e}")
    if errors:
        print("Not published — fix the errors in Excel and re-run.")
        return 1
    if warns and not args.allow_warnings:
        print(f"Not published — {len(warns)} warning(s) above. Re-run with "
              "--allow-warnings after confirming they are fine.")
        return 2

    db_dir = args.dir or load_db_config()["deals_dir"]
    path, replaced = publish(parsed, db_dir, published_by=args.by)
    verb = "Replaced snapshot" if replaced else "Published"
    print(f"{verb}: {path}  ({parsed.gp}, as of {parsed.as_of}, "
          f"{len(parsed.rows)} deals)")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main(sys.argv[1:]))
