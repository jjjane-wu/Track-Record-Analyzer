"""
transformer.py — Data transformation, computed fields, and bucket lookups.

Converts raw GP DataFrame rows into the structured records expected by the
three INPUT tabs: Fund Level Input, Deal Characteristics INPUT, Excluded Deals INPUT.
"""

from __future__ import annotations

import math
import re
from datetime import date, datetime
from typing import Any

import pandas as pd


# ── Status normalisation ───────────────────────────────────────────────────────

# Free-text GP status → template status. Only two output states exist:
#   "Realized"   — fully exited (includes written-off, i.e. exited at/near zero)
#   "Unrealized" — still held (includes active and partially-realized deals
#                  that retain a residual stake)
#
# ORDER MATTERS: normalise_status() returns on the first substring hit, and
# "realiz"/"realis" are substrings of "unrealized" and "partially realized".
# The qualified/negated forms are therefore listed before the bare ones.
_STATUS_MAP = {
    # Still held → Unrealized
    "unrealiz":  "Unrealized",
    "unrealis":  "Unrealized",
    "partial":   "Unrealized",   # partially realized: residual stake still held
    "active":    "Unrealized",
    "current":   "Unrealized",
    # Exited → Realized
    "writ":      "Realized",      # write-off / written off / writ. off
    "realiz":    "Realized",
    "realis":    "Realized",
    "exited":    "Realized",
    "divested":  "Realized",
    "sold":      "Realized",
}


def normalise_status(raw: Any) -> str:
    if raw is None:
        return "Unrealized"
    s = str(raw).lower().strip()
    for key, mapped in _STATUS_MAP.items():
        if key in s:
            return mapped
    return "Unrealized"


# ── Date helpers ───────────────────────────────────────────────────────────────

def _to_date(v: Any) -> date | None:
    if v is None:
        return None
    # pandas NaT / NaN are datetime subclasses; catch them before the isinstance
    # checks below, or NaT.date() leaks a NaT into date arithmetic.
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s or s in ("-", "n/a", "N/A", "#N/A", ""):
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y",
                "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s[:len(fmt)], fmt).date()
        except ValueError:
            pass
    return None


def _safe_float(v: Any) -> float | None:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v) if not math.isnan(float(v)) else None
    s = str(v).strip()
    if s in ("-", "", "n/a", "N/A", "#N/A", "#VALUE!", "#REF!", "#DIV/0!"):
        return None
    try:
        return float(s.replace(",", "").replace("%", "").rstrip("x"))
    except (ValueError, TypeError):
        return None


def holding_period(entry: date | None, exit_: date | None) -> float | None:
    """Return hold period in decimal years (365 days/year, matching template convention)."""
    if entry is None or exit_ is None:
        return None
    return (exit_ - entry).days / 365


# ── Monetary scale normalisation ───────────────────────────────────────────────
# The template expects every monetary figure in MILLIONS, but raw GP files
# report in whatever unit they like (millions, thousands, or absolute currency
# units — e.g. ILPA-format workbooks report plain dollars).

MONETARY_FIELD_IDS: tuple[str, ...] = (
    "ic_initial", "ic_total", "realized", "unrealized", "total_value",
    "entry_rev", "entry_ebitda", "entry_net_debt", "entry_ev",
    "exit_rev", "exit_ebitda", "exit_net_debt", "exit_ev",
)

# Magnitude bands for a monetary column's |median| (deal-level figures):
#   < 1,000      → already millions (or a mis-mapped ratio column) — leave alone
#   [1e3, 5e4)   → ambiguous: mega-cap millions vs thousands — millions unless
#                  a unit hint or the file's other columns prove otherwise
#   [5e4, 1e6)   → impossible as millions (≥ $50bn median) — thousands
#   ≥ 1e6        → absolute currency units
_AMBIG_MIN_MEDIAN    = 1_000.0
_MILLIONS_MAX_MEDIAN = 50_000.0
_ABS_UNIT_MIN_MEDIAN = 1_000_000.0
_RATIO_MAX_MEDIAN    = 100.0   # MOICs/multiples/percents live below this


def detect_monetary_scale(
    df: pd.DataFrame,
    field_map: dict[str, str | None],
    file_unit_hint: str | None = None,
) -> dict[str, float]:
    """
    Per-field multiplier that normalises raw monetary values to millions.

    Signals, strongest first — each gated by value plausibility so footnote
    letters like "EBITDA(k)" or a mis-mapped multiple column can't trigger it:
      1. column-header unit hint  ("'000"/"(k)" → ×1e-3; "($M)"/"(EUR m)" → ×1)
      2. file-level banner declaration (profiler.detect_unit_banner: "k"/"m")
      3. value magnitude, with a file-level consensus for the ambiguous band
         so every column of one file lands on the SAME unit (keeps MOIC =
         TV/IC arithmetic consistent).

    Returns {field_id: factor} containing only factors ≠ 1.
    """
    from extractor import _detect_unit

    med:  dict[str, float] = {}
    hint: dict[str, str | None] = {}
    for fid in MONETARY_FIELD_IDS:
        col = field_map.get(fid)
        if not col or col not in df.columns:
            continue
        vals = pd.to_numeric(df[col], errors="coerce").abs()
        vals = vals[vals > 0]
        if not len(vals):
            continue
        med[fid]  = float(vals.median())
        hint[fid] = _detect_unit(str(col))
    if not med:
        return {}

    # File regime from the largest medians: any column impossible-as-millions
    # decides the unit for its ambiguous-band siblings.
    if any(m >= _ABS_UNIT_MIN_MEDIAN for m in med.values()):
        file_regime = "abs"
    elif any(m >= _MILLIONS_MAX_MEDIAN for m in med.values()):
        file_regime = "k"
    else:
        file_regime = None

    scale: dict[str, float] = {}
    for fid, m in med.items():
        h = hint[fid]
        # 1. Column-header hint. "(k)" only counts as thousands when the
        #    values agree (median ≥ 1e3, i.e. figures ≥ ~$1m) — GP files also
        #    use "(k)" as a footnote letter. An explicit millions marker is
        #    trusted unless the median is impossible for millions.
        if h == "k" and m >= _AMBIG_MIN_MEDIAN:
            scale[fid] = 1e-3
            continue
        if h and h.endswith("m") and m < _MILLIONS_MAX_MEDIAN:
            continue
        # 2. File banner ("$ in thousands" etc.). Skipped for ratio-sized
        #    columns (a mis-mapped MOIC/multiple isn't monetary) and for
        #    medians so large the banner can't apply.
        if file_unit_hint == "k" and _RATIO_MAX_MEDIAN <= m < 1e7:
            scale[fid] = 1e-3
            continue
        if file_unit_hint == "m" and m < _MILLIONS_MAX_MEDIAN:
            continue
        # 3. Magnitude bands + file consensus for the ambiguous band.
        if m >= _ABS_UNIT_MIN_MEDIAN:
            scale[fid] = 1e-6
        elif m >= _AMBIG_MIN_MEDIAN and file_regime == "abs":
            scale[fid] = 1e-6
        elif m >= _AMBIG_MIN_MEDIAN and file_regime == "k":
            scale[fid] = 1e-3
    return {f: s for f, s in scale.items() if s != 1.0}


# ── Bucket lookups ─────────────────────────────────────────────────────────────

def _bucket(value: float | None, thresholds: list[float], labels: list[str]) -> str:
    if value is None or math.isnan(value):
        return ""
    for i, t in enumerate(thresholds):
        if value <= t:
            return labels[i]
    return labels[-1]


HP_THRESHOLDS = [2, 4, 6, 8]
HP_LABELS     = ["<=2 yrs", "2 yrs  - 4 yrs", "4 yrs  - 6 yrs", "6 yrs  - 8 yrs", ">=8 yrs"]

IC_THRESHOLDS = [250, 500, 1000, 1250, 1500]
IC_LABELS     = ["<=$250m", "$250 - $500m", "$500 - $1000m", "$1000 - $1250m", "$1250 - $1500m", ">=$1500m"]

MOIC_THRESHOLDS = [1.0, 2.0, 3.0]
MOIC_LABELS     = ["<=1.0x", "1.0x - 2.0x", "2.0x - 3.0x", ">=3.0x"]

EBITDA_THRESHOLDS = [0, 100, 200, 300, 400]
EBITDA_LABELS     = ["<=$0m", "$0 - $100m", "$100 - $200m", "$200 - $300m", "$300 - $400m", ">=$400m"]

REV_THRESHOLDS = [0, 250, 500, 1000, 1500]
REV_LABELS     = ["<=$0m", "$0 - $250m", "$250 - $500m", "$500 - $1000m", "$1000 - $1500m", ">=$1500m"]

ENTRY_EV_THRESHOLDS = [0, 500, 1000, 1500, 2000]
ENTRY_EV_LABELS     = ["<=$0m", "$0 - $500m", "$500 - $1000m", "$1000 - $1500m", "$1500 - $2000m", ">=$2000m"]

ENTRY_MULT_THRESHOLDS = [0, 5, 10, 15, 20]
ENTRY_MULT_LABELS     = ["<=0x", "0x - 5x", "5x - 10x", "10x - 15x", "15x - 20x", ">=20x"]

EBITDA_MARGIN_THRESHOLDS = [0.20, 0.30, 0.40, 0.50]
EBITDA_MARGIN_LABELS     = ["<=20%", "20% - 30%", "30% - 40%", "40% - 50%", ">=50%"]


def hp_bucket(hp: float | None) -> str:
    return _bucket(hp, HP_THRESHOLDS, HP_LABELS)

def ic_bucket(ic: float | None) -> str:
    return _bucket(ic, IC_THRESHOLDS, IC_LABELS)

def moic_bucket(moic: float | None) -> str:
    return _bucket(moic, MOIC_THRESHOLDS, MOIC_LABELS)

def ebitda_bucket(ebitda: float | None) -> str:
    return _bucket(ebitda, EBITDA_THRESHOLDS, EBITDA_LABELS)

def rev_bucket(rev: float | None) -> str:
    return _bucket(rev, REV_THRESHOLDS, REV_LABELS)

def entry_ev_bucket(ev: float | None) -> str:
    return _bucket(ev, ENTRY_EV_THRESHOLDS, ENTRY_EV_LABELS)

def entry_mult_bucket(mult: float | None) -> str:
    return _bucket(mult, ENTRY_MULT_THRESHOLDS, ENTRY_MULT_LABELS)

def ebitda_margin_bucket(margin: float | None) -> str:
    return _bucket(margin, EBITDA_MARGIN_THRESHOLDS, EBITDA_MARGIN_LABELS)


# ── Safe division ──────────────────────────────────────────────────────────────

def _div(a: float | None, b: float | None) -> float | None:
    if a is None or b is None or b == 0:
        return None
    return a / b


# ── CAGR ──────────────────────────────────────────────────────────────────────

def _cagr(start: float | None, end: float | None, years: float | None) -> float | None:
    if start is None or end is None or years is None or years <= 0:
        return None
    if start <= 0 or end <= 0:
        return None
    try:
        return (end / start) ** (1.0 / years) - 1.0
    except (ZeroDivisionError, ValueError):
        return None


# ── Main row transformer ───────────────────────────────────────────────────────

def transform_row(
    raw: dict[str, Any],
    field_map: dict[str, str | None],
    fund_vintage_map: dict[str, int],
    scale_map: dict[str, float] | None = None,
) -> dict[str, Any]:
    """
    Convert one raw GP row (keyed by column name) into a flat dict of
    Deal Characteristics INPUT fields (keyed by template column index).

    field_map: {field_id -> col_name_in_raw}  (from mapper.build_mapping)
    fund_vintage_map: {fund_name -> vintage_year}
    scale_map: {field_id -> multiplier} from detect_monetary_scale();
               normalises monetary fields to millions
    """
    scale_map = scale_map or {}

    def get(field_id: str) -> Any:
        col = field_map.get(field_id)
        return raw.get(col) if col else None

    def money(field_id: str) -> float | None:
        v = _safe_float(get(field_id))
        if v is None:
            return None
        return v * scale_map.get(field_id, 1.0)

    # ── Core identifiers ───────────────────────────────────────────────
    company      = str(get("company") or "").strip()
    fund         = str(get("fund") or "").strip()
    # Mixed columns (e.g. fund section headers inside the company column):
    # a "company" that is actually a fund name is a section label, not a deal.
    if company and (company == fund or company in fund_vintage_map):
        company = ""
    unique_id    = f"{company}-{fund}" if company and fund else ""
    vintage      = fund_vintage_map.get(fund)
    status       = normalise_status(get("status"))

    # ── Dates / hold period ────────────────────────────────────────────
    entry_date = _to_date(get("entry_date"))
    exit_date  = _to_date(get("exit_date"))
    exit_year  = exit_date.year if exit_date else None

    # Always compute hold period exactly from dates for precision;
    # raw value is often rounded to 1 decimal and causes mismatch.
    hp = holding_period(entry_date, exit_date)

    # ── Classifications ────────────────────────────────────────────────
    sector      = str(get("sector") or "").strip()
    geography   = str(get("region") or "").strip()
    tx_type     = str(get("transaction_type") or "").strip()
    role        = str(get("role") or "").strip()
    process     = str(get("competition") or "").strip()
    src_partner = str(get("sourcing_partner") or "").strip()
    exit_type   = str(get("exit_type") or "").strip()
    # COI Deal: raw GP files may include a co-investment column; default "No"
    coi_raw = get("coi_deal")
    if coi_raw is None:
        coi_deal = "n/a"   # older funds that don't report COI
    else:
        s = str(coi_raw).strip().lower()
        coi_deal = "Yes" if s in ("yes", "y", "true", "1") else "No"

    # ── Capital / values ───────────────────────────────────────────────
    ic_initial  = money("ic_initial")
    ic_total    = money("ic_total")
    realized    = money("realized")
    unrealized  = money("unrealized")
    total_value = money("total_value")
    gross_moic  = _safe_float(get("gross_moic"))
    gross_irr   = _safe_float(get("gross_irr"))
    val_method  = str(get("valuation_method") or "").strip()
    if val_method in ("-", "None", "none"):
        val_method = ""

    # If total_value not available, compute from realized + unrealized
    if total_value is None and realized is not None and unrealized is not None:
        total_value = realized + unrealized

    # If gross_moic not available, compute
    if gross_moic is None and ic_total and ic_total > 0 and total_value is not None:
        gross_moic = total_value / ic_total

    # Use ic_total as ic bucket base (matches template col14)
    ic_for_bucket = ic_total if ic_total is not None else ic_initial

    # ── Performing / loss ─────────────────────────────────────────────
    # performing = 1 (underperforming) whenever MOIC < 1, regardless of status.
    # For realized deals with MOIC >= 1, use '-'.
    # For unrealized deals with MOIC >= 1, use 0.
    is_under = gross_moic is not None and gross_moic < 1.0
    if is_under:
        performing = 1
    elif status == "Realized":
        performing = "-"
    else:
        performing = None   # Unrealized MOIC ≥ 1: leave blank (not 0)

    ic_in_loss = ic_for_bucket if is_under else 0
    # Impaired Value = IC - Total Value (positive number representing the loss magnitude)
    impaired   = ((ic_for_bucket or 0) - (total_value or 0)) if is_under else 0

    # ── Entry financial metrics ────────────────────────────────────────
    entry_rev    = money("entry_rev")
    entry_ebitda = money("entry_ebitda")
    entry_nd     = money("entry_net_debt")
    entry_ev     = money("entry_ev")

    entry_margin   = _div(entry_ebitda, entry_rev)
    entry_leverage = _div(entry_nd, entry_ebitda)
    entry_eq_val   = (entry_ev - entry_nd) if (entry_ev is not None and entry_nd is not None) else None
    entry_ebitda_m = _div(entry_ev, entry_ebitda)
    entry_eq_m     = _div(entry_eq_val, entry_ebitda)   # entry equity multiple = Equity Value / EBITDA
    entry_ev_sales = _div(entry_ev, entry_rev)

    # ── Exit / current financial metrics ──────────────────────────────
    exit_rev     = money("exit_rev")
    exit_ebitda  = money("exit_ebitda")
    exit_nd      = money("exit_net_debt")
    exit_ev      = money("exit_ev")

    exit_margin   = _div(exit_ebitda, exit_rev)
    exit_leverage = _div(exit_nd, exit_ebitda)
    exit_eq_val   = (exit_ev - exit_nd) if (exit_ev is not None and exit_nd is not None) else None
    exit_ebitda_m = _div(exit_ev, exit_ebitda)
    exit_ev_sales = _div(exit_ev, exit_rev)

    # ── IC-weighted metrics ────────────────────────────────────────────
    # Pattern: AdjInvCap[X] = ic_total if metric available else 0
    #          Wgtd[X] = AdjInvCap[X] * X
    ic_wgt = ic_total if ic_total is not None else 0

    adj_ic_rev = ic_wgt if (entry_rev is not None and exit_rev is not None) else 0
    rev_cagr   = _cagr(entry_rev, exit_rev, hp)
    wgtd_rev_cagr = (adj_ic_rev * rev_cagr) if (adj_ic_rev and rev_cagr is not None) else 0

    adj_ic_ebitda = ic_wgt if (entry_ebitda is not None and exit_ebitda is not None) else 0
    ebitda_cagr   = _cagr(entry_ebitda, exit_ebitda, hp)
    wgtd_ebitda_cagr = (adj_ic_ebitda * ebitda_cagr) if (adj_ic_ebitda and ebitda_cagr is not None) else 0

    adj_ic_margin = ic_wgt if (entry_margin is not None and exit_margin is not None) else 0
    wgtd_entry_margin = (adj_ic_margin * entry_margin) if adj_ic_margin else 0
    wgtd_exit_margin  = (adj_ic_margin * exit_margin) if adj_ic_margin else 0

    wgtd_hp = ic_wgt * hp if hp is not None else 0

    adj_ic_mult = ic_wgt if (entry_ebitda_m is not None and exit_ebitda_m is not None) else 0
    wgtd_entry_mult = (adj_ic_mult * entry_ebitda_m) if adj_ic_mult else 0
    wgtd_exit_mult  = (adj_ic_mult * exit_ebitda_m) if adj_ic_mult else 0

    adj_ic_lev = ic_wgt if (entry_leverage is not None and exit_leverage is not None) else 0
    wgtd_entry_lev = (adj_ic_lev * entry_leverage) if adj_ic_lev else 0
    wgtd_exit_lev  = (adj_ic_lev * exit_leverage) if adj_ic_lev else 0

    adj_ic_ev_sales = ic_wgt if (entry_ev_sales is not None and exit_ev_sales is not None) else 0
    wgtd_entry_ev_sales = (adj_ic_ev_sales * entry_ev_sales) if adj_ic_ev_sales else 0
    wgtd_exit_ev_sales  = (adj_ic_ev_sales * exit_ev_sales) if adj_ic_ev_sales else 0

    adj_ic_eq_mult = ic_wgt if (entry_leverage is not None and entry_eq_m is not None) else 0
    wgtd_entry_debt_mult = (adj_ic_eq_mult * entry_leverage) if adj_ic_eq_mult else 0
    wgtd_entry_eq_mult   = (adj_ic_eq_mult * entry_eq_m) if adj_ic_eq_mult else 0

    # ── Assemble result keyed by template column index (1-based) ──────
    return {
        1:  company,
        2:  fund,
        3:  unique_id,
        4:  vintage,
        5:  status,
        6:  entry_date,
        7:  exit_date,
        8:  exit_year,
        9:  hp,
        10: hp_bucket(hp),
        11: sector,
        12: geography,
        13: ic_initial,
        14: ic_for_bucket,
        15: ic_bucket(ic_for_bucket),
        16: ic_total,
        17: realized,
        18: unrealized,
        19: total_value,
        20: gross_moic,
        21: moic_bucket(gross_moic),
        22: performing,
        23: ic_in_loss,
        24: impaired,
        25: entry_ebitda,
        26: ebitda_bucket(entry_ebitda),
        27: entry_rev,
        28: rev_bucket(entry_rev),
        29: tx_type,
        30: role,
        31: process,
        32: src_partner,
        33: exit_type,
        34: coi_deal,
        35: gross_irr,  # Gross IRR as reported by the GP (decimal)
        36: entry_rev,
        37: entry_ebitda,
        38: entry_margin,
        39: entry_nd,
        40: entry_leverage,
        41: entry_eq_val,
        42: entry_ev,
        43: entry_ebitda_m,
        44: entry_eq_m,
        45: entry_ev_sales,
        46: exit_rev,
        47: exit_ebitda,
        48: exit_margin,
        49: exit_nd,
        50: exit_leverage,
        51: exit_eq_val,
        52: exit_ev,
        53: exit_ebitda_m,
        54: exit_ev_sales,
        55: val_method or None,
        56: entry_ev_bucket(entry_ev),
        57: entry_mult_bucket(entry_ebitda_m),
        58: ebitda_margin_bucket(entry_margin),
        59: adj_ic_rev,
        60: rev_cagr,
        61: wgtd_rev_cagr,
        62: adj_ic_ebitda,
        63: ebitda_cagr,
        64: wgtd_ebitda_cagr,
        65: adj_ic_margin,
        66: wgtd_entry_margin,
        67: wgtd_exit_margin,
        68: wgtd_hp,
        69: adj_ic_mult,
        70: wgtd_entry_mult,
        71: wgtd_exit_mult,
        72: adj_ic_lev,
        73: wgtd_entry_lev,
        74: wgtd_exit_lev,
        75: adj_ic_ev_sales,
        76: wgtd_entry_ev_sales,
        77: wgtd_exit_ev_sales,
        78: adj_ic_eq_mult,
        79: wgtd_entry_debt_mult,
        80: wgtd_entry_eq_mult,
    }


def compute_fund_vintages(df: pd.DataFrame, fund_col: str, date_col: str) -> dict[str, int]:
    """Return {fund_name: vintage_year} using earliest entry date per fund."""
    result = {}
    for _, row in df.iterrows():
        fund = str(row.get(fund_col, "") or "").strip()
        d = _to_date(row.get(date_col))
        if fund and d:
            if fund not in result or d.year < result[fund]:
                result[fund] = d.year
    return result


def compute_fund_aggregates(
    records: list[dict[str, Any]],
    fund_key: int = 2,
    ic_key: int = 16,
    realized_key: int = 17,
    unrealized_key: int = 18,
    total_key: int = 19,
) -> dict[str, dict[str, float]]:
    """
    Compute fund-level aggregates (IC, Realized, Unrealized, Total, GrossM)
    from a list of transformed deal records.
    Keys are fund names; values are {'ic', 'realized', 'unrealized', 'total', 'gross_moic'}.
    """
    agg: dict[str, dict[str, float]] = {}
    for rec in records:
        fund = str(rec.get(fund_key) or "")
        if not fund:
            continue
        if fund not in agg:
            agg[fund] = {"ic": 0.0, "realized": 0.0, "unrealized": 0.0, "total": 0.0}
        agg[fund]["ic"]         += rec.get(ic_key) or 0
        agg[fund]["realized"]   += rec.get(realized_key) or 0
        agg[fund]["unrealized"] += rec.get(unrealized_key) or 0
        agg[fund]["total"]      += rec.get(total_key) or 0

    for fund, vals in agg.items():
        vals["gross_moic"] = vals["total"] / vals["ic"] if vals["ic"] > 0 else 0

    return agg


def flag_excluded_deals(
    records: list[dict[str, Any]],
    pre_year: int = 2008,
    moic_threshold: float = 0.0,
) -> tuple[list[dict], list[dict]]:
    """All deals go to included; excluded tab is left empty."""
    return list(records), []
