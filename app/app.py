"""
app.py — GP Track Record Analysis Automation
Streamlit three-screen flow: Upload → Mapping Review → Output & Download.
"""

from __future__ import annotations

import io
import os
import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

# ── Path setup ────────────────────────────────────────────────────────────────
APP_DIR      = Path(__file__).resolve().parent
ROOT_DIR     = APP_DIR.parent
CONFIGS_DIR  = ROOT_DIR / "configs"
OUTPUTS_DIR  = ROOT_DIR / "outputs"

for d in [CONFIGS_DIR, OUTPUTS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

sys.path.insert(0, str(APP_DIR))

from parser import parse_gp_file, detect_funds_in_df, detect_deal_sheets
from profiler import (
    profile_workbook, group_candidate_tables, detect_track_record_date,
    detect_unit_banner, WorkbookProfile, TableCandidate, TableGroup,
)
from mapper import (
    build_mapping,
    TEMPLATE_FIELDS, confidence_tier, is_optional,
)
from transformer import (
    transform_row, compute_fund_vintages,
    flag_excluded_deals, normalise_status, detect_monetary_scale,
)
from build_output import build_inputs_workbook
from pipeline import GPParserPipeline, PipelineResult
from inferencer import CONFIDENCE_AUTO, CONFIDENCE_REVIEW
import db_publish

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="GP Track Record Automation",
    page_icon="📊",
    layout="wide",
)

# ── Session state initialisation ──────────────────────────────────────────────
def _init_state():
    defaults = {
        "screen": 1,
        "gp_name": "",           # ANONYMISED — shown in sidebar / written to outputs
        "gp_name_input": "",     # REAL name — kept only for the text-input widget
        "fund_name": "",         # kept empty — output filename falls back to the GP name
        "raw_df": None,
        "meta": None,
        "mapping": None,
        "fund_vintage_map": {},
        "all_funds": [],
        "profile": None,             # WorkbookProfile from profiler (Stage 1)
        "profile_file_id": None,     # tracks which file was profiled
        "upload_bytes": None,        # cached raw bytes from the last upload
        "pipeline_result": None,     # PipelineResult from Stages 1-3+5+6
        "analyst_name": "",          # from Screen 1
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()

# ── Sidebar navigation ────────────────────────────────────────────────────────
with st.sidebar:
    st.title("📊 GP TR Automation")
    st.caption("UNJSPF OIM — Phase 1")
    st.divider()

    app_mode = st.radio(
        "Mode",
        ["Analyze a GP file", "Publish to database"],
        key="app_mode",
        label_visibility="collapsed",
    )
    st.divider()

    if app_mode == "Analyze a GP file":
        screens = {1: "① Upload", 2: "② Mapping Review", 3: "③ Output & Download"}
        for num, label in screens.items():
            active = st.session_state.screen == num
            if active:
                st.markdown(f"**→ {label}**")
            else:
                st.markdown(f"  {label}")

        st.divider()
        if st.session_state.gp_name:
            st.caption(f"GP: **{st.session_state.gp_name}**")

# ─────────────────────────────────────────────────────────────────────────────
# Profile display helper
# ─────────────────────────────────────────────────────────────────────────────

def _display_workbook_profile(profile: "WorkbookProfile") -> None:
    """Render a compact workbook profile summary in the Streamlit UI."""

    st.subheader("Workbook Structure", divider="gray")

    # Top-level metrics row
    visible_count = len(profile.visible_sheets())
    total_tables  = sum(len(sp.candidate_tables) for sp in profile.sheet_profiles)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Sheets", len(profile.sheet_profiles))
    c2.metric("Visible", visible_count)
    c3.metric("Tables Found", total_tables)
    if profile.currency:
        c4.metric("Currency", profile.currency)
    elif profile.report_date:
        c4.metric("Report Date", str(profile.report_date))
    else:
        c4.metric("Report Date", "—")

    # Primary table highlight
    if profile.primary_table:
        pt = profile.primary_table
        hdr_label = (
            f"rows {pt.header_rows[0]}–{pt.header_rows[-1]}"
            if len(pt.header_rows) > 1
            else f"row {pt.header_rows[0]}"
        )
        st.success(
            f"**Primary table**: Sheet **{pt.sheet_name}** | {pt.bounds} | "
            f"{pt.row_count} data rows × {pt.col_count} cols | "
            f"header: {hdr_label} | score {pt.score:.0f}/100"
        )
        if pt.sample_headers:
            chips = "  ·  ".join(f"`{h}`" for h in pt.sample_headers[:12])
            st.caption(f"Detected columns: {chips}")
        for note in pt.notes:
            st.info(note, icon="ℹ️")
    else:
        st.warning("No candidate deal table auto-detected. Select a sheet manually below.")

    # Per-sheet detail (collapsed)
    with st.expander("All sheets detail", expanded=False):
        rows = []
        for sp in profile.sheet_profiles:
            best = sp.candidate_tables[0] if sp.candidate_tables else None
            rows.append({
                "Sheet": sp.name,
                "State": sp.state,
                "Used Range": str(sp.used_range) if sp.used_range else "—",
                "Tables": len(sp.candidate_tables),
                "Best Score": f"{best.score:.0f}" if best else "—",
                "Data Rows": best.row_count if best else "—",
                "Columns": best.col_count if best else "—",
                "Orientation": best.orientation if best else "—",
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # Detected metadata (GP name candidates suppressed — anonymised session)
    meta_lines = []
    if profile.report_date:
        meta_lines.append(f"Report date: **{profile.report_date}**")
    if profile.currency:
        meta_lines.append(f"Currency: **{profile.currency}**")
    if meta_lines:
        st.caption("  |  ".join(meta_lines))

    for w in profile.warnings:
        st.warning(w)


# ─────────────────────────────────────────────────────────────────────────────
# Publish-to-database page (separate mode, outside the 3-screen flow)
# ─────────────────────────────────────────────────────────────────────────────

def _render_publish_page() -> None:
    st.title("Publish to the Deal Database")
    st.markdown(
        "The database only takes **verified** data. First correct and check the "
        "downloaded Deal Level Input file in Excel; when you are sure the "
        "numbers are right, publish that file here. Each publish writes one CSV "
        "snapshot per GP per as-of date into the database folder — re-publishing "
        "the same GP and date **replaces** its snapshot, so corrections never "
        "create duplicates. Power BI reads the folder and combines every "
        "snapshot into one long deal table."
    )

    cfg = db_publish.load_db_config()
    with st.expander("Database folder", expanded=False):
        st.caption(
            "Published snapshots land in this folder. Point it at a folder "
            "inside your OneDrive-synced SharePoint library and every publish "
            "uploads itself automatically — no other steps needed. "
            "Power BI setup notes: `database/POWERBI_SETUP.md`."
        )
        new_dir = st.text_input("Folder path", value=cfg["deals_dir"])
        if st.button("Save folder") and new_dir.strip():
            db_publish.save_db_config(new_dir.strip())
            st.rerun()

    uploaded = st.file_uploader(
        "Verified Deal Level Input file",
        type=["xlsx"],
        help="The corrected [dd-mmm-yy - GP] - Gross Deal Level Input.xlsx",
    )

    if uploaded:
        try:
            parsed = db_publish.read_input_workbook(
                uploaded.getvalue(), source_name=uploaded.name)
        except Exception as e:
            st.error(f"Could not read the file: {e}")
            parsed = None

        if parsed:
            errors, warnings = db_publish.validate(parsed)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("GP", parsed.gp or "—")
            c2.metric("As of", parsed.as_of.isoformat() if parsed.as_of else "—")
            c3.metric("Currency", parsed.currency or "—")
            c4.metric("Deals", len(parsed.rows))

            for e in errors:
                st.error(e, icon="🚨")
            for w in warnings:
                st.warning(w, icon="⚠️")
            if not errors and not warnings:
                st.success("All checks passed.")

            with st.expander("Preview — exactly what will be published", expanded=False):
                st.dataframe(db_publish.to_long_table(parsed).head(10),
                             use_container_width=True, hide_index=True)

            snap_name = db_publish.snapshot_filename(parsed)
            target = Path(cfg["deals_dir"]).expanduser() / snap_name
            if target.exists():
                st.info(f"A snapshot for this GP and as-of date already exists — "
                        f"publishing will replace **{snap_name}**.")

            confirm_ok = True
            if warnings and not errors:
                confirm_ok = st.checkbox(
                    "I've checked the warnings above — the data is correct as-is.")

            by = st.text_input("Published by",
                               value=st.session_state.analyst_name or "")

            if errors:
                st.caption("The errors above must be fixed in Excel before "
                           "this file can be published.")
            if st.button(f"Publish {snap_name} →", type="primary",
                         disabled=bool(errors) or not confirm_ok):
                try:
                    path, replaced = db_publish.publish(
                        parsed, cfg["deals_dir"], published_by=by.strip())
                    st.session_state.analyst_name = by.strip()
                    verb = "Replaced" if replaced else "Published"
                    st.success(f"{verb} **{path.name}** — {len(parsed.rows)} "
                               f"deals → `{path.parent}`")
                except Exception as e:
                    st.error(f"Publish failed: {e}")

    st.divider()
    st.subheader("Database contents")
    snaps = db_publish.list_snapshots(cfg["deals_dir"])
    if len(snaps):
        st.dataframe(snaps, use_container_width=True, hide_index=True)
        st.caption(f"{len(snaps)} snapshot(s), {int(snaps['Deals'].sum())} deal "
                   f"rows — folder: `{cfg['deals_dir']}`")
    else:
        st.caption(f"No snapshots yet — folder: `{cfg['deals_dir']}`")


if st.session_state.get("app_mode") == "Publish to database":
    _render_publish_page()
    st.stop()


# ═══════════════════════════════════════════════════════════════════════════════
# Screen 1 — Upload
# ═══════════════════════════════════════════════════════════════════════════════
if st.session_state.screen == 1:
    st.title("Upload GP Track Record File")
    st.markdown(
        "**How this works:** Upload the GP's track record spreadsheet below and the "
        "app takes it from there — it scans the workbook, finds the deal-level table, "
        "and automatically matches each column to the standard analysis template. "
        "On the next screen you'll review what it found (and fix anything that looks "
        "off), then generate and download the finished file. No manual setup needed."
    )

    uploaded = st.file_uploader(
        "GP Track Record Excel File",
        type=["xlsx", "xls"],
        help="Upload the raw portfolio workbook from the GP.",
    )

    # ── Stage 1: Profile the workbook immediately on upload ───────────
    if uploaded:
        file_id = f"{uploaded.name}_{uploaded.size}"
        if st.session_state.profile_file_id != file_id:
            raw_bytes = uploaded.read()
            uploaded.seek(0)
            with st.spinner("Profiling workbook structure…"):
                profile = profile_workbook(raw_bytes, filename=uploaded.name)
            st.session_state.profile          = profile
            st.session_state.profile_file_id  = file_id
            st.session_state.upload_bytes     = raw_bytes
            # Pre-fill GP name input from profiler if not yet set
            if not st.session_state.gp_name_input and profile.gp_name_candidates:
                st.session_state.gp_name_input = profile.gp_name_candidates[0]
            st.rerun()

    # GP / analyst inputs — widget uses the *real* GP name so the user can
    # correct it; anonymised versions are derived after parsing.
    col1, col2 = st.columns(2)
    with col1:
        gp_name = st.text_input(
            "GP / Sponsor Name",
            value=st.session_state.gp_name_input or "",
            placeholder="e.g. Manager, Sponsor, Firm",
        )
    with col2:
        analyst_name = st.text_input(
            "Analyst Name",
            value=st.session_state.analyst_name or "",
            placeholder="e.g. Jane Smith",
        )

    # ── Show profile (if available) ───────────────────────────────────
    profile: WorkbookProfile | None = st.session_state.profile
    if profile:
        _display_workbook_profile(profile)

    # ── Layout detection + sheet override ────────────────────────────
    # The pipeline auto-detects whether the workbook is CONSOLIDATED (one sheet
    # holds every deal) or PER-FUND (deals sharded across same-schema tabs that
    # must be combined). Both cases are shown here and can be overridden.
    target_table_override: "TableCandidate | None" = None
    combine_tables_override: "list[TableCandidate] | None" = None
    if profile:
        all_sheet_names = [sp.name for sp in profile.sheet_profiles]
        # Best deal-like candidate per sheet — used to resolve names → tables.
        name_to_table: dict[str, TableCandidate] = {}
        for sp in profile.sheet_profiles:
            deal_c = [t for t in sp.candidate_tables if t.likely_deal_table]
            if deal_c:
                name_to_table[sp.name] = max(deal_c, key=lambda t: t.score)

        group = group_candidate_tables(profile)

        if group and group.layout == "per_fund":
            st.info(
                f"**Per-fund layout detected** — deals are split across "
                f"{len(group.tables)} same-schema tabs. They will be combined into "
                "one deal set (duplicates removed automatically)."
            )
            with st.expander("Review tabs to combine", expanded=True):
                st.caption(
                    "These tabs share the same column layout. Untick any that aren't "
                    "deal data, or add others. The fund is taken from each tab's name."
                )
                picked = st.multiselect(
                    "Fund tabs to combine",
                    options=all_sheet_names,
                    default=group.sheet_names(),
                )
                combine_tables_override = [
                    name_to_table[n] for n in picked if n in name_to_table
                ] or None
        else:
            primary_sheet = (
                group.tables[0].sheet_name if (group and group.tables)
                else (profile.best_sheet() or "")
            )
            with st.expander("Override sheet detection", expanded=False):
                st.caption(
                    f"**Consolidated layout** — the profiler selected **{primary_sheet}** "
                    "as the single sheet holding every deal. Change this only if the "
                    "detection is wrong."
                )
                choice = st.selectbox(
                    "Deal sheet to parse",
                    options=["(use profiler selection)"] + all_sheet_names,
                    index=0,
                    help="Pick the single tab that holds all deals in one table.",
                )
                if choice != "(use profiler selection)":
                    target_table_override = name_to_table.get(choice) or next(
                        (sp.candidate_tables[0] for sp in profile.sheet_profiles
                         if sp.name == choice and sp.candidate_tables),
                        None,
                    )

    # ── Parse button — runs Stages 2 (extraction) + 3 (schema inference) ─
    parse_disabled = not uploaded or not gp_name or profile is None
    if st.button("Parse File →", type="primary", disabled=parse_disabled):
        with st.spinner("Parsing and inferring schema (may take ~30 s on first run)…"):
            raw_bytes = st.session_state.upload_bytes or b""
            if not raw_bytes and uploaded:
                raw_bytes = uploaded.read()

            try:
                result = GPParserPipeline(use_semantic=True).run(
                    raw_bytes,
                    filename=uploaded.name if uploaded else "upload.xlsx",
                    target_table=target_table_override,
                    combine_tables=combine_tables_override,
                )
            except Exception as e:
                st.error(f"Failed to parse file: {e}")
                st.stop()

        # Store pipeline result (Stages 1–3 + 5 + 6)
        st.session_state.pipeline_result = result

        # Backwards-compatible mapping dict so Screen 3 still works unchanged
        st.session_state.mapping = {
            "field_to_col":   result.schema.field_to_col,
            "col_confidence": result.schema.col_confidence,
            "source":         "multi-signal inferencer",
        }

        st.session_state.raw_df = result.table.df
        st.session_state.meta   = {
            "raw_sheet_name": result.table.sheet_name,
            # As-of date the GP reports (labelled cell > filename > first-seen).
            "report_date":    detect_track_record_date(
                raw_bytes, uploaded.name if uploaded else "",
                fallback=result.profile.report_date),
            "currency":       result.profile.currency or "USD",
            # File-level monetary-unit declaration ("$ in thousands" banner)
            "unit_banner":    detect_unit_banner(raw_bytes),
        }
        st.session_state.gp_name     = gp_name
        st.session_state.analyst_name = analyst_name

        # Fund detection uses the inferred "fund" column
        fund_col  = result.schema.field_to_col.get("fund")
        df        = result.table.df
        all_funds = (
            detect_funds_in_df(df, fund_col)
            if fund_col and fund_col in df.columns
            else []
        )
        st.session_state.all_funds = all_funds
        st.rerun()

    # Show preview after parse
    if st.session_state.raw_df is not None:
        st.success(f"Parsed **{len(st.session_state.raw_df)}** rows from sheet "
                   f"**{st.session_state.meta.get('raw_sheet_name')}**")

        all_funds = st.session_state.all_funds
        if all_funds:
            st.caption(f"**{len(all_funds)} fund(s) detected — all will be included:** "
                       + ", ".join(all_funds))

        st.dataframe(
            st.session_state.raw_df.head(5),
            use_container_width=True,
            height=200,
        )

        if st.button("Next: Review Mapping →", type="primary"):
            st.session_state.screen = 2
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# Screen 2 — Mapping Review
# ═══════════════════════════════════════════════════════════════════════════════
elif st.session_state.screen == 2:
    st.title("Column Mapping Review")

    df   = st.session_state.raw_df
    meta = st.session_state.meta
    gp   = st.session_state.gp_name

    if df is None:
        st.error("No data loaded. Please go back to Screen 1.")
        if st.button("← Back"):
            st.session_state.screen = 1
            st.rerun()
        st.stop()

    # ── Mapping comes from pipeline (already run at parse time) ───────
    mapping_result   = st.session_state.mapping
    pipeline_result  = st.session_state.pipeline_result  # may be None for legacy sessions

    # Legacy fallback: if this session predates the pipeline, run old mapper
    if mapping_result is None:
        with st.spinner("Running field detection (this may take ~30 s on first run)…"):
            mapping_result = build_mapping(
                df, gp_name=gp, use_semantic=True, configs_dir=None,
            )
        st.session_state.mapping = mapping_result

    field_to_col   = mapping_result["field_to_col"]
    col_confidence = mapping_result["col_confidence"]   # {col_name: confidence}
    source         = mapping_result.get("source", "field inference")

    # Per-field confidence (keyed by field_id) — from pipeline when available
    field_confidences: dict[str, float] = {}
    field_explanations: dict[str, list[str]] = {}
    if pipeline_result is not None:
        field_confidences  = pipeline_result.schema.confidences
        field_explanations = pipeline_result.schema.explanations

    # Summary banner
    n_auto    = len([fid for fid, c in field_confidences.items() if c >= CONFIDENCE_AUTO])
    n_review  = len([fid for fid, c in field_confidences.items() if CONFIDENCE_REVIEW <= c < CONFIDENCE_AUTO])
    n_unmapped = len([fid for fid, c in field_confidences.items() if c < CONFIDENCE_REVIEW])

    if pipeline_result:
        report = pipeline_result.report
        st.info(
            f"**{source}** · "
            f"✅ {report.n_confirmed} auto-confirmed · "
            f"🟡 {report.n_review} need review · "
            f"🔴 {report.n_unmapped_fields} unmapped · "
            f"⚠️ {report.n_issues} validation issue(s)"
        )
    else:
        st.info(f"Mapping detected via **{source}** — review flagged fields below.")

    col_options = ["(unmapped)"] + [c for c in df.columns if str(c).strip()]

    # ── Split into confirmed / needs review / optional-unmapped ────────
    # Unmapped OPTIONAL fields (e.g. Initial Fund Equity, Board Seats) are
    # expected to be absent in many GP files — they go to a quiet collapsed
    # list instead of the prominent Needs Review section, so the analyst only
    # sees actionable items.
    needs_review_entries, confirmed_entries, optional_unmapped_entries = [], [], []
    for field in TEMPLATE_FIELDS:
        fid    = field["id"]
        mapped = field_to_col.get(fid)
        # Use field_id-keyed confidence when available, else fall back to col_name-keyed
        if field_confidences:
            conf = field_confidences.get(fid, 0.0)
        else:
            conf = col_confidence.get(mapped, 0.0) if mapped else 0.0
        tier  = confidence_tier(conf) if mapped else "red"
        entry = {
            "fid": fid, "label": field["label"],
            "mapped": mapped, "conf": conf, "tier": tier,
            "signals": field_explanations.get(fid, []),
        }
        if mapped is None and is_optional(fid):
            optional_unmapped_entries.append(entry)
        elif tier in ("yellow", "red") or mapped is None:
            needs_review_entries.append(entry)
        else:
            confirmed_entries.append(entry)

    # ── Needs-review section ──────────────────────────────────────────
    updated_map = dict(field_to_col)
    if needs_review_entries:
        st.subheader(f"Needs Review ({len(needs_review_entries)} fields)")
        for entry in needs_review_entries:
            fid, label, mapped, conf, tier = (
                entry["fid"], entry["label"], entry["mapped"], entry["conf"], entry["tier"]
            )
            signals = entry["signals"]
            badge   = "🟡" if tier == "yellow" else "🔴"

            c1, c2, c3 = st.columns([2, 3, 2])
            with c1:
                st.markdown(f"{badge} **{label}**")
                if signals:
                    st.caption("  \n".join(f"↳ {s}" for s in signals[:3]))
            with c2:
                current_idx = col_options.index(mapped) if mapped and mapped in col_options else 0
                new_col = st.selectbox(
                    f"__{fid}__",
                    options=col_options,
                    index=current_idx,
                    key=f"sel_{fid}",
                    label_visibility="collapsed",
                )
                updated_map[fid] = None if new_col == "(unmapped)" else new_col
            with c3:
                if mapped:
                    st.caption(f"Score: {conf:.0%}")
                else:
                    st.caption("unmapped")

    # ── Optional fields with no match (collapsed, non-alarming) ───────
    if optional_unmapped_entries:
        with st.expander(
            f"Optional fields not found ({len(optional_unmapped_entries)}) — safe to ignore",
            expanded=False,
        ):
            st.caption(
                "These fields aren't required for the output and weren't found in the "
                "source file. They'll be left blank (or auto-computed by the template). "
                "Map one only if your file actually contains it."
            )
            for entry in optional_unmapped_entries:
                fid = entry["fid"]
                c1, c2 = st.columns([2, 3])
                with c1:
                    st.markdown(f"⚪ {entry['label']}")
                with c2:
                    new_col = st.selectbox(
                        f"__opt_{fid}__",
                        options=col_options,
                        index=0,
                        key=f"sel_{fid}",
                        label_visibility="collapsed",
                    )
                    updated_map[fid] = None if new_col == "(unmapped)" else new_col

    st.divider()

    # ── Confirmed fields ──────────────────────────────────────────────
    with st.expander(f"✅ Auto-confirmed fields ({len(confirmed_entries)})", expanded=False):
        for entry in confirmed_entries:
            sigs_txt = f"  ·  {entry['signals'][0]}" if entry["signals"] else ""
            st.caption(
                f"**{entry['label']}** ← `{entry['mapped']}` ({entry['conf']:.0%}){sigs_txt}"
            )

    # ── Validation issues ─────────────────────────────────────────────
    if pipeline_result:
        report = pipeline_result.report
        val_errs  = report.validation_errors
        val_warns = report.validation_warnings

        if val_errs or val_warns:
            with st.expander(
                f"⚠️ Validation Issues ({len(val_errs)} error(s), {len(val_warns)} warning(s))",
                expanded=bool(val_errs),
            ):
                if val_errs:
                    st.markdown("**Errors** (likely data problems):")
                    for issue in val_errs:
                        st.error(issue.label(), icon="🚨")
                if val_warns:
                    st.markdown("**Warnings** (verify these rows):")
                    for issue in val_warns[:20]:   # cap at 20 to avoid flooding
                        st.warning(issue.label(), icon="⚠️")
                    if len(val_warns) > 20:
                        st.caption(f"… and {len(val_warns) - 20} more warnings.")

    # ── Unmapped source columns ───────────────────────────────────────
    if pipeline_result and pipeline_result.schema.unmapped_cols:
        with st.expander(f"Unmapped source columns ({len(pipeline_result.schema.unmapped_cols)})", expanded=False):
            st.caption(
                "These columns exist in the source file but could not be mapped to any "
                "standardised field.  They will not appear in the output template."
            )
            for col in pipeline_result.schema.unmapped_cols:
                st.caption(f"  · `{col}`")

    st.divider()

    col_back, col_gen = st.columns(2)
    with col_back:
        if st.button("← Back"):
            st.session_state.screen = 1
            st.rerun()
    with col_gen:
        if st.button("Generate Analysis →", type="primary"):
            st.session_state.mapping["field_to_col"] = updated_map
            st.session_state.screen = 3
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# Screen 3 — Output & Download
# ═══════════════════════════════════════════════════════════════════════════════
elif st.session_state.screen == 3:
    st.title("Output & Download")

    df             = st.session_state.raw_df
    meta           = st.session_state.meta
    gp             = st.session_state.gp_name
    fund_name      = st.session_state.fund_name
    field_to_col   = st.session_state.mapping["field_to_col"]
    selected_funds = st.session_state.all_funds   # every detected fund

    if df is None:
        st.error("No data. Go back to Screen 1.")
        st.stop()

    # Collect all non-fatal errors here; shown in a summary at the end.
    phase_errors: list[dict] = []   # {"phase": ..., "detail": ..., "row": ...}

    # ── Phase 1: Transform rows ───────────────────────────────────────
    with st.spinner("Transforming deal rows…"):
        fund_col = field_to_col.get("fund")
        date_col = field_to_col.get("entry_date")

        df_filtered = df.copy()               # all funds are included

        fund_vintage_map = {}
        if fund_col and date_col and date_col in df_filtered.columns:
            try:
                fund_vintage_map = compute_fund_vintages(df_filtered, fund_col, date_col)
            except Exception as e:
                phase_errors.append({"phase": "Vintage detection", "detail": str(e), "row": ""})

        # Normalise monetary units to millions (template convention). Raw GP
        # files may report in absolute currency units or thousands instead.
        scale_map = detect_monetary_scale(
            df_filtered, field_to_col,
            file_unit_hint=meta.get("unit_banner"),
        )

        all_records = []
        for i, (_, row) in enumerate(df_filtered.iterrows()):
            try:
                rec = transform_row(
                    raw=row.to_dict(),
                    field_map=field_to_col,
                    fund_vintage_map=fund_vintage_map,
                    scale_map=scale_map,
                )
                all_records.append(rec)
            except Exception as e:
                company = str(row.get(field_to_col.get("company", ""), f"row {i+1}"))
                phase_errors.append({
                    "phase": "Row transform",
                    "detail": str(e),
                    "row": company,
                })

        included_records, excluded_records = flag_excluded_deals(all_records)

    if scale_map:
        _label = {f["id"]: f["label"] for f in TEMPLATE_FIELDS}
        for factor, note in ((1e-6, "absolute currency units — divided by 1,000,000"),
                             (1e-3, "thousands — divided by 1,000")):
            hit = [_label.get(f, f) for f, m in scale_map.items() if m == factor]
            if hit:
                st.info(f"Unit normalisation: {len(hit)} monetary column(s) "
                        f"reported in {note} to convert to millions: "
                        f"{', '.join(sorted(hit))}")

    report_date = meta.get("report_date")

    # ── Phase 2: Build mapping log ───────────────────────────────────
    mapping_log = []
    for field in TEMPLATE_FIELDS:
        fid  = field["id"]
        col  = field_to_col.get(fid)
        conf = st.session_state.mapping["col_confidence"].get(col, 1.0 if col else 0.0)
        tier = confidence_tier(conf) if col else "red"
        mapping_log.append({
            "field_id":    fid,
            "field_label": field["label"],
            "source_col":  col,
            "confidence":  conf,
            "tier":        tier,
        })

    # ── Phase 4: Build the Deal Level Input workbook ──────────────────
    # The hand-off file for the VBA analyzer: TR-Analyzer.xlsm imports it
    # and builds the analysis tabs. (The all-Python output builder still
    # lives in build_output.py; it is no longer wired into this screen.)
    gp_clean = re.sub(r"[\\/:*?\"<>|]", "-", (gp or "GP")).strip() or "GP"
    date_tag = date.today().strftime("%d-%b-%y")
    inputs_name = f"[{date_tag} - {gp_clean}] - Gross Deal Level Input.xlsx"

    inputs_bytes = None
    with st.spinner("Building the Deal Level Input workbook…"):
        try:
            inputs_bytes = build_inputs_workbook(
                included_records, gp,
                currency=meta.get("currency", "USD"),
                # Track Record Date = as-of date reported in the GP file;
                # the filename uses TODAY (processing date) instead.
                track_record_date=report_date,
            )
        except Exception as e:
            phase_errors.append({"phase": "Inputs workbook", "detail": str(e), "row": ""})

    if inputs_bytes:
        st.success(f"✅ Deal Level Input built — {len(included_records)} deals, "
                   "ready for the VBA analyzer.")
    else:
        st.error("❌ Could not build the Deal Level Input file. See the error log below.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Deals",  len(included_records))
    c2.metric("Realized",     sum(1 for r in included_records if str(r.get(5, "")).lower() == "realized"))
    c3.metric("Excluded",     len(excluded_records))
    c4.metric("Funds",        len(selected_funds))

    if inputs_bytes:
        st.download_button(
            label     = "⬇️ Download Deal Level Input",
            data      = inputs_bytes,
            file_name = inputs_name,
            mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type      = "primary",
        )
        try:
            out_path = OUTPUTS_DIR / inputs_name
            out_path.write_bytes(inputs_bytes)
            st.caption(f"A copy was saved to: `{out_path}`")
        except Exception as e:
            phase_errors.append({"phase": "Save to disk", "detail": str(e), "row": ""})

        st.markdown("#### Next: build the analysis in TR-Analyzer.xlsm")
        st.markdown(
            "1. Open **TR-Analyzer.xlsm** — it comes with the app, in the `vba` "
            "folder (click *Enable Macros* if prompted).\n"
            "2. Press **Opt+F8** on Mac / **Alt+F8** on Windows "
            "(or Tools → Macro → Macros…).\n"
            "3. Run **`ImportInputsAndBuild`** and pick the file you just "
            "downloaded.\n"
            "4. It imports the file's data and builds the 7 analysis tabs; "
            "a confirmation dialog appears when done.\n\n"
            "*Processing another GP later? Just run the macro again with the "
            "new file — everything rebuilds, nothing to clean up first.*\n\n"
            "*Once you've verified the numbers, the same file can go into the "
            "deal database — switch to **Publish to database** in the sidebar.*"
        )

    # ── Mapping log preview ──────────────────────────────────────────
    st.subheader("Mapping Log")
    log_df = pd.DataFrame(mapping_log)[["field_label", "source_col", "confidence", "tier"]]
    log_df["confidence"] = log_df["confidence"].map(lambda x: f"{x:.0%}" if x else "")

    def _style_tier(row):
        colors = {"green": "#C6EFCE", "yellow": "#FFFF99", "red": "#FFB3B3", "": ""}
        return [f"background-color: {colors.get(row['tier'], '')}"] * len(row)

    st.dataframe(
        log_df.style.apply(_style_tier, axis=1),
        use_container_width=True,
        height=300,
    )

    # ── Error log ────────────────────────────────────────────────────
    if phase_errors:
        st.divider()
        st.subheader(f"⚠️ Error Log ({len(phase_errors)} issue{'s' if len(phase_errors) != 1 else ''})")
        st.caption("These rows/steps were skipped or fell back to blank. All other data is still included in the output.")
        err_df = pd.DataFrame(phase_errors, columns=["phase", "row", "detail"])
        err_df.columns = ["Phase", "Row / Company", "Error Detail"]
        st.dataframe(err_df, use_container_width=True)

    st.divider()
    if st.button("← Start a new analysis"):
        st.session_state.screen               = 1
        st.session_state.raw_df               = None
        st.session_state.meta                 = None
        st.session_state.mapping              = None
        st.session_state.fund_vintage_map     = {}
        st.session_state.all_funds            = []
        st.session_state.profile              = None
        st.session_state.profile_file_id      = None
        st.session_state.upload_bytes         = None
        st.session_state.pipeline_result      = None
        st.rerun()

