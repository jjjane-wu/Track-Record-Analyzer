#!/usr/bin/env python3
"""Generate vba/modSpec.bas from the Python single-source-of-truth specs.

Emits the Deal List column schema (headers / formula specs / tag row), the
Deal Level Inputs column order, the Deal List header block (bucket threshold
tables etc.), and the Return & Loss Ratios pivot list as VBA data functions.

Re-run after changing deal_list_spec.py or build_output.py:
    python3 vba/generate_vba_spec.py
"""
import sys
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "app"))

from deal_list_spec import DL_COLS, DL_FORMATS, HEADER_BLOCK, TAG_ROW   # noqa: E402
from build_output import (INPUT_COLS, PIVOT_SPECS, _PAGE_FIELD_HEADERS,  # noqa: E402
                          _CANONICAL_ORDER)


def vstr(s: str) -> str:
    """Python string -> VBA string literal (quotes doubled, newlines via Chr)."""
    if s is None:
        return '""'
    parts = s.replace("\r", "").split("\n")
    lits = ['"' + p.replace('"', '""') + '"' for p in parts]
    out = ' & Chr(10) & '.join(lits)
    for ch in s:
        if ord(ch) > 126 and ch != "\n":
            raise SystemExit(f"non-ASCII char {ch!r} in spec string {s[:60]!r} — "
                             "extend vstr() with a Chr() mapping first")
    return out


lines = []
w = lines.append
w('Attribute VB_Name = "modSpec"')
w("' ===================================================================")
w("'  GENERATED FILE -- do not edit by hand.")
w("'  Source of truth: app/deal_list_spec.py + app/build_output.py")
w("'  Regenerate with:  python3 vba/generate_vba_spec.py")
w("' ===================================================================")
w("Option Explicit")
w("")
w(f"Public Const DL_NCOLS As Long = {len(DL_COLS)}")
w(f"Public Const IN_NCOLS As Long = {len(INPUT_COLS)}")
w(f"Public Const RLR_NPIVOTS As Long = {len(PIVOT_SPECS)}")
w("")

# -- Deal List columns: header + formula spec (same in:/in0:/F:/FT: prefixes
#    the Python builder uses) + number format -----------------------------
w("Public Sub LoadDealListSpec(h() As String, f() As String, fmt() As String, tag() As String)")
w(f"    ReDim h(1 To {len(DL_COLS)}): ReDim f(1 To {len(DL_COLS)})")
w(f"    ReDim fmt(1 To {len(DL_COLS)}): ReDim tag(1 To {len(DL_COLS)})")
tags = {L: t for L, t in TAG_ROW}
from openpyxl.utils import get_column_letter  # noqa: E402
for i, (hdr, spec, _key, _kind) in enumerate(DL_COLS):
    n = i + 1
    letter = get_column_letter(2 + i)
    w(f"    h({n}) = {vstr(hdr)}")
    w(f"    f({n}) = {vstr(spec or '')}")
    if i in DL_FORMATS:
        w(f"    fmt({n}) = {vstr(DL_FORMATS[i])}")
    if letter in tags:
        w(f"    tag({n}) = {vstr(tags[letter])}")
w("End Sub")
w("")

# -- Deal Level Inputs column order (for in:-link letter resolution) ------
w("Public Sub LoadInputCols(h() As String)")
w(f"    ReDim h(1 To {len(INPUT_COLS)})")
for i, (hdr, _key) in enumerate(INPUT_COLS):
    w(f"    h({i + 1}) = {vstr(hdr)}")
w("End Sub")
w("")

# -- Deal List header block (rows 1-12: bucket threshold tables, counters).
#    {LAST} resolves to the last data row at build time. ------------------
w("Public Sub LoadHeaderBlock(refs() As String, vals() As String, isNum() As Boolean, numv() As Double)")
w(f"    ReDim refs(1 To {len(HEADER_BLOCK)}): ReDim vals(1 To {len(HEADER_BLOCK)})")
w(f"    ReDim isNum(1 To {len(HEADER_BLOCK)}): ReDim numv(1 To {len(HEADER_BLOCK)})")
for i, cell in enumerate(HEADER_BLOCK):
    n = i + 1
    v = cell["v"]
    w(f"    refs({n}) = {vstr(cell['ref'])}")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        w(f"    isNum({n}) = True: numv({n}) = {v}")
    else:
        w(f"    vals({n}) = {vstr(str(v))}")
w("End Sub")
w("")

# -- Return & Loss Ratios pivot list --------------------------------------
w("Public Sub LoadPivotSpecs(title() As String, fld() As String, variant_() As String)")
w(f"    ReDim title(1 To {len(PIVOT_SPECS)}): ReDim fld(1 To {len(PIVOT_SPECS)})")
w(f"    ReDim variant_(1 To {len(PIVOT_SPECS)})")
for i, (title, fld, variant) in enumerate(PIVOT_SPECS):
    n = i + 1
    w(f"    title({n}) = {vstr(title)}")
    w(f"    fld({n}) = {vstr(fld)}")
    w(f"    variant_({n}) = {vstr(variant)}")
w("End Sub")
w("")

w("Public Sub LoadPageFields(p() As String)")
w(f"    ReDim p(1 To {len(_PAGE_FIELD_HEADERS)})")
for i, ph in enumerate(_PAGE_FIELD_HEADERS):
    w(f"    p({i + 1}) = {vstr(ph)}")
w("End Sub")
w("")

# -- canonical item order for bucket-type axis fields ---------------------
w("' Canonical display order for bucket axes (native pivots would otherwise")
w("' sort alphabetically). Returns Empty when the field has no fixed order.")
w("Public Function CanonicalOrder(ByVal fieldName As String) As Variant")
w("    Select Case fieldName")
for hdr, labels in _CANONICAL_ORDER.items():
    w(f"        Case {vstr(hdr)}")
    parts = ", ".join(vstr(l) for l in labels)
    w(f"            CanonicalOrder = Array({parts})")
w("        Case Else")
w("            CanonicalOrder = Empty")
w("    End Select")
w("End Function")

out = ROOT / "vba" / "modSpec.bas"
out.write_text("\r\n".join(lines) + "\r\n", encoding="ascii")
print(f"wrote {out} ({len(lines)} lines)")
