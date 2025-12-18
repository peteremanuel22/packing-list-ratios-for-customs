
# --- Optional runtime bootstrap (fallback only) ---
# Tries to install openpyxl at runtime if it's missing.
# Not recommended for Streamlit Cloud, but handy for local runs without requirements setup.

def _ensure_openpyxl():
    try:
        import openpyxl  # noqa: F401
        return
    except Exception:
        pass

    import sys, subprocess
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--no-cache-dir", "openpyxl>=3.1"])
        import openpyxl  # noqa: F401
    except Exception as e:
        import streamlit as st
        st.error(
            "Failed to install **openpyxl** at runtime. "
            "Please ensure `requirements.txt` includes `openpyxl>=3.1` and redeploy."
        )
        st.stop()

_ensure_openpyxl()
# --- End fallback ---




# app.py
# -*- coding: utf-8 -*-
"""
Packing List Recalculator (Gas Cookers)
- Multi-sheet Excel input (each sheet = a container).
- Detect multiple tables by header row:
  S.N | Box code | component in arabic | component in E | Codes | Qu. | Box
- Preserve structure & styles; only update 'Qu.' using per-cooker ratios.
- Safe READ and WRITE for merged cells (read/write the master—top-left—cell).
- Only updates rows whose original 'Qu.' is numeric; text/unit rows remain as-is.
- Preview shows the ARABIC component name.
- GROUPING BY CODE ONLY. If Codes is empty, treat the row as unique (no cross-row merging).
- Robust error handling (no blank page; errors are shown in the UI).
"""

import io
import re
import traceback
from typing import Dict, List, Tuple, Any, Optional

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

# -------------------------
# UI setup
# -------------------------
st.set_page_config(page_title="Packing List Recalculator", layout="wide")
st.title("Packing List Recalculator (Gas Cookers)")
st.write("Upload your Excel packing list (one or more worksheets).")
st.write("Enter the original cookers count represented by the file and the desired cookers count.")
st.write("The app will adjust the 'Qu.' column proportionally and preserve the original structure and styles.")

# -------------------------
# Constants & helpers
# -------------------------
EXPECTED_HEADERS = {
    "S.N", "Box code", "component in arabic", "component in E", "Codes", "Qu.", "Box"
}

PACKAGING_PATTERNS = [
    r"foam", r"فوم", r"carton", r"كرتون", r"كرتونة", r"زاوية فوم", r"قاعدة فوم", r"شريحة فوم"
]
packaging_re = re.compile("|".join(PACKAGING_PATTERNS), flags=re.IGNORECASE)

def norm(x: Any) -> str:
    """Normalize a cell value to trimmed string."""
    if x is None:
        return ""
    return str(x).strip()

def parse_quantity(q_raw: Any) -> Tuple[bool, int]:
    """
    Extract a numeric quantity from a cell value.
    Returns (is_numeric, int_value).
    - Accepts numbers stored as text/float ('120', '120.0').
    - If text contains a number with words ('4 بالته'), extracts the first number.
    - If nothing numeric is found, returns (False, 0).
    """
    s = norm(q_raw)
    if s == "":
        return (False, 0)
    try:
        val = float(s)
        return (True, int(round(val)))
    except Exception:
        pass
    m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
    if m:
        try:
            val = float(m.group(0))
            return (True, int(round(val)))
        except Exception:
            return (False, 0)
    return (False, 0)

def is_header_row(values: List[Any]) -> bool:
    """Treat a row as header if essential subset appears."""
    vals = {norm(v) for v in values if norm(v)}
    essential = {"S.N", "Codes", "Qu."}
    return essential.issubset(vals)

# -------------------------
# merged-cell helpers
# -------------------------
def find_merged_master(ws: Worksheet, row: int, col: int) -> Optional[Tuple[int, int]]:
    """
    If (row, col) is inside a merged range, return (min_row, min_col) of that range.
    Otherwise return None.
    """
    coord = ws.cell(row=row, column=col).coordinate
    for rng in ws.merged_cells.ranges:
        if coord in rng:
            return (rng.min_row, rng.min_col)
    return None

def safe_read(ws: Worksheet, row: int, col: int) -> Any:
    """
    Read value from cell (row, col); if it's part of a merged range, return the master (top-left) value.
    """
    master = find_merged_master(ws, row, col)
    if master is not None:
        return ws.cell(row=master[0], column=master[1]).value
    return ws.cell(row=row, column=col).value

def safe_write(ws: Worksheet, row: int, col: int, value: int) -> None:
    """
    Write to cell (row, col); if it's a merged child cell, redirect to the master (top-left) cell.
    """
    master = find_merged_master(ws, row, col)
    target_row, target_col = (master if master is not None else (row, col))
    ws.cell(row=target_row, column=target_col).value = int(value)

# -------------------------
# table detection
# -------------------------
def find_tables(ws: Worksheet) -> List[Dict[str, Any]]:
    """
    Scan a worksheet to find tables by header rows.
    Return list of dicts: {header_row, col_map, data_rows}
    """
    tables = []
    max_row = ws.max_row
    max_col = ws.max_column

    row_idx = 1
    while row_idx <= max_row:
        row_vals = [safe_read(ws, row_idx, c) for c in range(1, max_col + 1)]
        if is_header_row(row_vals):
            col_map = {}
            for c in range(1, max_col + 1):
                h = norm(safe_read(ws, row_idx, c))
                if h in EXPECTED_HEADERS:
                    col_map[h] = c
            data_rows = []
            r = row_idx + 1
            while r <= max_row:
                candidate = [safe_read(ws, r, c) for c in range(1, max_col + 1)]
                if is_header_row(candidate):
                    break
                data_rows.append(r)
                r += 1
            tables.append({"header_row": row_idx, "col_map": col_map, "data_rows": data_rows})
            row_idx = r
        else:
            row_idx += 1

    return tables

# -------------------------
# component & occurrences
# -------------------------
def component_id_from_row(ws: Worksheet, row: int, col_map: Dict[str, int]) -> str:
    """
    Build component identifier: GROUP BY CODE ONLY.
    - If Codes is present: use it as the unique key.
    - If Codes is empty: create a per-row unique key so we don't merge unrelated "بدون" lines.
    """
    code = norm(safe_read(ws, row, col_map.get("Codes", 0)))
    if code:
        return code
    return f"__NO_CODE__@{ws.title}@{row}"

# Occurrence tuple includes ARABIC NAME:
# (ws_name, row_idx, q_col, numeric_q, is_numeric, arabic_name)
Occurrence = Tuple[str, int, int, int, bool, str]

def extract_component_occurrences(wb) -> Dict[str, List[Occurrence]]:
    """
    Returns: component_id -> list of occurrences
    Each occurrence: (worksheet_name, row_idx, q_col, numeric_q, is_numeric, arabic_name)
    """
    comp_occ: Dict[str, List[Occurrence]] = {}
    for ws in wb.worksheets:
        tables = find_tables(ws)
        for tbl in tables:
            cmap = tbl["col_map"]
            q_col = cmap.get("Qu.")
            if not q_col:
                continue
            for r in tbl["data_rows"]:
                all_vals = [safe_read(ws, r, c) for c in range(1, ws.max_column + 1)]
                if not any(norm(v) for v in all_vals):
                    continue
                q_raw = safe_read(ws, r, q_col)
                is_num, q_val = parse_quantity(q_raw)
                arabic_name = norm(safe_read(ws, r, cmap.get("component in arabic", 0)))
                cid = component_id_from_row(ws, r, cmap)
                comp_occ.setdefault(cid, []).append((ws.title, r, q_col, q_val, is_num, arabic_name))
    return comp_occ

# -------------------------
# allocation & application
# -------------------------
def largest_remainder_allocate(originals: List[int], target_total: int) -> List[int]:
    """
    Allocate target_total proportionally to originals; ensure integer sum via largest remainder.
    """
    if target_total < 0:
        target_total = 0
    n = len(originals)
    if n == 0:
        return []
    s = sum(originals)
    if s == 0:
        base = target_total // n
        rem = target_total - base * n
        alloc = [base] * n
        for i in range(rem):
            alloc[i] += 1
        return alloc
    weights = np.array(originals, dtype=float) / float(s)
    ideal = weights * float(target_total)
    floors = np.floor(ideal).astype(int)
    residual = int(target_total - floors.sum())
    fracs = ideal - floors
    order = np.argsort(-fracs)
    alloc = floors.copy()
    for i in range(residual):
        alloc[order[i]] += 1
    return alloc.tolist()

def apply_allocations(wb, comp_occ: Dict[str, List[Occurrence]], comp_targets: Dict[str, int]) -> None:
    """
    Overwrite numeric 'Qu.' cells in-place only. Styles preserved by openpyxl.
    - Non-numeric original 'Qu.' cells are left unchanged.
    - If a target total must be distributed but a component has zero numeric occurrences, nothing is written.
    """
    ws_index = {ws.title: ws for ws in wb.worksheets}
    for cid, occs in comp_occ.items():
        target_total = comp_targets.get(cid)
        if target_total is None:
            continue
        numeric_occs = [(ws, r, c, q, is_num, ar) for (ws, r, c, q, is_num, ar) in occs if is_num]
        if not numeric_occs:
            continue
        originals = [q for (_, _, _, q, _, _) in numeric_occs]
        new_vals = largest_remainder_allocate(originals, target_total)
        for new_q, (ws_name, row_idx, q_col, _, _, _) in zip(new_vals, numeric_occs):
            ws = ws_index[ws_name]
            safe_write(ws, row_idx, q_col, int(new_q))

# -------------------------
# ratios & targets
# -------------------------
def compute_ratios(comp_occ: Dict[str, List[Occurrence]], original_cookers: int) -> Dict[str, float]:
    """Per-component ratio = (sum of numeric quantities) / original cookers."""
    if original_cookers <= 0:
        return {k: 0.0 for k in comp_occ}
    ratios = {}
    for cid, occs in comp_occ.items():
        total_q_numeric = sum(q for (_, _, _, q, is_num, _) in occs if is_num)
        ratios[cid] = total_q_numeric / float(original_cookers)
    return ratios

def compute_targets_from_ratios(ratios: Dict[str, float], desired_cookers: int) -> Dict[str, int]:
    """Target totals = ratio * desired cookers, rounded to nearest int (>=0)."""
    targets = {}
    for cid, r in ratios.items():
        val = r * float(desired_cookers)
        targets[cid] = max(int(round(val)), 0)
    return targets

def build_preview_dataframe(comp_occ: Dict[str, List[Occurrence]], ratios: Dict[str, float], targets: Dict[str, int]) -> pd.DataFrame:
    """
    Build the preview DataFrame.
    Shows Arabic component name from the first occurrence within each CODE group.
    """
    rows = []
    for cid, occs in comp_occ.items():
        code_display = cid if not cid.startswith("__NO_CODE__@") else ""
        arabic_name = occs[0][5] if occs else ""
        original_total_numeric = sum(q for (_, _, _, q, is_num, _) in occs if is_num)
        rows.append({
            "Code": code_display,
            "Component (Arabic)": arabic_name,
            "Original total (numeric Qu.)": original_total_numeric,
            "Per-cooker ratio": ratios.get(cid, 0.0),
            "Target total (Qu.)": targets.get(cid, original_total_numeric),
            "Occurrences (numeric/all)": f"{sum(1 for o in occs if o[4])}/{len(occs)}"
        })
    df = pd.DataFrame(rows).sort_values(["Code", "Component (Arabic)"]).reset_index(drop=True)
    return df

def is_packaging_component(occs_for_one_component: List[Occurrence]) -> bool:
    """
    Decide packaging based on Arabic name of the first occurrence in this component group.
    """
    if not occs_for_one_component:
        return False
    arabic_name = occs_for_one_component[0][5] or ""
    return bool(packaging_re.search(arabic_name))

# -------------------------
# Inputs
# -------------------------
uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])
col_input = st.columns(2)
with col_input[0]:
    original_cookers = st.number_input("Original shipment cookers count", min_value=1, value=250, step=1)
with col_input[1]:
    desired_cookers = st.number_input("Desired cookers count", min_value=1, value=250, step=1)

exclude_packaging = st.checkbox("Exclude packaging-only items (foam/cartons) from scaling", value=False)

# -------------------------
# Main processing
# -------------------------
if uploaded is not None:
    try:
        with st.spinner("Reading workbook and computing allocations..."):
            data = uploaded.read()
            bio_in = io.BytesIO(data)

            # Read values for analysis
            wb_in = load_workbook(filename=bio_in, data_only=True)
            comp_occ = extract_component_occurrences(wb_in)
            ratios = compute_ratios(comp_occ, original_cookers)
            targets = compute_targets_from_ratios(ratios, desired_cookers)

            if exclude_packaging:
                for cid, occs in list(comp_occ.items()):
                    if is_packaging_component(occs):
                        orig_total_numeric = sum(q for (_, _, _, q, is_num, _) in occs if is_num)
                        targets[cid] = orig_total_numeric

            # Preview
            st.subheader("Preview of totals and ratios")
            df_prev = build_preview_dataframe(comp_occ, ratios, targets)
            st.dataframe(df_prev, use_container_width=True)

            # Rewrite quantities while preserving styles
            bio_in.seek(0)
            wb_out = load_workbook(filename=io.BytesIO(bio_in.read()), data_only=False)
            apply_allocations(wb_out, comp_occ, targets)

            # Save to buffer for download
            out_buf = io.BytesIO()
            wb_out.save(out_buf)
            out_buf.seek(0)

        st.success("New workbook is ready. Only numeric 'Qu.' values were changed; styles and layout remain unchanged.")
        st.download_button(
            label="Download recalculated packing list (.xlsx)",
            data=out_buf,
            file_name="packing_list_recalculated.xlsx",
            mime="application/vnd.openxmlformats" ,     
      )

    except Exception as e:
        st.error("An error occurred while processing the file.")
        st.exception(e)
        tb = traceback.format_exc()
        st.text("Traceback:")
        st.code(tb)

# ==== Centered footer ====
footer_css = """
<style>
.app-footer {
  position: fixed;
  left: 50%;
  bottom: 12px;
  transform: translateX(-50%);
  z-index: 9999;
  background: rgba(255,255,255,0.85);
  border: 1px solid #e6e6e6;
  border-radius: 14px;
  padding: 8px 14px;
  font-weight: 600;
  font-size: 14px;
}
</style>
"""
footer_html = """
<div class="app-footer">✨ تم التنفيذ بواسطة م / بيتر عمانوئيل – جميع الحقوق محفوظة © 2025 ✨</div>
"""
st.markdown(footer_css, unsafe_allow_html=True)
st.markdown(footer_html, unsafe_allow_html=True)


