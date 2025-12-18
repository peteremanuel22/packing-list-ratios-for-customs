
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
- Detect multiple tables by header row (S.N, Box code, component in arabic, component in E, Codes, Qu., Box).
- Keep structure & style; only update 'Qu.' using per-cooker ratios.
- Handles merged cells safely (writes to master cell only).
- Only updates rows whose original 'Qu.' is numeric; leaves text/non-numeric as-is.
- Robust error handling to avoid white screen.
"""

import io
import re
import traceback
from typing import Dict, List, Tuple, Any

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import MergedCell

# -------------------------
# UI setup (render immediately)
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

# Optional: detect packaging items by keywords (used only if you check the box)
PACKAGING_PATTERNS = [
    r"foam", r"فوم", r"carton", r"كرتون", r"كرتونة", r"زاوية فوم", r"قاعدة فوم", r"شريحة فوم"
]
packaging_re = re.compile("|".join(PACKAGING_PATTERNS), flags=re.IGNORECASE)

def norm(x: Any) -> str:
    """Normalize cell value to trimmed string."""
    if x is None:
        return ""
    return str(x).strip()

def parse_quantity(q_raw: Any) -> Tuple[bool, int]:
    """
    Try to extract a numeric quantity from a cell value.
    Returns (is_numeric, int_value).

    - Accepts numbers stored as text or float (e.g., '120', '120.0').
    - If the text contains a number with words (e.g., '4 بالته'), extracts the first number.
    - If nothing numeric is found, returns (False, 0).
    """
    s = norm(q_raw)
    if s == "":
        return (False, 0)
    # Direct numeric
    try:
        val = float(s)
        return (True, int(round(val)))
    except Exception:
        pass
    # Try to find first number inside text
    m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
    if m:
        try:
            val = float(m.group(0))
            return (True, int(round(val)))
        except Exception:
            return (False, 0)
    return (False, 0)

def is_header_row(values: List[Any]) -> bool:
    """Header row if essential subset appears."""
    vals = {norm(v) for v in values if norm(v)}
    essential = {"S.N", "Codes", "Qu."}
    return essential.issubset(vals)

def find_tables(ws: Worksheet) -> List[Dict[str, Any]]:
    """
    Scan worksheet to find tables by header rows.
    Return list of dicts: {header_row, col_map, data_rows}
    """
    tables = []
    max_row = ws.max_row
    max_col = ws.max_column

    row_idx = 1
    while row_idx <= max_row:
        row_vals = [ws.cell(row=row_idx, column=c).value for c in range(1, max_col + 1)]
        if is_header_row(row_vals):
            # Build column map (header -> col index)
            col_map = {}
            for c in range(1, max_col + 1):
                h = norm(ws.cell(row=row_idx, column=c).value)
                if h in EXPECTED_HEADERS:
                    col_map[h] = c

            # Collect data rows until next header or end
            data_rows = []
            r = row_idx + 1
            while r <= max_row:
                candidate = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
                if is_header_row(candidate):
                    break
                data_rows.append(r)
                r += 1

            tables.append({"header_row": row_idx, "col_map": col_map, "data_rows": data_rows})
            row_idx = r
        else:
            row_idx += 1

    return tables

def component_id_from_row(ws: Worksheet, row: int, col_map: Dict[str, int]) -> str:
    """Prefer Codes; otherwise use English/Arabic names to keep 'بدون' items distinct."""
    code = norm(ws.cell(row=row, column=col_map.get("Codes", 0)).value)
    comp_e = norm(ws.cell(row=row, column=col_map.get("component in E", 0)).value)
    comp_ar = norm(ws.cell(row=row, column=col_map.get("component in arabic", 0)).value)
    comp_name = comp_e if comp_e else comp_ar
    return f"{code}|{comp_name}"

# Occurrence tuple type:
# (ws_name, row_idx, q_col, numeric_q, is_numeric)
Occurrence = Tuple[str, int, int, int, bool]

def extract_component_occurrences(wb) -> Dict[str, List[Occurrence]]:
    """
    Returns: component_id -> list of occurrences
    Each occurrence: (worksheet_name, row_idx, q_col, numeric_q, is_numeric)
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
                # Skip rows that are entirely empty
                if all(norm(ws.cell(row=r, column=c).value) == "" for c in range(1, ws.max_column + 1)):
                    continue
                q_raw = ws.cell(row=r, column=q_col).value
                is_num, q_val = parse_quantity(q_raw)
                cid = component_id_from_row(ws, r, cmap)
                comp_occ.setdefault(cid, []).append((ws.title, r, q_col, q_val, is_num))
    return comp_occ

def largest_remainder_allocate(originals: List[int], target_total: int) -> List[int]:
    """Allocate target_total proportionally to originals; ensure integer sum via largest remainder."""
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
    order = np.argsort(-fracs)  # indices sorted by descending fractional part
    alloc = floors.copy()
    for i in range(residual):
        alloc[order[i]] += 1
    return alloc.tolist()

def safe_write(ws: Worksheet, row: int, col: int, value: int) -> None:
    """
    Write to cell (row, col); if it's a merged child cell, redirect to the master (top-left) cell.
    """
    cell = ws.cell(row=row, column=col)
    # If the cell is a merged child, find its merged range and write to the top-left cell
    if isinstance(cell, MergedCell) or cell.coordinate in ws.merged_cells:
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                master = ws.cell(row=rng.min_row, column=rng.min_col)
                master.value = int(value)
                return
    # Normal cell
    cell.value = int(value)

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

        # Split occurrences into numeric vs non-numeric
        numeric_occs = [(ws, r, c, q, is_num) for (ws, r, c, q, is_num) in occs if is_num]
        if not numeric_occs:
            # Nothing to write for this component
            continue

        originals = [q for (_, _, _, q, _) in numeric_occs]
        new_vals = largest_remainder_allocate(originals, target_total)

        for new_q, (ws_name, row_idx, q_col, _, _) in zip(new_vals, numeric_occs):
            ws = ws_index[ws_name]
            safe_write(ws, row_idx, q_col, int(new_q))

def compute_ratios(comp_occ: Dict[str, List[Occurrence]], original_cookers: int) -> Dict[str, float]:
    """Per-component ratio = (sum of numeric quantities) / original cookers."""
    if original_cookers <= 0:
        return {k: 0.0 for k in comp_occ}
    ratios = {}
    for cid, occs in comp_occ.items():
        total_q_numeric = sum(q for (_, _, _, q, is_num) in occs if is_num)
        ratios[cid] = total_q_numeric / float(original_cookers)
    return ratios

def compute_targets_from_ratios(ratios: Dict[str, float], desired_cookers: int) -> Dict[str, int]:
    """Target totals = ratio * desired cookers, rounded to nearest int (>=0)."""
    targets = {}
    for cid, r in ratios.items():
        val = r * float(desired_cookers)
        targets[cid] = max(int(round(val)), 0)
    return targets

def split_cid(cid: str) -> Tuple[str, str]:
    parts = cid.split("|", 1)
    code = parts[0]
    name = parts[1] if len(parts) > 1 else ""
    return code, name

def build_preview_dataframe(comp_occ: Dict[str, List[Occurrence]], ratios: Dict[str, float], targets: Dict[str, int]):
    rows = []
    for cid, occs in comp_occ.items():
        code, name = split_cid(cid)
        original_total_numeric = sum(q for (_, _, _, q, is_num) in occs if is_num)
        rows.append({
            "Code": code,
            "Component": name,
            "Original total (numeric Qu.)": original_total_numeric,
            "Per-cooker ratio": ratios.get(cid, 0.0),
            "Target total (Qu.)": targets.get(cid, original_total_numeric),
            "Occurrences (numeric/all)": f"{sum(1 for o in occs if o[4])}/{len(occs)}"
        })
    df = pd.DataFrame(rows).sort_values(["Code", "Component"]).reset_index(drop=True)
    return df

def is_packaging_component(cid: str) -> bool:
    _, name = split_cid(cid)
    return bool(packaging_re.search(name))

# -------------------------
# Inputs (ensures first paint)
# -------------------------
uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])
col_input = st.columns(2)
with col_input[0]:
    original_cookers = st.number_input("Original shipment cookers count", min_value=1, value=250, step=1)
with col_input[1]:
    desired_cookers = st.number_input("Desired cookers count", min_value=1, value=250, step=1)

exclude_packaging = st.checkbox("Exclude packaging-only items (foam/cartons) from scaling", value=False)

# -------------------------
# Main processing (with spinner & try/except)
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
                for cid in list(comp_occ.keys()):
                    if is_packaging_component(cid):
                        # Keep original numeric total for packaging items
                        orig_total_numeric = sum(q for (_, _, _, q, is_num) in comp_occ[cid] if is_num)
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
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
