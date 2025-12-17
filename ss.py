
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
- Robust error handling and no stray backticks/duplicated tokens.
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

PACKAGING_PATTERNS = [
    r"foam", r"فوم", r"carton", r"كرتون", r"كرتونة", r"زاوية فوم", r"قاعدة فوم", r"شريحة فوم"
]
packaging_re = re.compile("|".join(PACKAGING_PATTERNS), flags=re.IGNORECASE)

def norm(x: Any) -> str:
    """Normalize cell value to trimmed string."""
    if x is None:
        return ""
    return str(x).strip()

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

def extract_component_occurrences(wb) -> Dict[str, List[Tuple[str, int, int, int]]]:
    """
    Returns: component_id -> list of (worksheet_name, row_idx, q_col, original_q)
    """
    comp_occ: Dict[str, List[Tuple[str, int, int, int]]] = {}
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
                q_val = 0
                if norm(q_raw):
                    try:
                        q_val = int(float(norm(q_raw)))
                    except Exception:
                        q_val = 0
                cid = component_id_from_row(ws, r, cmap)
                comp_occ.setdefault(cid, []).append((ws.title, r, q_col, q_val))
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

def apply_allocations(wb, comp_occ: Dict[str, List[Tuple[str, int, int, int]]],
                      comp_targets: Dict[str, int]) -> None:
    """Overwrite Qu. cells in-place. Styles are preserved by openpyxl."""
    ws_index = {ws.title: ws for ws in wb.worksheets}
    for cid, occs in comp_occ.items():
        target_total = comp_targets.get(cid)
        if target_total is None:
            continue
        originals = [q for (_, _, _, q) in occs]
        new_vals = largest_remainder_allocate(originals, target_total)
        for new_q, (ws_name, row_idx, q_col, _) in zip(new_vals, occs):
            ws = ws_index[ws_name]
            ws.cell(row=row_idx, column=q_col).value = int(new_q)

def compute_ratios(comp_occ: Dict[str, List[Tuple[str, int, int, int]]], original_cookers: int) -> Dict[str, float]:
    """Per-component ratio = total quantity / original cookers."""
    if original_cookers <= 0:
        return {k: 0.0 for k in comp_occ}
    ratios = {}
    for cid, occs in comp_occ.items():
        total_q = sum(q for (_, _, _, q) in occs)
        ratios[cid] = total_q / float(original_cookers)
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

def build_preview_dataframe(comp_occ, ratios, targets):
    rows = []
    for cid, occs in comp_occ.items():
        code, name = split_cid(cid)
        original_total = sum(q for (_, _, _, q) in occs)
        rows.append({
            "Code": code,
            "Component": name,
            "Original total (Qu.)": original_total,
            "Per-cooker ratio": ratios.get(cid, 0.0),
            "Target total (Qu.)": targets.get(cid, original_total),
            "Occurrences": len(occs)
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
                        orig_total = sum(q for (_, _, _, q) in comp_occ[cid])
                        targets[cid] = orig_total

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

        st.success("New workbook is ready. Only 'Qu.' values were changed; styles and layout remain unchanged.")
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
else:
    st.info("Upload an Excel")

