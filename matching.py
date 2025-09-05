# matching.py
import re
import io
import time
from collections import defaultdict

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ---------- Normalization helpers ----------
def collapse_spaces(s: str) -> str:
    return re.sub(r'\s+', ' ', s).strip()

def normalize_lp_cons(val):
    if val is None:
        return ""
    return collapse_spaces(str(val)).lower()

def normalize_fund(val):
    if val is None:
        return ""
    s = collapse_spaces(str(val))
    # remove trailing variants like " lp", "l.p.", "llc"
    s = re.sub(r'\b(l\.?p\.?|llc)\b\.?$', '', s, flags=re.IGNORECASE).strip()
    return s.lower()

def find_col_by_lower(df_cols, target_lower):
    """Return actual column name from df_cols whose stripped-lower equals target_lower"""
    for col in df_cols:
        if str(col).strip().lower() == target_lower:
            return col
    raise KeyError(target_lower)

# ---------- Core function ----------
def process_files(master_bytes: bytes, output_bytes: bytes) -> bytes:
    """
    master_bytes, output_bytes -> returns bytes of the updated output workbook (xlsx)
    """
    t_start = time.time()

    # read pandas dataframes (preserve original master fund strings)
    master_df = pd.read_excel(io.BytesIO(master_bytes), dtype=str)
    output_df = pd.read_excel(io.BytesIO(output_bytes), dtype=str)

    # find columns case-insensitively
    master_cols = list(master_df.columns)
    output_cols = list(output_df.columns)

    master_lp_col = find_col_by_lower(master_cols, "lp name")
    master_fund_col = find_col_by_lower(master_cols, "fund name")
    master_cons_col = find_col_by_lower(master_cols, "consultant")

    output_lp_col = find_col_by_lower(output_cols, "lpname")
    output_fund_col = find_col_by_lower(output_cols, "fundname")
    output_cons_col = find_col_by_lower(output_cols, "reportingconsultant")

    # normalize master values and keep original fund strings
    master_lp_norm = master_df[master_lp_col].fillna("").apply(normalize_lp_cons)
    master_fund_norm = master_df[master_fund_col].fillna("").apply(normalize_fund)
    master_cons_norm = master_df[master_cons_col].fillna("").apply(normalize_lp_cons)
    master_fund_orig = master_df[master_fund_col].fillna("").astype(str)

    # build lookups (normalized)
    exact_map = {}  # (lp_norm, fund_norm, cons_norm) -> original fund string
    funds_by_lp_cons = defaultdict(list)  # (lp_norm, cons_norm) -> list of (fund_norm, fund_orig)
    for lp_n, fund_n, cons_n, fund_orig in zip(master_lp_norm, master_fund_norm, master_cons_norm, master_fund_orig):
        exact_map[(lp_n, fund_n, cons_n)] = fund_orig
        funds_by_lp_cons[(lp_n, cons_n)].append((fund_n, fund_orig))

    # normalize output columns for matching
    output_lp_norm = output_df[output_lp_col].fillna("").apply(normalize_lp_cons).tolist()
    output_fund_norm = output_df[output_fund_col].fillna("").apply(normalize_fund).tolist()
    output_cons_norm = output_df[output_cons_col].fillna("").apply(normalize_lp_cons).tolist()

    # Open output workbook via openpyxl (from bytes)
    out_wb = load_workbook(io.BytesIO(output_bytes))
    out_ws = out_wb.active

    # highlight colors
    GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    YELLOW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # insert masterentity column header at end
    last_col_index = (out_ws.max_column or 0) + 1
    out_ws.cell(row=1, column=last_col_index).value = "masterentity"

    exact_count = partial_count = no_match_count = 0
    log_lines = []
    total = len(output_df)
    t0 = time.time()

    # iterate rows
    for i in range(total):
        excel_row = i + 2  # assuming header row 1
        lp = output_lp_norm[i]
        fund = output_fund_norm[i]
        cons = output_cons_norm[i]

        matched_original_fund = None
        fill_color = None

        # exact match
        if (lp, fund, cons) in exact_map:
            matched_original_fund = exact_map[(lp, fund, cons)]
            fill_color = GREEN
            exact_count += 1
            log_lines.append(f"Row {excel_row} | Exact -> Master fund: {matched_original_fund}")
        else:
            # partial
            candidates = funds_by_lp_cons.get((lp, cons), [])
            for cand_norm, cand_orig in candidates:
                if fund and cand_norm and (fund in cand_norm or cand_norm in fund):
                    matched_original_fund = cand_orig
                    fill_color = YELLOW
                    partial_count += 1
                    log_lines.append(f"Row {excel_row} | Partial -> Master fund: {matched_original_fund}")
                    break

        # write masterentity or No Match
        if matched_original_fund:
            out_ws.cell(row=excel_row, column=last_col_index).value = matched_original_fund
        else:
            out_ws.cell(row=excel_row, column=last_col_index).value = "No Match"
            no_match_count += 1
            log_lines.append(f"Row {excel_row} | No Match")

        # highlight entire row if matched (fill_color set)
        if fill_color:
            # highlight columns 1..last_col_index (include masterentity cell too)
            for col in range(1, last_col_index + 1):
                out_ws.cell(row=excel_row, column=col).fill = fill_color

        # progress sampling (not printed to Streamlit console here)
        if (i + 1) % 100 == 0 or (i + 1) == total:
            elapsed = time.time() - t0
            # simple progress log appended
            log_lines.append(f"Processed {i+1}/{total} rows (elapsed {elapsed:.1f}s)")

    # save workbook to bytes
    out_bytes_io = io.BytesIO()
    out_wb.save(out_bytes_io)
    out_bytes = out_bytes_io.getvalue()

    # final log summary appended
    total_time = time.time() - t_start
    log_lines.append(f"Done in {total_time:.1f}s — exact={exact_count}, partial={partial_count}, no-match={no_match_count}")

    # return bytes and small summary as tuple (we'll return bytes and keep log externally in app)
    return out_bytes, {
        "exact": exact_count,
        "partial": partial_count,
        "nomatch": no_match_count,
        "log_lines": log_lines
    }
