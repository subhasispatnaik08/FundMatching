# matching.py
import re
import time
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from collections import defaultdict

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
    s = re.sub(r'\b(l\.?p\.?|llc)\b\.?$', '', s, flags=re.IGNORECASE).strip()
    return s.lower()

def find_column_ignore_case(df, target_lower):
    for col in df.columns:
        if str(col).strip().lower() == target_lower:
            return col
    raise KeyError(f"Column '{target_lower}' not found.")

def read_uploaded(uploaded_bytes: bytes, filename: str):
    """Read uploaded file (CSV or Excel) into a DataFrame"""
    buffer = BytesIO(uploaded_bytes)
    if filename.lower().endswith(".csv"):
        try:
            return pd.read_csv(buffer, dtype=str, encoding="utf-8")
        except UnicodeDecodeError:
            buffer.seek(0)  # reset pointer
            return pd.read_csv(buffer, dtype=str, encoding="latin1")
    else:
        return pd.read_excel(buffer, dtype=str)

def process_files(master_bytes: bytes, output_bytes: bytes, 
                  master_filename: str = "master.xlsx", 
                  output_filename: str = "output.xlsx"):
    """
    Run fund matching. 
    Inputs: file bytes + filenames (to detect CSV vs Excel).
    Returns: (BytesIO result_xlsx, stats dict)
    """

    # ---------- Read DataFrames ----------
    master_orig = read_uploaded(master_bytes, master_filename)
    output_orig = read_uploaded(output_bytes, output_filename)

    # ---------- Locate required columns ----------
    master_lp_col = find_column_ignore_case(master_orig, "lp name")
    master_fund_col = find_column_ignore_case(master_orig, "fund name")
    master_cons_col = find_column_ignore_case(master_orig, "consultant")

    output_lp_col = find_column_ignore_case(output_orig, "lpname")
    output_fund_col = find_column_ignore_case(output_orig, "fundname")
    output_cons_col = find_column_ignore_case(output_orig, "reportingconsultant")

    # ---------- Build normalized master structures ----------
    master_lp_norm = master_orig[master_lp_col].fillna("").apply(normalize_lp_cons)
    master_fund_norm = master_orig[master_fund_col].fillna("").apply(normalize_fund)
    master_cons_norm = master_orig[master_cons_col].fillna("").apply(normalize_lp_cons)
    master_fund_orig = master_orig[master_fund_col].fillna("").astype(str)

    exact_map = {}
    funds_by_lp_cons = defaultdict(list)
    for lp_n, fund_n, cons_n, fund_orig in zip(master_lp_norm, master_fund_norm, master_cons_norm, master_fund_orig):
        exact_map[(lp_n, fund_n, cons_n)] = fund_orig
        funds_by_lp_cons[(lp_n, cons_n)].append((fund_n, fund_orig))

    # ---------- Normalize output ----------
    output_lp_norm = output_orig[output_lp_col].fillna("").apply(normalize_lp_cons).tolist()
    output_fund_norm = output_orig[output_fund_col].fillna("").apply(normalize_fund).tolist()
    output_cons_norm = output_orig[output_cons_col].fillna("").apply(normalize_lp_cons).tolist()

    # ---------- Create workbook ----------
    tmp_buffer = BytesIO()
    output_orig.to_excel(tmp_buffer, index=False)
    tmp_buffer.seek(0)
    wb = load_workbook(tmp_buffer)
    ws = wb.active

    GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    YELLOW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    last_col_index = ws.max_column + 1
    ws.cell(row=1, column=last_col_index).value = "masterentity"

    exact_count = 0
    partial_count = 0
    no_match_count = 0
    log_lines = []

    total = len(output_orig)
    t0 = time.time()

    for i in range(total):
        excel_row = i + 2
        lp = output_lp_norm[i]
        fund = output_fund_norm[i]
        cons = output_cons_norm[i]

        matched_original_fund = None
        fill_color = None

        if (lp, fund, cons) in exact_map:
            matched_original_fund = exact_map[(lp, fund, cons)]
            fill_color = GREEN
            exact_count += 1
            log_lines.append(f"Row {excel_row} | Exact -> {matched_original_fund}")
        else:
            candidates = funds_by_lp_cons.get((lp, cons), [])
            for cand_norm, cand_orig in candidates:
                if fund and cand_norm and (fund in cand_norm or cand_norm in fund):
                    matched_original_fund = cand_orig
                    fill_color = YELLOW
                    partial_count += 1
                    log_lines.append(f"Row {excel_row} | Partial -> {matched_original_fund}")
                    break

        if matched_original_fund:
            ws.cell(row=excel_row, column=last_col_index).value = matched_original_fund
        else:
            ws.cell(row=excel_row, column=last_col_index).value = "No Match"
            no_match_count += 1
            log_lines.append(f"Row {excel_row} | No Match")

        if fill_color:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=excel_row, column=col).fill = fill_color

    elapsed = time.time() - t0

    # ---------- Save workbook into BytesIO ----------
    result_buffer = BytesIO()
    wb.save(result_buffer)
    result_buffer.seek(0)

    # ---------- Stats ----------
    stats = {
        "exact": exact_count,
        "partial": partial_count,
        "nomatch": no_match_count,
        "rows": total,
        "elapsed": elapsed,
        "log_lines": log_lines,
    }

    return result_buffer, stats

