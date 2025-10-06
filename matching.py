# fund_matcher_v4.py
# Requires: pandas, openpyxl
# pip install pandas openpyxl

import re
import os
import time
from tkinter import Tk, filedialog, messagebox

import pandas as pd
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
    # remove trailing variants like " lp", "l.p.", "llc" (word-boundary)
    s = re.sub(r'\b(l\.?p\.?|llc)\b\.?$', '', s, flags=re.IGNORECASE).strip()
    return s.lower()

def find_column_ignore_case(df, target_lower):
    """Return the actual column name in df whose stripped-lower equals target_lower.
       Raises KeyError if not found."""
    for col in df.columns:
        if str(col).strip().lower() == target_lower:
            return col
    raise KeyError(f"Column '{target_lower}' not found (checked case-insensitively).")

# ---------- File pick ----------
Tk().withdraw()
print("Select MASTER fund file (Excel or CSV, must contain headers 'LP name', 'Fund name', 'Consultant'):")
master_path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx;*.csv")])
if not master_path:
    raise SystemExit("Master file not selected. Exiting.")

print("Select OUTPUT file to highlight (Excel or CSV, must contain headers 'lpname', 'fundname', 'reportingconsultant'):")
output_path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx;*.csv")])
if not output_path:
    raise SystemExit("Output file not selected. Exiting.")

# ---------- Read originals (preserve exact master fund strings) ----------
def read_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(path, dtype=str)
    else:
        return pd.read_excel(path, dtype=str)

master_orig = read_file(master_path)
output_orig = read_file(output_path)

# ---------- Locate required columns ----------
try:
    master_lp_col = find_column_ignore_case(master_orig, "lp name")
    master_fund_col = find_column_ignore_case(master_orig, "fund name")
    master_cons_col = find_column_ignore_case(master_orig, "consultant")
except KeyError as e:
    messagebox.showerror("Missing Column", f"In Master file: {e}")
    raise SystemExit(e)

try:
    output_lp_col = find_column_ignore_case(output_orig, "lpname")
    output_fund_col = find_column_ignore_case(output_orig, "fundname")
    output_cons_col = find_column_ignore_case(output_orig, "reportingconsultant")
except KeyError as e:
    messagebox.showerror("Missing Column", f"In Output file: {e}")
    raise SystemExit(e)

# ---------- Build normalized master structures but keep original fund strings ----------
master_lp_norm = master_orig[master_lp_col].fillna("").apply(normalize_lp_cons)
master_fund_norm = master_orig[master_fund_col].fillna("").apply(normalize_fund)
master_cons_norm = master_orig[master_cons_col].fillna("").apply(normalize_lp_cons)
master_fund_orig = master_orig[master_fund_col].fillna("").astype(str)  # original form to write back

# Build lookups
exact_map = {}
funds_by_lp_cons = defaultdict(list)

for lp_n, fund_n, cons_n, fund_orig in zip(master_lp_norm, master_fund_norm, master_cons_norm, master_fund_orig):
    exact_map[(lp_n, fund_n, cons_n)] = fund_orig
    funds_by_lp_cons[(lp_n, cons_n)].append((fund_n, fund_orig))

print(f"Master loaded: {len(master_orig)} rows; groups (LP+Cons): {len(funds_by_lp_cons)}")

# ---------- Normalize output columns for matching ----------
output_lp_norm = output_orig[output_lp_col].fillna("").apply(normalize_lp_cons).tolist()
output_fund_norm = output_orig[output_fund_col].fillna("").apply(normalize_fund).tolist()
output_cons_norm = output_orig[output_cons_col].fillna("").apply(normalize_lp_cons).tolist()

# ---------- Prepare Excel workbook for OUTPUT ----------
# Always create a fresh XLSX file for results
xlsx_output_path = os.path.splitext(output_path)[0] + "_processed.xlsx"
output_orig.to_excel(xlsx_output_path, index=False)   # create base Excel
wb = load_workbook(xlsx_output_path)
ws = wb.active

# Colors
GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")   # light green
YELLOW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")  # light yellow

# Insert "masterentity" header at the end
last_col_index = ws.max_column + 1
ws.cell(row=1, column=last_col_index).value = "masterentity"

exact_count = 0
partial_count = 0
no_match_count = 0
log_lines = []

total = len(output_orig)
t0 = time.time()

for i in range(total):
    excel_row = i + 2  # header is row 1
    lp = output_lp_norm[i]
    fund = output_fund_norm[i]
    cons = output_cons_norm[i]

    matched_original_fund = None
    fill_color = None

    # Exact match
    if (lp, fund, cons) in exact_map:
        matched_original_fund = exact_map[(lp, fund, cons)]
        fill_color = GREEN
        exact_count += 1
        log_lines.append(f"Row {excel_row} | Exact -> Master fund: {matched_original_fund}")
    else:
        # Partial match
        candidates = funds_by_lp_cons.get((lp, cons), [])
        for cand_norm, cand_orig in candidates:
            if fund and cand_norm and (fund in cand_norm or cand_norm in fund):
                matched_original_fund = cand_orig
                fill_color = YELLOW
                partial_count += 1
                log_lines.append(f"Row {excel_row} | Partial -> Master fund: {matched_original_fund}")
                break

    # Write masterentity or "No Match"
    if matched_original_fund:
        ws.cell(row=excel_row, column=last_col_index).value = matched_original_fund
    else:
        ws.cell(row=excel_row, column=last_col_index).value = "No Match"
        no_match_count += 1
        log_lines.append(f"Row {excel_row} | No Match")

    # Highlight entire row if match
    if fill_color:
        for col in range(1, ws.max_column + 1):
            ws.cell(row=excel_row, column=col).fill = fill_color

    # progress
    if (i + 1) % 100 == 0 or (i + 1) == total:
        elapsed = time.time() - t0
        print(f"Processed {i+1}/{total} rows — exact: {exact_count}, partial: {partial_count}, no-match: {no_match_count} (elapsed {elapsed:.1f}s)")

# ---------- Save workbook and log ----------
try:
    wb.save(xlsx_output_path)
except PermissionError:
    messagebox.showerror("Save Error", "Cannot save the output file. Please close it in Excel and run again.")
    raise

log_path = os.path.join(os.path.dirname(xlsx_output_path), "match_log.txt")
with open(log_path, "w", encoding="utf-8") as lf:
    lf.write("\n".join(log_lines))

# ---------- Summary popup ----------
summary = (
    f"Done!\n\n"
    f"Exact matches: {exact_count}\n"
    f"Partial matches: {partial_count}\n"
    f"No matches: {no_match_count}\n\n"
    f"Log: {log_path}\n"
    f"Output file updated: {xlsx_output_path}"
)
messagebox.showinfo("Match Summary", summary)
print(summary)
