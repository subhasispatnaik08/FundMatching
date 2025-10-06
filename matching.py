# fund_matcher_streamlit.py
# Requires: pandas, openpyxl, streamlit
# pip install pandas openpyxl streamlit

import re
import os
import time
import pandas as pd
import streamlit as st
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
    # remove trailing variants like " lp", "l.p.", "llc"
    s = re.sub(r'\b(l\.?p\.?|llc)\b\.?$', '', s, flags=re.IGNORECASE).strip()
    return s.lower()

def find_column_ignore_case(df, target_lower):
    """Return the actual column name in df whose stripped-lower equals target_lower."""
    for col in df.columns:
        if str(col).strip().lower() == target_lower:
            return col
    raise KeyError(f"Column '{target_lower}' not found (checked case-insensitively).")

def read_file(uploaded_file):
    if uploaded_file is None:
        return None
    if uploaded_file.name.endswith(".csv"):
        return pd.read_csv(uploaded_file, dtype=str)
    else:
        return pd.read_excel(uploaded_file, dtype=str)

# ---------- Streamlit UI ----------
st.title("Fund Matcher v4 (Streamlit Edition)")

st.write("Upload your **MASTER fund file** and **OUTPUT file** (CSV or Excel).")
master_file = st.file_uploader("Upload MASTER file", type=["csv", "xlsx"])
output_file = st.file_uploader("Upload OUTPUT file", type=["csv", "xlsx"])

if master_file and output_file:
    try:
        # ---------- Read data ----------
        master_orig = read_file(master_file)
        output_orig = read_file(output_file)

        # ---------- Locate required columns ----------
        try:
            master_lp_col = find_column_ignore_case(master_orig, "lp name")
            master_fund_col = find_column_ignore_case(master_orig, "fund name")
            master_cons_col = find_column_ignore_case(master_orig, "consultant")
        except KeyError as e:
            st.error(f"Missing Column in MASTER file: {e}")
            st.stop()

        try:
            output_lp_col = find_column_ignore_case(output_orig, "lpname")
            output_fund_col = find_column_ignore_case(output_orig, "fundname")
            output_cons_col = find_column_ignore_case(output_orig, "reportingconsultant")
        except KeyError as e:
            st.error(f"Missing Column in OUTPUT file: {e}")
            st.stop()

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

        # ---------- Create Excel workbook ----------
        xlsx_output_path = "output_processed.xlsx"
        output_orig.to_excel(xlsx_output_path, index=False)  # temporary base
        wb = load_workbook(xlsx_output_path)
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

            # Exact
            if (lp, fund, cons) in exact_map:
                matched_original_fund = exact_map[(lp, fund, cons)]
                fill_color = GREEN
                exact_count += 1
                log_lines.append(f"Row {excel_row} | Exact -> Master fund: {matched_original_fund}")
            else:
                # Partial
                candidates = funds_by_lp_cons.get((lp, cons), [])
                for cand_norm, cand_orig in candidates:
                    if fund and cand_norm and (fund in cand_norm or cand_norm in fund):
                        matched_original_fund = cand_orig
                        fill_color = YELLOW
                        partial_count += 1
                        log_lines.append(f"Row {excel_row} | Partial -> Master fund: {matched_original_fund}")
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

        # ---------- Save workbook into memory ----------
        output_buffer = BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)

        # ---------- Log file ----------
        log_str = "\n".join(log_lines)
        log_buffer = BytesIO(log_str.encode("utf-8"))

        # ---------- Summary ----------
        summary = (
            f"✅ Done!\n\n"
            f"Exact matches: {exact_count}\n"
            f"Partial matches: {partial_count}\n"
            f"No matches: {no_match_count}\n"
            f"Processed {total} rows in {elapsed:.1f} seconds."
        )
        st.success(summary)

        # ---------- Download buttons ----------
        st.download_button(
            label="⬇️ Download Processed Excel",
            data=output_buffer,
            file_name="output_processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="⬇️ Download Log File",
            data=log_buffer,
            file_name="match_log.txt",
            mime="text/plain"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
