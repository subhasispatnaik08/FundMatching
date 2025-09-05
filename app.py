# app.py
import streamlit as st
from io import BytesIO
from matching import process_files

st.set_page_config("Fund Matcher", layout="centered")
st.title("Fund Matching Tool")

st.markdown("Upload **Master Fund** and **Output** Excel files. Column names expected (case-insensitive):\n"
            "- Master: `LP name`, `Fund name`, `Consultant`\n"
            "- Output: `lpname`, `fundname`, `reportingconsultant`")

master_file = st.file_uploader("Upload Master Fund Excel", type=["xlsx"])
output_file = st.file_uploader("Upload Output Excel", type=["xlsx"])

if master_file and output_file:
    if st.button("Run Matching"):
        # read bytes
        master_bytes = master_file.read()
        output_bytes = output_file.read()

        with st.spinner("Running match — this may take some time for large files..."):
            result_bytes, stats = process_files(master_bytes, output_bytes)

        st.success("Done!")
        st.write(f"Exact matches: **{stats['exact']}**  Partial: **{stats['partial']}**  No Match: **{stats['nomatch']}**")

        # show small log tail
        if stats.get("log_lines"):
            st.subheader("Recent log lines")
            st.text("\n".join(stats["log_lines"][-10:]))

        # download button
        st.download_button(
            "Download highlighted Output (Excel)",
            data=result_bytes,
            file_name="output_highlighted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
