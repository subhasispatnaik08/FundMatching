# app.py
import streamlit as st
from matching import process_files

st.set_page_config("Fund Matcher", layout="centered")
st.title("Fund Matching Tool")

st.markdown("""
Upload **Master Fund** and **Output** files (CSV or Excel).  
Column names expected (case-insensitive):  
- Master: `LP name`, `Fund name`, `Consultant`  
- Output: `lpname`, `fundname`, `reportingconsultant`
""")

# ✅ Accept both CSV and Excel now
master_file = st.file_uploader("Upload Master Fund File", type=["csv", "xlsx"])
output_file = st.file_uploader("Upload Output File", type=["csv", "xlsx"])

if master_file and output_file:
    if st.button("Run Matching"):
        # read bytes
        master_bytes = master_file.read()
        output_bytes = output_file.read()

        with st.spinner("Running match — this may take some time for large files..."):
            result_bytes, stats = process_files(
                master_bytes, output_bytes, master_file.name, output_file.name
            )

        st.success("✅ Matching complete!")
        st.write(
            f"Exact matches: **{stats['exact']}** | "
            f"Partial matches: **{stats['partial']}** | "
            f"No Match: **{stats['nomatch']}**"
        )

        # show recent log lines
        if stats.get("log_lines"):
            st.subheader("Recent log lines")
            st.text("\n".join(stats["log_lines"][-10:]))

        # ✅ derive output name based on uploaded file
        base_name = output_file.name.rsplit(".", 1)[0] + ".xlsx"

        st.download_button(
            "⬇️ Download highlighted Output (Excel)",
            data=result_bytes.getvalue(),  # convert BytesIO → raw bytes
            file_name=base_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
