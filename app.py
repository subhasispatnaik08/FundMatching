# app.py
import streamlit as st
from matching import process_files

st.set_page_config("Fund Name Validator", layout="centered")

st.markdown("""
<style>
  /* ── Page background ── */
  .stApp { background: #f8f8f6; }
  .block-container { max-width: 740px; padding-top: 2.5rem; padding-bottom: 3rem; }

  /* ── Hide default Streamlit chrome ── */
  #MainMenu, footer, header { visibility: hidden; }

  /* ── Badge ── */
  .fnv-badge {
    display: inline-block;
    background: #e6f1fb;
    color: #185fa5;
    font-size: 11px;
    font-weight: 600;
    padding: 4px 12px;
    border-radius: 6px;
    border: 1px solid #b5d4f4;
    letter-spacing: 0.05em;
    margin-bottom: 0.6rem;
  }

  /* ── Title & subtitle ── */
  .fnv-title {
    font-size: 28px;
    font-weight: 700;
    color: #1c1c1a;
    margin: 0.2rem 0 0.4rem;
  }
  .fnv-sub {
    font-size: 14px;
    color: #646460;
    margin-bottom: 1.6rem;
    line-height: 1.6;
  }

  /* ── Schema cards row ── */
  .fnv-schema-row {
    display: flex;
    gap: 12px;
    margin-bottom: 1.6rem;
  }
  .fnv-schema-card {
    flex: 1;
    background: #ffffff;
    border: 1px solid #d2d0ca;
    border-radius: 10px;
    padding: 12px 14px;
  }
  .fnv-schema-label {
    font-size: 10px;
    font-weight: 700;
    color: #b4b2a9;
    letter-spacing: 0.07em;
    text-transform: uppercase;
    margin-bottom: 8px;
  }
  .fnv-pill {
    display: inline-block;
    font-size: 11.5px;
    font-family: monospace;
    background: #f1efe8;
    border: 1px solid #c3c1bb;
    border-radius: 5px;
    padding: 3px 9px;
    margin: 2px 3px 2px 0;
    color: #185fa5;
  }

  /* ── Step label ── */
  .fnv-step-row {
    display: flex;
    align-items: center;
    gap: 8px;
    margin-bottom: 8px;
    margin-top: 0.5rem;
  }
  .fnv-step-num {
    width: 22px; height: 22px;
    border-radius: 50%;
    background: #e6f1fb;
    color: #185fa5;
    font-size: 11px;
    font-weight: 700;
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0;
  }
  .fnv-step-label {
    font-size: 13px;
    font-weight: 600;
    color: #1c1c1a;
  }

  /* ── File uploader override ── */
  [data-testid="stFileUploader"] {
    background: #ffffff;
    border: 1.5px dashed #c3c1bb;
    border-radius: 10px;
    padding: 8px 4px;
    transition: border-color 0.15s;
  }
  [data-testid="stFileUploader"]:hover {
    border-color: #185fa5;
    background: #f0f6fd;
  }
  [data-testid="stFileUploader"] label { display: none; }
  [data-testid="stFileUploaderDropzone"] {
    background: transparent !important;
    border: none !important;
  }

  /* ── Divider ── */
  .fnv-divider {
    height: 1px;
    background: #d2d0ca;
    margin: 1.4rem 0 1.2rem;
  }

  /* ── Run button ── */
  .stButton > button {
    width: 100%;
    background: #1c1c1a !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 12px !important;
    font-size: 14px !important;
    font-weight: 600 !important;
    letter-spacing: 0.01em;
    transition: opacity 0.15s !important;
  }
  .stButton > button:hover { opacity: 0.85 !important; }
  .stButton > button:active { transform: scale(0.99); }

  /* ── Stat cards ── */
  .fnv-stats-row {
    display: flex;
    gap: 10px;
    margin-top: 1.4rem;
  }
  .fnv-stat {
    flex: 1;
    background: #ffffff;
    border: 1px solid #d2d0ca;
    border-radius: 10px;
    padding: 14px 16px;
  }
  .fnv-stat-label {
    font-size: 11px;
    color: #b4b2a9;
    margin-bottom: 4px;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.05em;
  }
  .fnv-stat-val {
    font-size: 26px;
    font-weight: 700;
  }
  .fnv-stat-val.exact   { color: #3b6d11; }
  .fnv-stat-val.partial { color: #ba7517; }
  .fnv-stat-val.nomatch { color: #a32d2d; }

  /* ── Success / download banner ── */
  .fnv-success {
    background: #eaf3de;
    border: 1px solid #98d078;
    border-radius: 8px;
    padding: 10px 16px;
    font-size: 13px;
    color: #3b6d11;
    font-weight: 600;
    margin-top: 1rem;
  }

  /* ── Download button ── */
  [data-testid="stDownloadButton"] > button {
    width: 100%;
    background: #ffffff !important;
    color: #185fa5 !important;
    border: 1px solid #b5d4f4 !important;
    border-radius: 8px !important;
    font-size: 13px !important;
    font-weight: 600 !important;
    margin-top: 0.6rem;
    transition: background 0.15s !important;
  }
  [data-testid="stDownloadButton"] > button:hover {
    background: #e6f1fb !important;
  }

  /* ── Log expander ── */
  .stExpander {
    border: 1px solid #d2d0ca !important;
    border-radius: 8px !important;
    background: #ffffff !important;
    margin-top: 0.8rem;
  }
</style>
""", unsafe_allow_html=True)


# ── Header ──────────────────────────────────────────────
st.markdown('<div class="fnv-badge">Fund Name Validator</div>', unsafe_allow_html=True)
st.markdown('<div class="fnv-title">Match &amp; validate fund names</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="fnv-sub">Upload your master and output files to reconcile fund names and flag discrepancies.</div>',
    unsafe_allow_html=True
)

# ── Schema cards ────────────────────────────────────────
st.markdown("""
<div class="fnv-schema-row">
  <div class="fnv-schema-card">
    <div class="fnv-schema-label">Master file expects</div>
    <span class="fnv-pill">lp name</span>
    <span class="fnv-pill">fund name</span>
    <span class="fnv-pill">consultant</span>
  </div>
  <div class="fnv-schema-card">
    <div class="fnv-schema-label">Output file expects</div>
    <span class="fnv-pill">lpname</span>
    <span class="fnv-pill">fundname</span>
    <span class="fnv-pill">reportingconsultant</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ── File uploaders ───────────────────────────────────────
st.markdown("""
<div class="fnv-step-row">
  <div class="fnv-step-num">1</div>
  <div class="fnv-step-label">Upload master fund file</div>
</div>
""", unsafe_allow_html=True)
master_file = st.file_uploader("master", type=["csv", "xlsx"], label_visibility="collapsed")

st.markdown("""
<div class="fnv-step-row">
  <div class="fnv-step-num">2</div>
  <div class="fnv-step-label">Upload output file</div>
</div>
""", unsafe_allow_html=True)
output_file = st.file_uploader("output", type=["csv", "xlsx"], label_visibility="collapsed")

# ── Run button ───────────────────────────────────────────
st.markdown('<div class="fnv-divider"></div>', unsafe_allow_html=True)

run_clicked = st.button("Run validation", disabled=not (master_file and output_file))

# ── Processing ───────────────────────────────────────────
if run_clicked and master_file and output_file:
    master_bytes = master_file.read()
    output_bytes = output_file.read()

    with st.spinner("Running match — this may take a moment for large files…"):
        result_bytes, stats = process_files(
            master_bytes, output_bytes,
            master_file.name, output_file.name
        )

    # Success banner
    st.markdown('<div class="fnv-success">Matching complete — your file is ready to download.</div>', unsafe_allow_html=True)

    # Stat cards
    st.markdown(f"""
    <div class="fnv-stats-row">
      <div class="fnv-stat">
        <div class="fnv-stat-label">Exact match</div>
        <div class="fnv-stat-val exact">{stats['exact']:,}</div>
      </div>
      <div class="fnv-stat">
        <div class="fnv-stat-label">Partial match</div>
        <div class="fnv-stat-val partial">{stats['partial']:,}</div>
      </div>
      <div class="fnv-stat">
        <div class="fnv-stat-label">No match</div>
        <div class="fnv-stat-val nomatch">{stats['nomatch']:,}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Download button
    base_name = output_file.name.rsplit(".", 1)[0] + "_matched.xlsx"
    st.download_button(
        "Download highlighted output (Excel)",
        data=result_bytes.getvalue(),
        file_name=base_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Log expander
    if stats.get("log_lines"):
        with st.expander("View recent log lines"):
            st.text("\n".join(stats["log_lines"][-20:]))
