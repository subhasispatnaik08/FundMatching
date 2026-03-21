# app.py
import time
import random
import threading
import streamlit as st
from matching import process_files, PARTIAL_SIMILARITY_THRESHOLD

st.set_page_config("Fund Name Validator", layout="centered")

# ── Funny loading messages ───────────────────────────────
LOADING_MESSAGES = [
    "Sit back, relax — we've got this. ☕",
    "Hang on, incoming files… don't panic.",
    "Teaching the algorithm to read fund names… slowly.",
    "LP or not LP, that is the question.",
    "Matching in progress. Please do not turn off your brain.",
    "We promise this is faster than doing it manually.",
    "Running at full speed. (The speed of coffee.)",
    "This is fine. Everything is fine. 🔥",
    "Asking the data nicely to cooperate…",
    "Almost there. Probably. We think.",
    "The algorithm is giving it everything it's got.",
    "Fun fact: this is still faster than a VLOOKUP.",
]

st.markdown("""
<style>
  .stApp { background: #f8f8f6; }
  .block-container { max-width: 740px; padding-top: 2.5rem; padding-bottom: 3rem; }
  #MainMenu, footer, header { visibility: hidden; }

  .fnv-badge {
    display: inline-block;
    background: #e6f1fb; color: #185fa5;
    font-size: 11px; font-weight: 600;
    padding: 4px 12px; border-radius: 6px;
    border: 1px solid #b5d4f4;
    letter-spacing: 0.05em; margin-bottom: 0.6rem;
  }
  .fnv-title { font-size: 28px; font-weight: 700; color: #1c1c1a; margin: 0.2rem 0 0.4rem; }
  .fnv-sub   { font-size: 14px; color: #646460; margin-bottom: 1.6rem; line-height: 1.6; }

  .fnv-schema-row { display: flex; gap: 12px; margin-bottom: 1.6rem; }
  .fnv-schema-card {
    flex: 1; background: #ffffff;
    border: 1px solid #d2d0ca; border-radius: 10px; padding: 12px 14px;
  }
  .fnv-schema-label {
    font-size: 10px; font-weight: 700; color: #b4b2a9;
    letter-spacing: 0.07em; text-transform: uppercase; margin-bottom: 8px;
  }
  .fnv-pill {
    display: inline-block; font-size: 11.5px; font-family: monospace;
    background: #f1efe8; border: 1px solid #c3c1bb; border-radius: 5px;
    padding: 3px 9px; margin: 2px 3px 2px 0; color: #185fa5;
  }

  .fnv-step-row { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; margin-top: 0.5rem; }
  .fnv-step-num {
    width: 22px; height: 22px; border-radius: 50%;
    background: #e6f1fb; color: #185fa5;
    font-size: 11px; font-weight: 700;
    display: flex; align-items: center; justify-content: center; flex-shrink: 0;
  }
  .fnv-step-label { font-size: 13px; font-weight: 600; color: #1c1c1a; }

  [data-testid="stFileUploader"] {
    background: #ffffff; border: 1.5px dashed #c3c1bb;
    border-radius: 10px; padding: 8px 4px; transition: border-color 0.15s;
  }
  [data-testid="stFileUploader"]:hover { border-color: #185fa5; background: #f0f6fd; }
  [data-testid="stFileUploader"] label { display: none; }
  [data-testid="stFileUploaderDropzone"] { background: transparent !important; border: none !important; }


  /* Threshold section */
  .fnv-threshold-card {
    background: #ffffff; border: 1px solid #d2d0ca;
    border-radius: 10px; padding: 14px 16px; margin-bottom: 1rem;
  }
  .fnv-threshold-title { font-size: 13px; font-weight: 600; color: #1c1c1a; margin-bottom: 2px; }
  .fnv-threshold-hint  { font-size: 12px; color: #888780; margin-bottom: 0; }

  .fnv-divider { height: 1px; background: #d2d0ca; margin: 1.4rem 0 1.2rem; }

  .stButton > button {
    width: 100%; background: #1c1c1a !important; color: #ffffff !important;
    border: none !important; border-radius: 8px !important; padding: 12px !important;
    font-size: 14px !important; font-weight: 600 !important;
    letter-spacing: 0.01em; transition: opacity 0.15s !important;
  }
  .stButton > button:hover  { opacity: 0.85 !important; }
  .stButton > button:active { transform: scale(0.99); }

  /* Loading box */
  .fnv-loading {
    background: #ffffff; border: 1px solid #d2d0ca; border-radius: 10px;
    padding: 20px 20px 14px; margin-top: 1rem; text-align: center;
  }
  .fnv-loading-msg  { font-size: 14px; font-weight: 600; color: #1c1c1a; margin-bottom: 10px; }
  .fnv-loading-pct  { font-size: 12px; color: #888780; margin-top: 8px; }
  .fnv-progress-bar-bg {
    background: #f1efe8; border-radius: 99px; height: 6px; overflow: hidden; margin: 6px 0;
  }
  .fnv-progress-bar-fill {
    height: 6px; border-radius: 99px; background: #185fa5; transition: width 0.4s ease;
  }

  /* Stat cards */
  .fnv-stats-row { display: flex; gap: 10px; margin-top: 1.4rem; }
  .fnv-stat {
    flex: 1; background: #ffffff;
    border: 1px solid #d2d0ca; border-radius: 10px; padding: 14px 16px;
  }
  .fnv-stat-label {
    font-size: 11px; color: #b4b2a9; margin-bottom: 2px;
    font-weight: 500; text-transform: uppercase; letter-spacing: 0.05em;
  }
  .fnv-stat-val  { font-size: 26px; font-weight: 700; }
  .fnv-stat-sub  { font-size: 11px; color: #b4b2a9; margin-top: 2px; }
  .fnv-stat-val.exact   { color: #3b6d11; }
  .fnv-stat-val.partial { color: #ba7517; }
  .fnv-stat-val.nomatch { color: #a32d2d; }
  .fnv-stat-val.total   { color: #1c1c1a; }

  .fnv-success {
    background: #eaf3de; border: 1px solid #98d078; border-radius: 8px;
    padding: 10px 16px; font-size: 13px; color: #3b6d11;
    font-weight: 600; margin-top: 1rem;
  }

  [data-testid="stDownloadButton"] > button {
    width: 100%; background: #ffffff !important; color: #185fa5 !important;
    border: 1px solid #b5d4f4 !important; border-radius: 8px !important;
    font-size: 13px !important; font-weight: 600 !important;
    margin-top: 0.6rem; transition: background 0.15s !important;
  }
  [data-testid="stDownloadButton"] > button:hover { background: #e6f1fb !important; }

  .stExpander {
    border: 1px solid #d2d0ca !important;
    border-radius: 8px !important; background: #ffffff !important; margin-top: 0.8rem;
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

# ── Similarity threshold slider ──────────────────────────
st.markdown('<div class="fnv-divider"></div>', unsafe_allow_html=True)

with st.expander("⚙️  Advanced — similarity threshold", expanded=False):
    st.markdown("""
    <div class="fnv-threshold-hint">
    Controls how similar two fund names must be for a <b>Partial</b> match.<br>
    Higher = stricter (fewer partials). Lower = more lenient (more partials).<br>
    Default is <b>0.75</b> — recommended for most cases.
    </div>
    """, unsafe_allow_html=True)
    threshold = st.slider(
        "Similarity threshold",
        min_value=0.50, max_value=1.00, value=PARTIAL_SIMILARITY_THRESHOLD,
        step=0.05, label_visibility="collapsed"
    )
    col1, col2, col3 = st.columns(3)
    col1.caption("0.50 — lenient")
    col2.caption(f"← current: {threshold}")
    col3.caption("1.00 — strict")

# ── Results placeholder (renders above run button) ──────
results_slot = st.empty()

# ── Run button ───────────────────────────────────────────
run_clicked = st.button("Run validation", disabled=not (master_file and output_file))

# ── Processing ───────────────────────────────────────────
if run_clicked and master_file and output_file:
    master_bytes = master_file.read()
    output_bytes = output_file.read()

    # Shared state for threading
    result_container = {}
    result_container["done"] = False
    result_container["result"] = None
    result_container["error"] = None

    def run_matching():
        try:
            result_container["result"] = process_files(
                master_bytes, output_bytes,
                master_file.name, output_file.name,
                similarity_threshold=threshold
            )
        except Exception as e:
            result_container["error"] = str(e)
        finally:
            result_container["done"] = True

    thread = threading.Thread(target=run_matching)
    thread.start()

    # Pick one random message for this entire run
    run_msg = random.choice(LOADING_MESSAGES)

    loading_slot = st.empty()

    tick = 0
    FAKE_DURATION = 8   # seconds over which progress bar fills to ~90%

    while not result_container["done"]:
        pct = min(int((tick / FAKE_DURATION) * 90), 90)
        loading_slot.markdown(f"""
        <div class="fnv-loading">
          <div class="fnv-loading-msg">{run_msg}</div>
          <div class="fnv-progress-bar-bg">
            <div class="fnv-progress-bar-fill" style="width:{pct}%"></div>
          </div>
          <div class="fnv-loading-pct">{pct}% done — hang tight…</div>
        </div>
        """, unsafe_allow_html=True)
        time.sleep(1)
        tick += 1

    # Snap to 100%
    loading_slot.markdown("""
    <div class="fnv-loading">
      <div class="fnv-loading-msg">Done! That wasn't so bad, was it? 🎉</div>
      <div class="fnv-progress-bar-bg">
        <div class="fnv-progress-bar-fill" style="width:100%"></div>
      </div>
      <div class="fnv-loading-pct">100% — your file is ready.</div>
    </div>
    """, unsafe_allow_html=True)
    time.sleep(0.8)
    loading_slot.empty()

    if result_container["error"]:
        st.error(f"Something went wrong: {result_container['error']}")
        st.stop()

    result_bytes, stats = result_container["result"]
    elapsed = stats.get("elapsed", 0)
    base_name = output_file.name.rsplit(".", 1)[0] + "_matched.xlsx"

    # Render stats + download into the placeholder ABOVE the run button
    with results_slot.container():
        st.markdown(
            f'<div class="fnv-success">Matching complete — {stats["rows"]:,} rows processed in {elapsed:.1f}s. Your file is ready.</div>',
            unsafe_allow_html=True
        )
        st.markdown(f"""
        <div class="fnv-stats-row">
          <div class="fnv-stat">
            <div class="fnv-stat-label">Total rows</div>
            <div class="fnv-stat-val total">{stats['rows']:,}</div>
            <div class="fnv-stat-sub">in {elapsed:.1f}s</div>
          </div>
          <div class="fnv-stat">
            <div class="fnv-stat-label">Exact match</div>
            <div class="fnv-stat-val exact">{stats['exact']:,}</div>
            <div class="fnv-stat-sub">{stats['exact']/stats['rows']*100:.1f}% of rows</div>
          </div>
          <div class="fnv-stat">
            <div class="fnv-stat-label">Partial match</div>
            <div class="fnv-stat-val partial">{stats['partial']:,}</div>
            <div class="fnv-stat-sub">{stats['partial']/stats['rows']*100:.1f}% of rows</div>
          </div>
          <div class="fnv-stat">
            <div class="fnv-stat-label">No match</div>
            <div class="fnv-stat-val nomatch">{stats['nomatch']:,}</div>
            <div class="fnv-stat-sub">{stats['nomatch']/stats['rows']*100:.1f}% of rows</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.download_button(
            "Download highlighted output (Excel)",
            data=result_bytes.getvalue(),
            file_name=base_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if stats.get("log_lines"):
            with st.expander("View recent log lines"):
                st.text("\n".join(stats["log_lines"][-20:]))
