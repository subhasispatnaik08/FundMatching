# app.py
import time
import random
import threading
import streamlit as st
from matching import process_files, PARTIAL_SIMILARITY_THRESHOLD

st.set_page_config("Fund Name Validator", layout="wide", initial_sidebar_state="expanded")

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
  #MainMenu, footer, header { visibility: hidden; }

  /* ── Sidebar styling ── */
  [data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid #d2d0ca !important;
    padding-top: 1.5rem;
  }
  [data-testid="stSidebar"] .block-container { padding: 1rem 1.2rem; }

  /* ── Main canvas ── */
  .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 100%; }

  /* ── Badge ── */
  .fnv-badge {
    display: inline-block;
    background: #e6f1fb; color: #185fa5;
    font-size: 11px; font-weight: 600;
    padding: 4px 12px; border-radius: 6px;
    border: 1px solid #b5d4f4;
    letter-spacing: 0.05em; margin-bottom: 0.5rem;
  }
  .fnv-title { font-size: 22px; font-weight: 700; color: #1c1c1a; margin: 0.2rem 0 0.3rem; line-height: 1.3; }
  .fnv-sub   { font-size: 12px; color: #646460; margin-bottom: 1.2rem; line-height: 1.5; }

  /* ── Schema pills (sidebar) ── */
  .fnv-schema-card {
    background: #f8f8f6; border: 1px solid #d2d0ca;
    border-radius: 8px; padding: 10px 12px; margin-bottom: 0.6rem;
  }
  .fnv-schema-label {
    font-size: 10px; font-weight: 700; color: #b4b2a9;
    letter-spacing: 0.07em; text-transform: uppercase; margin-bottom: 6px;
  }
  .fnv-pill {
    display: inline-block; font-size: 10.5px; font-family: monospace;
    background: #ffffff; border: 1px solid #c3c1bb; border-radius: 4px;
    padding: 2px 7px; margin: 2px 2px 2px 0; color: #185fa5;
  }

  /* ── Step labels ── */
  .fnv-step-row { display: flex; align-items: center; gap: 7px; margin-bottom: 6px; margin-top: 0.8rem; }
  .fnv-step-num {
    width: 20px; height: 20px; border-radius: 50%;
    background: #e6f1fb; color: #185fa5;
    font-size: 10px; font-weight: 700;
    display: flex; align-items: center; justify-content: center; flex-shrink: 0;
  }
  .fnv-step-label { font-size: 12px; font-weight: 600; color: #1c1c1a; }

  /* ── File uploader ── */
  [data-testid="stFileUploader"] {
    background: #ffffff; border: 1.5px dashed #c3c1bb;
    border-radius: 8px; padding: 4px 2px;
  }
  [data-testid="stFileUploader"]:hover { border-color: #185fa5; background: #f0f6fd; }
  [data-testid="stFileUploader"] label { display: none; }
  [data-testid="stFileUploaderDropzone"] { background: transparent !important; border: none !important; }

  /* ── Divider ── */
  .fnv-divider { height: 1px; background: #d2d0ca; margin: 1rem 0; }

  /* ── Sidebar run button ── */
  .stButton > button {
    width: 100%; background: #1c1c1a !important; color: #ffffff !important;
    border: none !important; border-radius: 8px !important; padding: 11px !important;
    font-size: 13px !important; font-weight: 600 !important;
    letter-spacing: 0.01em; transition: opacity 0.15s !important;
  }
  .stButton > button:hover  { opacity: 0.85 !important; }
  .stButton > button:active { transform: scale(0.99); }

  /* ── Right canvas — empty state ── */
  .fnv-empty-state {
    display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    min-height: 60vh; text-align: center;
    color: #b4b2a9;
  }
  .fnv-empty-icon { font-size: 48px; margin-bottom: 1rem; }
  .fnv-empty-title { font-size: 16px; font-weight: 600; color: #888780; margin-bottom: 0.4rem; }
  .fnv-empty-hint  { font-size: 13px; color: #b4b2a9; }

  /* ── Loading box ── */
  .fnv-loading {
    background: #ffffff; border: 1px solid #d2d0ca; border-radius: 10px;
    padding: 28px 28px 20px; text-align: center; max-width: 520px; margin: 4rem auto;
  }
  .fnv-loading-msg  { font-size: 15px; font-weight: 600; color: #1c1c1a; margin-bottom: 14px; }
  .fnv-loading-pct  { font-size: 12px; color: #888780; margin-top: 10px; }
  .fnv-progress-bar-bg {
    background: #f1efe8; border-radius: 99px; height: 6px; overflow: hidden; margin: 6px 0;
  }
  .fnv-progress-bar-fill {
    height: 6px; border-radius: 99px; background: #185fa5; transition: width 0.4s ease;
  }

  /* ── Stat cards ── */
  .fnv-stats-grid {
    display: grid;
    grid-template-columns: repeat(4, minmax(0,1fr));
    gap: 12px; margin-bottom: 1.4rem;
  }
  .fnv-stat {
    background: #ffffff; border: 1px solid #d2d0ca;
    border-radius: 10px; padding: 16px 18px;
  }
  .fnv-stat-label {
    font-size: 11px; color: #b4b2a9; margin-bottom: 4px;
    font-weight: 500; text-transform: uppercase; letter-spacing: 0.05em;
  }
  .fnv-stat-val  { font-size: 28px; font-weight: 700; }
  .fnv-stat-sub  { font-size: 11px; color: #b4b2a9; margin-top: 3px; }
  .fnv-stat-val.exact   { color: #3b6d11; }
  .fnv-stat-val.partial { color: #ba7517; }
  .fnv-stat-val.nomatch { color: #a32d2d; }
  .fnv-stat-val.total   { color: #1c1c1a; }

  /* ── Success banner ── */
  .fnv-success {
    background: #eaf3de; border: 1px solid #98d078; border-radius: 8px;
    padding: 10px 16px; font-size: 13px; color: #3b6d11;
    font-weight: 600; margin-bottom: 1.2rem;
  }

  /* ── Download button ── */
  [data-testid="stDownloadButton"] > button {
    width: 100%; background: #ffffff !important; color: #185fa5 !important;
    border: 1px solid #b5d4f4 !important; border-radius: 8px !important;
    font-size: 13px !important; font-weight: 600 !important;
    margin-top: 0.4rem; transition: background 0.15s !important;
  }
  [data-testid="stDownloadButton"] > button:hover { background: #e6f1fb !important; }

  /* ── Log expander ── */
  .stExpander {
    border: 1px solid #d2d0ca !important;
    border-radius: 8px !important; background: #ffffff !important; margin-top: 0.8rem;
  }

  /* ── Threshold hint text ── */
  .fnv-threshold-hint { font-size: 11px; color: #888780; line-height: 1.5; margin-bottom: 0.4rem; }
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════
# SIDEBAR — controls panel
# ════════════════════════════════════════════════════════
with st.sidebar:

    st.markdown('<div class="fnv-badge">Fund Name Validator</div>', unsafe_allow_html=True)
    st.markdown('<div class="fnv-title">Match &amp; validate fund names</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="fnv-sub">Reconcile fund names between your master and output files.</div>',
        unsafe_allow_html=True
    )

    # Schema reference cards
    st.markdown("""
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
    """, unsafe_allow_html=True)

    # Upload 1
    st.markdown("""
    <div class="fnv-step-row">
      <div class="fnv-step-num">1</div>
      <div class="fnv-step-label">Master fund file</div>
    </div>
    """, unsafe_allow_html=True)
    master_file = st.file_uploader("master", type=["csv", "xlsx"], label_visibility="collapsed")

    # Upload 2
    st.markdown("""
    <div class="fnv-step-row">
      <div class="fnv-step-num">2</div>
      <div class="fnv-step-label">Output file</div>
    </div>
    """, unsafe_allow_html=True)
    output_file = st.file_uploader("output", type=["csv", "xlsx"], label_visibility="collapsed")

    # Threshold
    st.markdown('<div class="fnv-divider"></div>', unsafe_allow_html=True)
    with st.expander("⚙️  Similarity threshold", expanded=False):
        st.markdown("""
        <div class="fnv-threshold-hint">
        Higher = stricter (fewer partials).<br>
        Lower = more lenient (more partials).<br>
        Default <b>0.75</b> is recommended.
        </div>
        """, unsafe_allow_html=True)
        threshold = st.slider(
            "threshold", min_value=0.50, max_value=1.00,
            value=PARTIAL_SIMILARITY_THRESHOLD, step=0.05,
            label_visibility="collapsed"
        )
        c1, c2, c3 = st.columns(3)
        c1.caption("0.50")
        c2.caption(f"✦ {threshold}")
        c3.caption("1.00")

    # Run button
    st.markdown('<div class="fnv-divider"></div>', unsafe_allow_html=True)
    run_clicked = st.button("Run validation", disabled=not (master_file and output_file))


# ════════════════════════════════════════════════════════
# MAIN CANVAS — results panel
# ════════════════════════════════════════════════════════

canvas = st.empty()

# Empty state — shown before any run
if not run_clicked:
    canvas.markdown("""
    <div class="fnv-empty-state">
      <div class="fnv-empty-icon">📂</div>
      <div class="fnv-empty-title">No results yet</div>
      <div class="fnv-empty-hint">Upload both files and click Run validation to get started.</div>
    </div>
    """, unsafe_allow_html=True)

# ── Processing ───────────────────────────────────────────
if run_clicked and master_file and output_file:
    canvas.empty()

    master_bytes = master_file.read()
    output_bytes = output_file.read()

    result_container = {"done": False, "result": None, "error": None}

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

    # Single random message, stays for entire run
    run_msg = random.choice(LOADING_MESSAGES)
    loading_slot = st.empty()
    tick = 0
    FAKE_DURATION = 8

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

    # Snap to 100% briefly then clear
    loading_slot.markdown(f"""
    <div class="fnv-loading">
      <div class="fnv-loading-msg">{run_msg}</div>
      <div class="fnv-progress-bar-bg">
        <div class="fnv-progress-bar-fill" style="width:100%"></div>
      </div>
      <div class="fnv-loading-pct">100% done.</div>
    </div>
    """, unsafe_allow_html=True)
    time.sleep(0.6)
    loading_slot.empty()

    if result_container["error"]:
        st.error(f"Something went wrong: {result_container['error']}")
        st.stop()

    result_bytes, stats = result_container["result"]
    elapsed = stats.get("elapsed", 0)
    base_name = output_file.name.rsplit(".", 1)[0] + "_matched.xlsx"

    # ── Results ──
    st.markdown(
        f'<div class="fnv-success">Matching complete — {stats["rows"]:,} rows processed in {elapsed:.1f}s.</div>',
        unsafe_allow_html=True
    )

    st.markdown(f"""
    <div class="fnv-stats-grid">
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
