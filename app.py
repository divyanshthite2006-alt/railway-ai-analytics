"""
app.py — Railway Point Testing Automation System
SECR · Signal & Telecommunication Department
Redesigned UI — no track animation, no operational window widget.
All business logic unchanged.
"""

import base64
from pathlib import Path

import streamlit as st
import pandas as pd
from io import BytesIO

from modules.processor import process_dataframe
from modules.report_builder import (
    build_excel_package,
    build_executive_pdf,
    build_station_summary,
    build_data_quality_report,
)
from modules.logger import write_log, read_logs
from modules.query_engine import run_query
from modules.batch_processor import process_multiple_files

# ──────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SECR · Point Testing",
    page_icon="🚆",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────────────────────────
# LOGO
# ──────────────────────────────────────────────────────────────────────────────
def _logo_data_uri() -> str:
    logo_path = Path(__file__).parent / "Assets" / "secr_logo.png"
    if logo_path.exists():
        encoded = base64.b64encode(logo_path.read_bytes()).decode("utf-8")
        return f"data:image/png;base64,{encoded}"
    return ""

LOGO_URI = _logo_data_uri()

FALLBACK_SVG = """
<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg" width="52" height="52">
  <circle cx="32" cy="32" r="30" fill="#0D1B2A" stroke="#E8A020" stroke-width="2"/>
  <path d="M20 42 L20 24 Q20 18 26 18 L38 18 Q44 18 44 24 L44 42 Z"
        fill="none" stroke="#E8A020" stroke-width="2.5"/>
  <circle cx="26" cy="40" r="3" fill="#E8A020"/>
  <circle cx="38" cy="40" r="3" fill="#E8A020"/>
  <path d="M14 48 L50 48" stroke="#E8A020" stroke-width="2"/>
</svg>"""

def logo_html(size: int = 52) -> str:
    if LOGO_URI:
        return (f'<img src="{LOGO_URI}" alt="SECR" '
                f'style="height:{size}px;width:{size}px;object-fit:contain;'
                f'border-radius:50%;background:#fff;padding:3px;" />')
    return FALLBACK_SVG

# ──────────────────────────────────────────────────────────────────────────────
# GLOBAL STYLES
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">

<style>
/* ── Reset ─────────────────────────────────────────────────────────────── */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --ink:          #0D1B2A;
  --ink-mid:      #1A3050;
  --ink-soft:     #2C4A6E;
  --gold:         #E8A020;
  --gold-dim:     #C4861A;
  --gold-tint:    rgba(232,160,32,0.12);
  --gold-tint2:   rgba(232,160,32,0.06);
  --surface:      #F7F8FA;
  --white:        #FFFFFF;
  --border:       #E4E8EF;
  --border-dark:  #C8D0DC;
  --text:         #1A2535;
  --muted:        #6B7A8D;
  --muted-light:  #9AAABB;
  --green:        #0F9E6E;
  --green-bg:     #EDFAF4;
  --red:          #D94040;
  --red-bg:       #FEF1F1;
  --blue:         #2563EB;
  --blue-bg:      #EFF4FF;
  --r4:           4px;
  --r8:           8px;
  --r12:          12px;
  --r16:          16px;
  --sh-xs:        0 1px 3px rgba(13,27,42,0.07);
  --sh-sm:        0 2px 8px rgba(13,27,42,0.09);
  --sh-md:        0 4px 18px rgba(13,27,42,0.11);
  --ease:         cubic-bezier(0.4,0,0.2,1);
}

/* ── Base ─────────────────────────────────────────────────────────────── */
html, body, [class*="css"] {
  font-family: 'DM Sans', sans-serif !important;
  background: var(--surface) !important;
  color: var(--text) !important;
}
.stApp { background: var(--surface) !important; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 0 !important; max-width: 100% !important; }

/* ── Sidebar ──────────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
  background: var(--ink) !important;
  border-right: 1px solid rgba(255,255,255,0.05) !important;
  width: 260px !important;
  box-shadow: 2px 0 16px rgba(0,0,0,0.18) !important;
}
[data-testid="stSidebar"] > div:first-child { padding: 0 !important; }
section[data-testid="stSidebar"] .block-container { padding: 0 !important; }

/* Sidebar brand block */
.sb-brand {
  padding: 24px 20px 20px;
  border-bottom: 1px solid rgba(255,255,255,0.07);
  margin-bottom: 8px;
  display: flex;
  align-items: center;
  gap: 14px;
}
.sb-brand-text {}
.sb-org {
  font-family: 'DM Sans', sans-serif;
  font-size: 18px;
  font-weight: 700;
  color: #fff;
  letter-spacing: 0.5px;
  line-height: 1.2;
}
.sb-tag {
  display: inline-block;
  margin-top: 4px;
  font-size: 10px;
  color: var(--gold);
  font-weight: 600;
  letter-spacing: 1.4px;
  text-transform: uppercase;
}
.sb-dept {
  margin-top: 3px;
  font-size: 11px;
  color: rgba(255,255,255,0.38);
  line-height: 1.4;
}

/* Sidebar nav label */
.sb-nav-label {
  font-size: 10px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 1.5px;
  color: rgba(255,255,255,0.3);
  padding: 14px 20px 6px;
}

/* Sidebar radio → nav items */
[data-testid="stSidebar"] .stRadio { }
[data-testid="stSidebar"] .stRadio > div {
  gap: 2px !important;
  padding: 0 10px !important;
}
[data-testid="stSidebar"] .stRadio label {
  background: transparent !important;
  border-radius: var(--r8) !important;
  padding: 10px 14px !important;
  color: rgba(255,255,255,0.55) !important;
  font-size: 13.5px !important;
  font-weight: 500 !important;
  transition: all 0.18s var(--ease) !important;
  border: none !important;
  cursor: pointer !important;
  width: 100% !important;
  display: flex !important;
  align-items: center !important;
}
[data-testid="stSidebar"] .stRadio label:hover {
  background: rgba(255,255,255,0.06) !important;
  color: #fff !important;
}
[data-testid="stSidebar"] .stRadio [data-testid="stMarkdownContainer"] p {
  color: inherit !important;
  font-size: 13.5px !important;
}
[data-testid="stSidebar"] .stRadio input[type="radio"]:checked + div label,
[data-testid="stSidebar"] .stRadio input[type="radio"]:checked ~ label {
  background: var(--gold-tint) !important;
  color: var(--gold) !important;
  font-weight: 600 !important;
  border-left: 2px solid var(--gold) !important;
  padding-left: 12px !important;
}
[data-testid="stSidebar"] .stRadio input[type="radio"] { display: none !important; }
[data-testid="stSidebar"] .stRadio > label { display: none !important; }

/* Sidebar footer */
.sb-footer {
  position: absolute;
  bottom: 0; left: 0; right: 0;
  padding: 14px 20px;
  border-top: 1px solid rgba(255,255,255,0.06);
  font-size: 11px;
  color: rgba(255,255,255,0.22);
  letter-spacing: 0.3px;
}

/* ── Page header (no animation, no track) ─────────────────────────────── */
.pg-header {
  background: var(--ink);
  padding: 22px 32px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  border-bottom: 2px solid var(--gold);
  margin-bottom: 28px;
}
.pg-header-left {
  display: flex;
  align-items: center;
  gap: 16px;
}
.pg-divider {
  width: 1px;
  height: 40px;
  background: rgba(255,255,255,0.15);
}
.pg-title {
  font-size: 22px;
  font-weight: 700;
  color: #fff;
  letter-spacing: 0.2px;
  line-height: 1.2;
}
.pg-sub {
  font-size: 12.5px;
  color: rgba(255,255,255,0.45);
  margin-top: 3px;
}
.pg-badge {
  background: var(--gold-tint);
  border: 1px solid rgba(232,160,32,0.35);
  border-radius: 20px;
  padding: 5px 14px;
  font-size: 11.5px;
  font-weight: 600;
  color: var(--gold);
  letter-spacing: 0.4px;
  text-transform: uppercase;
  white-space: nowrap;
}
.status-dot {
  width: 6px; height: 6px;
  border-radius: 50%;
  background: var(--green);
  display: inline-block;
  margin-right: 5px;
  box-shadow: 0 0 0 2px rgba(15,158,110,0.25);
}

/* ── Main content wrap ────────────────────────────────────────────────── */
.main-wrap {
  padding: 0 28px 48px;
  max-width: 1360px;
  margin: 0 auto;
}

/* ── Section title ────────────────────────────────────────────────────── */
.sec-title {
  font-size: 17px;
  font-weight: 700;
  color: var(--ink);
  margin-bottom: 3px;
  display: flex;
  align-items: center;
  gap: 9px;
}
.sec-icon {
  width: 26px; height: 26px;
  background: var(--gold);
  border-radius: var(--r4);
  display: inline-flex;
  align-items: center;
  justify-content: center;
  font-size: 13px;
  flex-shrink: 0;
}
.sec-sub {
  font-size: 12.5px;
  color: var(--muted);
  margin-bottom: 14px;
  margin-left: 35px;
}

/* ── Cards ────────────────────────────────────────────────────────────── */
.card {
  background: var(--white);
  border-radius: var(--r12);
  border: 1px solid var(--border);
  padding: 20px 22px;
  box-shadow: var(--sh-xs);
  margin-bottom: 18px;
  transition: box-shadow 0.2s var(--ease);
}
.card:hover { box-shadow: var(--sh-sm); }
.card-l { border-left: 3px solid var(--gold) !important; }

/* ── Metric grid ──────────────────────────────────────────────────────── */
.metrics {
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 12px;
  margin-bottom: 22px;
}
.mc {
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--r12);
  padding: 18px 14px 16px;
  text-align: center;
  box-shadow: var(--sh-xs);
  transition: all 0.2s var(--ease);
}
.mc:hover { box-shadow: var(--sh-sm); transform: translateY(-1px); }
.mc-top { border-top: 3px solid var(--gold); }
.mc-green { border-top: 3px solid var(--green); }
.mc-red   { border-top: 3px solid var(--red); }
.mv {
  font-family: 'DM Mono', monospace;
  font-size: 32px;
  font-weight: 500;
  color: var(--ink);
  line-height: 1.1;
}
.mv-g { color: var(--green); }
.mv-r { color: var(--red); }
.mv-a { color: var(--gold-dim); }
.ml {
  font-size: 11px;
  font-weight: 600;
  color: var(--muted);
  text-transform: uppercase;
  letter-spacing: 0.5px;
  margin-top: 5px;
}
.exec-bar {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  color: var(--muted);
  margin-bottom: 20px;
  padding: 9px 14px;
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--r8);
  box-shadow: var(--sh-xs);
}
.exec-bar strong { color: var(--ink); }
.exec-sep { color: var(--border-dark); margin: 0 4px; }

/* ── Banners ──────────────────────────────────────────────────────────── */
.bn {
  padding: 10px 14px;
  border-radius: var(--r8);
  font-size: 13px;
  font-weight: 500;
  margin-bottom: 14px;
  display: flex;
  align-items: flex-start;
  gap: 8px;
  line-height: 1.5;
}
.bn-ok    { background: var(--green-bg); color: #0A6A4A; border-left: 3px solid var(--green); }
.bn-info  { background: var(--blue-bg);  color: #1A49B8; border-left: 3px solid var(--blue); }
.bn-warn  { background: #FFF8EC;         color: #8A5A00; border-left: 3px solid var(--gold); }
.bn-err   { background: var(--red-bg);   color: #9B2020; border-left: 3px solid var(--red); }

/* ── Streamlit overrides ─────────────────────────────────────────────── */
/* Primary button */
.stButton > button[kind="primary"] {
  background: var(--gold) !important;
  color: var(--ink) !important;
  border: none !important;
  border-radius: var(--r8) !important;
  font-weight: 700 !important;
  font-size: 13.5px !important;
  padding: 9px 22px !important;
  box-shadow: 0 2px 8px rgba(232,160,32,0.3) !important;
  transition: all 0.18s var(--ease) !important;
}
.stButton > button[kind="primary"]:hover {
  background: var(--gold-dim) !important;
  box-shadow: 0 4px 14px rgba(232,160,32,0.4) !important;
  transform: translateY(-1px) !important;
}
/* Secondary button */
.stButton > button:not([kind="primary"]) {
  background: var(--white) !important;
  color: var(--ink) !important;
  border: 1.5px solid var(--border-dark) !important;
  border-radius: var(--r8) !important;
  font-weight: 500 !important;
  font-size: 13px !important;
  transition: all 0.18s var(--ease) !important;
}
.stButton > button:not([kind="primary"]):hover {
  border-color: var(--ink-soft) !important;
  box-shadow: var(--sh-xs) !important;
}
/* Download button */
.stDownloadButton > button {
  background: var(--white) !important;
  color: var(--ink) !important;
  border: 1.5px solid var(--border) !important;
  border-radius: var(--r8) !important;
  font-weight: 600 !important;
  font-size: 13px !important;
  padding: 9px 18px !important;
  transition: all 0.18s var(--ease) !important;
}
.stDownloadButton > button:hover {
  border-color: var(--gold) !important;
  color: var(--gold-dim) !important;
  box-shadow: 0 2px 10px rgba(232,160,32,0.15) !important;
}
/* File uploader */
[data-testid="stFileUploader"] {
  background: var(--white) !important;
  border: 2px dashed var(--border-dark) !important;
  border-radius: var(--r12) !important;
  transition: border-color 0.18s var(--ease) !important;
}
[data-testid="stFileUploader"]:hover { border-color: var(--gold) !important; }
[data-testid="stFileUploaderDropzone"] { background: transparent !important; }
/* Text input */
.stTextInput > div > div > input {
  border: 1.5px solid var(--border) !important;
  border-radius: var(--r8) !important;
  font-size: 14px !important;
  padding: 10px 14px !important;
  background: var(--white) !important;
  transition: border-color 0.18s var(--ease) !important;
}
.stTextInput > div > div > input:focus {
  border-color: var(--gold) !important;
  box-shadow: 0 0 0 3px rgba(232,160,32,0.12) !important;
  outline: none !important;
}
/* Expanders */
.streamlit-expanderHeader {
  background: var(--white) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--r8) !important;
  font-weight: 600 !important;
  font-size: 13.5px !important;
  color: var(--ink) !important;
  padding: 11px 15px !important;
}
.streamlit-expanderHeader:hover {
  border-color: var(--gold) !important;
}
.streamlit-expanderContent {
  border: 1px solid var(--border) !important;
  border-top: none !important;
  border-radius: 0 0 var(--r8) var(--r8) !important;
  background: var(--white) !important;
  padding: 14px !important;
}
/* Dataframe */
[data-testid="stDataFrame"] {
  border-radius: var(--r12) !important;
  overflow: hidden !important;
  border: 1px solid var(--border) !important;
  box-shadow: var(--sh-xs) !important;
}
/* Spinner */
.stSpinner > div { border-top-color: var(--gold) !important; }
/* Alerts */
.stSuccess {
  background: var(--green-bg) !important; color: #0A6A4A !important;
  border-left: 3px solid var(--green) !important;
  border-radius: var(--r8) !important; font-weight: 500 !important;
}
.stError {
  background: var(--red-bg) !important; color: #9B2020 !important;
  border-left: 3px solid var(--red) !important;
  border-radius: var(--r8) !important;
}
.stInfo {
  background: var(--blue-bg) !important; color: #1A49B8 !important;
  border-left: 3px solid var(--blue) !important;
  border-radius: var(--r8) !important;
}
.stWarning {
  background: #FFF8EC !important; color: #8A5A00 !important;
  border-left: 3px solid var(--gold) !important;
  border-radius: var(--r8) !important;
}
/* Divider */
hr { border: none !important; border-top: 1px solid var(--border) !important; margin: 22px 0 !important; }
.stCaption, .stCaption p { color: var(--muted) !important; font-size: 12px !important; }

/* ── Query chips ──────────────────────────────────────────────────────── */
.chip {
  display: inline-block;
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: 20px;
  padding: 4px 12px;
  font-size: 12px;
  font-weight: 500;
  color: var(--ink-soft);
  margin: 3px 3px 3px 0;
  cursor: pointer;
  transition: all 0.15s var(--ease);
}
.chip:hover { background: #FFF8EC; border-color: var(--gold); color: var(--gold-dim); }

/* ── Scrollbar ────────────────────────────────────────────────────────── */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: var(--surface); }
::-webkit-scrollbar-thumb { background: var(--border-dark); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: var(--ink-soft); }

/* ── Responsive ───────────────────────────────────────────────────────── */
@media (max-width: 900px) {
  .metrics { grid-template-columns: repeat(2, 1fr); }
  .main-wrap { padding: 0 14px 32px; }
  .pg-header { padding: 18px 18px; flex-direction: column; align-items: flex-start; gap: 10px; }
}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div class="sb-brand">
      <div>{logo_html(52)}</div>
      <div class="sb-brand-text">
        <div class="sb-org">SECR</div>
        <div class="sb-tag">Indian Railways</div>
        <div class="sb-dept">Signal &amp; Telecommunication<br>Department</div>
      </div>
    </div>
    <div class="sb-nav-label">Modules</div>
    """, unsafe_allow_html=True)

    mode = st.radio(
        "nav",
        [
            "🗂️  Single File",
            "📦  Batch Processing",
            "📜  Processing Logs",
            "🔍  AI Query Engine",
        ],
        label_visibility="collapsed",
    )

# ──────────────────────────────────────────────────────────────────────────────
# PAGE HEADER
# ──────────────────────────────────────────────────────────────────────────────
PAGE_META = {
    "🗂️  Single File":      ("Single File Processing",   "Upload and process one Excel file"),
    "📦  Batch Processing":  ("Batch Processing",          "Process multiple files and merge results"),
    "📜  Processing Logs":   ("Processing Logs",           "Audit trail of all processing runs"),
    "🔍  AI Query Engine":   ("AI Query Engine",           "Query your results in plain English"),
}
title_text, sub_text = PAGE_META.get(mode, ("Point Testing", ""))

st.markdown(f"""
<div class="pg-header">
  <div class="pg-header-left">
    <div>{logo_html(44)}</div>
    <div class="pg-divider"></div>
    <div>
      <div class="pg-title">{title_text}</div>
      <div class="pg-sub">
        <span class="status-dot"></span>System Online
        &nbsp;·&nbsp; {sub_text}
      </div>
    </div>
  </div>
  <div class="pg-badge">SECR · S&amp;T Dept.</div>
</div>
<div class="main-wrap">
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────────────────────
def render_metrics(result: dict):
    t     = result["tested_count"]
    r     = result["review_count"]
    total = t + r
    pct   = f"{round(t / total * 100)}%" if total else "—"

    st.markdown(f"""
    <div class="metrics">
      <div class="mc mc-top">
        <div class="mv">{result['rows_processed']}</div>
        <div class="ml">Records</div>
      </div>
      <div class="mc">
        <div class="mv">{result['stations_processed']}</div>
        <div class="ml">Stations</div>
      </div>
      <div class="mc">
        <div class="mv">{result['points_processed']}</div>
        <div class="ml">Points</div>
      </div>
      <div class="mc mc-green">
        <div class="mv mv-g">{t}</div>
        <div class="ml">Tested</div>
      </div>
      <div class="mc mc-red">
        <div class="mv mv-r">{r}</div>
        <div class="ml">Manual Review</div>
      </div>
    </div>
    <div class="exec-bar">
      <span class="status-dot"></span>
      Execution <strong>{result['execution_time']}s</strong>
      <span class="exec-sep">|</span>
      Generated <strong>{result['generated_on']}</strong>
      <span class="exec-sep">|</span>
      Compliance <strong style="color:var(--green)">{pct}</strong>
    </div>
    """, unsafe_allow_html=True)


def render_section(icon: str, title: str, sub: str = ""):
    sub_html = f'<div class="sec-sub">{sub}</div>' if sub else ""
    st.markdown(f"""
    <div style="margin-bottom:6px;">
      <div class="sec-title">
        <span class="sec-icon">{icon}</span>{title}
      </div>
      {sub_html}
    </div>
    """, unsafe_allow_html=True)


def show_station_summary(result: dict):
    render_section("📊", "Station Summary", "Point testing status grouped by station")
    summary_df = build_station_summary(result["final_rows"])
    if summary_df.empty:
        st.markdown('<div class="bn bn-info">ℹ️ No station data available.</div>', unsafe_allow_html=True)
    else:
        st.dataframe(summary_df, use_container_width=True, hide_index=True)


def show_downloads(result: dict, df_raw: pd.DataFrame):
    excel_buf = build_excel_package(result, df_raw)
    pdf_buf   = build_executive_pdf(result)

    render_section("⬇️", "Download Reports")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="card card-l">'
                    '<div style="font-weight:700;color:var(--ink);margin-bottom:3px;">📊 Full Report Package</div>'
                    '<div style="font-size:12px;color:var(--muted);margin-bottom:12px;">'
                    'Final · Exception · Audit · Station Summary · Data Quality</div>',
                    unsafe_allow_html=True)
        st.download_button(
            "Download .xlsx Package",
            data=excel_buf,
            file_name="SECR_Railway_Report_Package.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="card card-l">'
                    '<div style="font-weight:700;color:var(--ink);margin-bottom:3px;">📄 Executive Summary</div>'
                    '<div style="font-size:12px;color:var(--muted);margin-bottom:12px;">'
                    'PDF report for senior officers and records</div>',
                    unsafe_allow_html=True)
        st.download_button(
            "Download .pdf Summary",
            data=pdf_buf,
            file_name="SECR_Executive_Summary.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# MODE 1 — SINGLE FILE
# ──────────────────────────────────────────────────────────────────────────────
if mode == "🗂️  Single File":

    st.markdown('<div class="card">', unsafe_allow_html=True)
    render_section("📁", "Upload File", "Accepts .xlsx files exported from the point testing system")
    uploaded_file = st.file_uploader(
        "Drop Excel file here or click to browse",
        type=["xlsx"],
        label_visibility="collapsed",
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file, header=4)
            st.markdown(
                f'<div class="bn bn-ok">✓ &nbsp;<strong>{uploaded_file.name}</strong> '
                f'loaded — {len(df)} rows detected</div>',
                unsafe_allow_html=True,
            )

            with st.expander("Preview Input Data", expanded=False):
                st.dataframe(df.head(10), use_container_width=True, hide_index=True)

            st.markdown('<div style="height:4px;"></div>', unsafe_allow_html=True)

            if st.button("⚙️  Generate Report", type="primary"):
                with st.spinner("Processing records…"):
                    result = process_dataframe(df, filename=uploaded_file.name)
                    write_log(result)

                st.markdown('<hr>', unsafe_allow_html=True)
                render_section("⚡", "Execution Metrics")
                render_metrics(result)

                st.markdown('<hr>', unsafe_allow_html=True)
                render_section("📋", "Final Report", "All classified point-position combinations")
                final_df = pd.DataFrame(result["final_rows"])
                st.dataframe(
                    final_df.style.apply(
                        lambda row: [
                            "background-color:#C6EFCE" if row["Status"] == "Tested"
                            else "background-color:#FFC7CE"
                            for _ in row
                        ], axis=1,
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

                st.markdown('<hr>', unsafe_allow_html=True)

                with st.expander("🔍  Audit Report — Verification Trail"):
                    audit_df = pd.DataFrame(result["audit_rows"])
                    st.dataframe(audit_df, use_container_width=True, hide_index=True)

                review_n = result["review_count"]
                exc_label = (
                    f"⚠️  Exception Report — {review_n} record(s) need review"
                    if review_n > 0 else
                    "✅  Exception Report — No exceptions found"
                )
                with st.expander(exc_label, expanded=review_n > 0):
                    if result["exception_rows"]:
                        exc_df = pd.DataFrame(result["exception_rows"])
                        st.dataframe(exc_df, use_container_width=True, hide_index=True)
                    else:
                        st.markdown(
                            '<div class="bn bn-ok">✓ All records classified as Tested. No exceptions.</div>',
                            unsafe_allow_html=True,
                        )

                with st.expander("🧪  Data Quality Report"):
                    dq_df = build_data_quality_report(df)
                    st.dataframe(dq_df, use_container_width=True, hide_index=True)

                st.markdown('<hr>', unsafe_allow_html=True)
                show_station_summary(result)

                st.markdown('<hr>', unsafe_allow_html=True)
                show_downloads(result, df)

                st.markdown(
                    '<div class="bn bn-ok" style="margin-top:14px;">'
                    '✓ &nbsp;Report generated successfully and ready for download.</div>',
                    unsafe_allow_html=True,
                )

        except Exception as e:
            st.markdown(
                f'<div class="bn bn-err">✗ &nbsp;{str(e)}</div>',
                unsafe_allow_html=True,
            )


# ──────────────────────────────────────────────────────────────────────────────
# MODE 2 — BATCH PROCESSING
# ──────────────────────────────────────────────────────────────────────────────
elif mode == "📦  Batch Processing":

    st.markdown('<div class="card">', unsafe_allow_html=True)
    render_section("📦", "Batch Upload", "Select multiple .xlsx files — results merged with deduplication")
    st.markdown(
        '<div class="bn bn-info">ℹ️ Hold <kbd>Ctrl</kbd> (or <kbd>Cmd</kbd> on Mac) '
        'to select multiple files in the picker.</div>',
        unsafe_allow_html=True,
    )
    uploaded_files = st.file_uploader(
        "Select multiple Excel files",
        type=["xlsx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_files:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(
            f'<div style="font-weight:700;color:var(--ink);margin-bottom:8px;">'
            f'{len(uploaded_files)} file(s) queued</div>',
            unsafe_allow_html=True,
        )
        for uf in uploaded_files:
            st.markdown(
                f'<div style="font-size:13px;color:var(--muted);padding:3px 0;">📄 &nbsp;{uf.name}</div>',
                unsafe_allow_html=True,
            )
        st.markdown('</div>', unsafe_allow_html=True)

        if st.button(f"⚙️  Process {len(uploaded_files)} File(s)", type="primary"):
            with st.spinner(f"Processing {len(uploaded_files)} file(s)…"):
                result = process_multiple_files(uploaded_files)
                for uf in uploaded_files:
                    write_log({**result, "filename": uf.name})

            if result:
                render_section("⚡", "Batch Metrics")
                render_metrics(result)
                st.markdown('<hr>', unsafe_allow_html=True)

                render_section("📋", "Combined Final Report", "Merged and deduplicated across all files")
                final_df = pd.DataFrame(result["final_rows"])
                st.dataframe(final_df, use_container_width=True, hide_index=True)
                st.markdown('<hr>', unsafe_allow_html=True)

                show_station_summary(result)
                st.markdown('<hr>', unsafe_allow_html=True)

                try:
                    uploaded_files[0].seek(0)
                    df_sample = pd.read_excel(uploaded_files[0], header=4)
                except Exception:
                    df_sample = pd.DataFrame()

                show_downloads(result, df_sample)

                st.markdown(
                    '<div class="bn bn-ok" style="margin-top:14px;">'
                    '✓ &nbsp;Batch report generated successfully.</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    '<div class="bn bn-err">✗ &nbsp;No valid data could be processed.</div>',
                    unsafe_allow_html=True,
                )


# ──────────────────────────────────────────────────────────────────────────────
# MODE 3 — PROCESSING LOGS
# ──────────────────────────────────────────────────────────────────────────────
elif mode == "📜  Processing Logs":

    logs = read_logs()

    if not logs:
        st.markdown("""
        <div class="card" style="text-align:center;padding:48px 20px;">
          <div style="font-size:36px;margin-bottom:12px;">📭</div>
          <div style="font-size:18px;font-weight:700;color:var(--ink);margin-bottom:6px;">No logs yet</div>
          <div style="font-size:13px;color:var(--muted);">
            Generate a report from Single File or Batch Processing to start logging.
          </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        logs_df = pd.DataFrame(logs)
        total_tested = logs_df["Tested Count"].astype(int).sum()
        total_review = logs_df["Manual Review Count"].astype(int).sum()
        total_rows   = logs_df["Rows Processed"].astype(int).sum()

        st.markdown(f"""
        <div class="metrics" style="grid-template-columns:repeat(3,1fr);">
          <div class="mc mc-top">
            <div class="mv">{len(logs_df)}</div>
            <div class="ml">Total Runs</div>
          </div>
          <div class="mc">
            <div class="mv">{total_rows}</div>
            <div class="ml">Rows Processed</div>
          </div>
          <div class="mc mc-green">
            <div class="mv mv-g">{total_tested}</div>
            <div class="ml">Total Tested</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        render_section("📜", "Processing History", "Auto-saved to logs/processing_log.csv")
        st.dataframe(logs_df, use_container_width=True, hide_index=True)

        log_csv = logs_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥  Download Log as CSV",
            data=log_csv,
            file_name="SECR_processing_log.csv",
            mime="text/csv",
        )


# ──────────────────────────────────────────────────────────────────────────────
# MODE 4 — AI QUERY ENGINE
# ──────────────────────────────────────────────────────────────────────────────
elif mode == "🔍  AI Query Engine":

    EXAMPLE_QUERIES = [
        "Show all points tested on 15",
        "Which station has most reviews?",
        "Show all manual review points",
        "Show testing activity for station DPH",
        "Show all tested points",
        "Show all points",
    ]

    st.markdown('<div class="card">', unsafe_allow_html=True)
    render_section("🔍", "AI Query Engine", "Ask questions about your data in plain English — no SQL, no formulas")
    st.markdown(
        '<div class="bn bn-info">ℹ️ Runs entirely on Pandas. Your data never leaves your machine.</div>',
        unsafe_allow_html=True,
    )
    uploaded_file = st.file_uploader(
        "Upload Excel file to query",
        type=["xlsx"],
        key="query_uploader",
        label_visibility="collapsed",
    )
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown(
        '<div style="font-size:11.5px;font-weight:700;color:var(--muted);'
        'text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">Example queries</div>'
        + "".join(f'<span class="chip">{q}</span>' for q in EXAMPLE_QUERIES),
        unsafe_allow_html=True,
    )
    st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)

    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file, header=4)
            result = process_dataframe(df_raw, filename=uploaded_file.name)
            final_df = pd.DataFrame(result["final_rows"])

            st.markdown(
                f'<div class="bn bn-ok">✓ &nbsp;<strong>{uploaded_file.name}</strong> '
                f'ready — {len(final_df)} classified records available.</div>',
                unsafe_allow_html=True,
            )

            query = st.text_input(
                "Your query",
                placeholder="e.g.  Show all manual review points for station BLP",
                label_visibility="collapsed",
            )

            if query:
                with st.spinner("Querying…"):
                    filtered, explanation = run_query(query, final_df)

                st.markdown(
                    f'<div class="bn bn-info">🔎 &nbsp;{explanation}</div>',
                    unsafe_allow_html=True,
                )

                if filtered.empty:
                    st.markdown(
                        '<div class="bn bn-warn">No records matched. '
                        'Try rephrasing or use the example queries above.</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        f'<div style="font-size:13px;font-weight:600;color:var(--green);'
                        f'margin-bottom:10px;">{len(filtered)} record(s) found</div>',
                        unsafe_allow_html=True,
                    )
                    st.dataframe(filtered, use_container_width=True, hide_index=True)

        except Exception as e:
            st.markdown(
                f'<div class="bn bn-err">✗ &nbsp;{str(e)}</div>',
                unsafe_allow_html=True,
            )

# Close main-wrap
st.markdown('</div>', unsafe_allow_html=True)
