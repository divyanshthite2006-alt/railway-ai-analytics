"""
app.py — Railway Point Testing Automation System
SECR Enterprise Edition — Full UI/UX Upgrade
All business logic unchanged. UI layer completely redesigned.
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

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

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="SECR · Point Testing Automation",
    page_icon="🚆",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════════
# DESIGN SYSTEM — injected once, governs everything
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Rajdhani:wght@500;600;700&family=Inter:wght@300;400;500;600&display=swap" rel="stylesheet">

<style>
/* ── Reset & Base ─────────────────────────────────────────────────────── */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --navy:        #0A1628;
  --navy-mid:    #152644;
  --navy-light:  #1E3A5F;
  --amber:       #F59E0B;
  --amber-dark:  #D97706;
  --cloud:       #F0F4F8;
  --white:       #FFFFFF;
  --steel:       #CBD5E1;
  --steel-dark:  #94A3B8;
  --text-pri:    #1E293B;
  --text-muted:  #64748B;
  --green:       #10B981;
  --green-bg:    #ECFDF5;
  --red:         #EF4444;
  --red-bg:      #FEF2F2;
  --amber-bg:    #FFFBEB;
  --radius-sm:   6px;
  --radius-md:   12px;
  --radius-lg:   18px;
  --shadow-card: 0 1px 3px rgba(10,22,40,0.08), 0 4px 16px rgba(10,22,40,0.06);
  --shadow-lift: 0 4px 20px rgba(10,22,40,0.14);
  --font-display: 'Rajdhani', sans-serif;
  --font-body:    'Inter', sans-serif;
  --transition:   all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
}

/* ── Global body ─────────────────────────────────────────────────────── */
html, body, [class*="css"] {
  font-family: var(--font-body) !important;
  background-color: var(--cloud) !important;
  color: var(--text-pri) !important;
}

/* ── Hide Streamlit chrome ───────────────────────────────────────────── */
#MainMenu, footer, header { visibility: hidden; }
.block-container {
  padding: 0 !important;
  max-width: 100% !important;
}
.stApp { background: var(--cloud) !important; }

/* ── Sidebar ─────────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
  background: var(--navy) !important;
  border-right: 1px solid rgba(245,158,11,0.15) !important;
  width: 260px !important;
}
[data-testid="stSidebar"] > div:first-child {
  padding: 0 !important;
}
section[data-testid="stSidebar"] .block-container {
  padding: 0 !important;
}

/* ── Sidebar radio → custom nav pills ───────────────────────────────── */
[data-testid="stSidebar"] .stRadio > div {
  gap: 4px !important;
  padding: 8px 12px !important;
}
[data-testid="stSidebar"] .stRadio label {
  background: transparent !important;
  border-radius: var(--radius-sm) !important;
  padding: 10px 14px !important;
  color: var(--steel) !important;
  font-family: var(--font-body) !important;
  font-size: 14px !important;
  font-weight: 500 !important;
  transition: var(--transition) !important;
  border: 1px solid transparent !important;
  cursor: pointer !important;
  width: 100% !important;
  display: flex !important;
  align-items: center !important;
}
[data-testid="stSidebar"] .stRadio label:hover {
  background: var(--navy-mid) !important;
  color: var(--white) !important;
  border-color: rgba(245,158,11,0.3) !important;
}
[data-testid="stSidebar"] .stRadio [data-testid="stMarkdownContainer"] p {
  color: inherit !important;
  font-size: 14px !important;
}
[data-testid="stSidebar"] .stRadio input[type="radio"]:checked + div label,
[data-testid="stSidebar"] .stRadio input[type="radio"]:checked ~ label {
  background: var(--navy-light) !important;
  color: var(--amber) !important;
  border-color: var(--amber) !important;
  font-weight: 600 !important;
}
/* Hide the actual radio dot */
[data-testid="stSidebar"] .stRadio input[type="radio"] { display: none !important; }
[data-testid="stSidebar"] .stRadio > label { display: none !important; }

/* ── Main content wrapper ─────────────────────────────────────────────── */
.main-content {
  padding: 0 32px 40px 32px;
  max-width: 1400px;
  margin: 0 auto;
}

/* ── Page header ─────────────────────────────────────────────────────── */
.page-header {
  background: var(--navy);
  padding: 22px 32px 20px;
  margin-bottom: 28px;
  position: relative;
  overflow: hidden;
  border-bottom: 3px solid var(--amber);
}
.page-header-inner {
  display: flex;
  align-items: center;
  justify-content: space-between;
  position: relative;
  z-index: 2;
}
.page-header-title {
  font-family: var(--font-display);
  font-size: 26px;
  font-weight: 700;
  color: var(--white);
  letter-spacing: 0.5px;
  line-height: 1.2;
}
.page-header-sub {
  font-family: var(--font-body);
  font-size: 13px;
  color: var(--steel-dark);
  margin-top: 3px;
  font-weight: 400;
}
.page-header-badge {
  background: rgba(245,158,11,0.15);
  border: 1px solid rgba(245,158,11,0.4);
  border-radius: 20px;
  padding: 5px 14px;
  font-size: 12px;
  font-weight: 600;
  color: var(--amber);
  font-family: var(--font-body);
  letter-spacing: 0.5px;
  text-transform: uppercase;
}

/* Animated track SVG overlay in header */
.track-animation {
  position: absolute;
  bottom: 0; left: 0; right: 0;
  height: 6px;
  overflow: hidden;
  z-index: 1;
}
.track-line {
  position: absolute;
  bottom: 3px;
  height: 2px;
  background: repeating-linear-gradient(
    90deg,
    rgba(245,158,11,0.6) 0px,
    rgba(245,158,11,0.6) 18px,
    transparent 18px,
    transparent 32px
  );
  width: 200%;
  animation: trackMove 3s linear infinite;
}
@keyframes trackMove {
  from { transform: translateX(0); }
  to   { transform: translateX(-50%); }
}

/* ── Section header (inside content) ─────────────────────────────────── */
.section-title {
  font-family: var(--font-display);
  font-size: 20px;
  font-weight: 700;
  color: var(--navy);
  letter-spacing: 0.3px;
  margin-bottom: 4px;
  display: flex;
  align-items: center;
  gap: 10px;
}
.section-title-icon {
  width: 28px; height: 28px;
  background: var(--amber);
  border-radius: 6px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  font-size: 14px;
}
.section-sub {
  font-size: 13px;
  color: var(--text-muted);
  margin-bottom: 16px;
  margin-left: 38px;
}

/* ── Cards ───────────────────────────────────────────────────────────── */
.card {
  background: var(--white);
  border-radius: var(--radius-md);
  border: 1px solid #E2E8F0;
  padding: 20px 22px;
  box-shadow: var(--shadow-card);
  margin-bottom: 20px;
  transition: var(--transition);
}
.card:hover { box-shadow: var(--shadow-lift); }
.card-accent {
  border-left: 4px solid var(--amber) !important;
}

/* ── Metric cards ─────────────────────────────────────────────────────── */
.metric-grid {
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 14px;
  margin-bottom: 24px;
}
.metric-card {
  background: var(--white);
  border-radius: var(--radius-md);
  padding: 18px 16px;
  border: 1px solid #E2E8F0;
  box-shadow: var(--shadow-card);
  text-align: center;
  transition: var(--transition);
}
.metric-card:hover { transform: translateY(-2px); box-shadow: var(--shadow-lift); }
.metric-card.highlight { border-top: 3px solid var(--amber); }
.metric-card.success   { border-top: 3px solid var(--green); }
.metric-card.alert     { border-top: 3px solid var(--red); }
.metric-value {
  font-family: var(--font-display);
  font-size: 36px;
  font-weight: 700;
  color: var(--navy);
  line-height: 1.1;
}
.metric-value.green { color: var(--green); }
.metric-value.red   { color: var(--red); }
.metric-value.amber { color: var(--amber-dark); }
.metric-label {
  font-size: 12px;
  font-weight: 500;
  color: var(--text-muted);
  text-transform: uppercase;
  letter-spacing: 0.6px;
  margin-top: 4px;
}
.metric-exec {
  font-size: 12px;
  color: var(--text-muted);
  margin-top: 4px;
  display: flex;
  align-items: center;
  gap: 6px;
}
.exec-dot {
  width: 6px; height: 6px;
  border-radius: 50%;
  background: var(--green);
  display: inline-block;
  animation: pulse 2s ease-in-out infinite;
}
@keyframes pulse {
  0%, 100% { opacity: 1; transform: scale(1); }
  50%       { opacity: 0.5; transform: scale(0.85); }
}

/* ── Upload area ─────────────────────────────────────────────────────── */
.upload-zone {
  background: var(--white);
  border: 2px dashed var(--steel);
  border-radius: var(--radius-lg);
  padding: 40px 24px;
  text-align: center;
  transition: var(--transition);
  margin-bottom: 20px;
}
.upload-zone:hover {
  border-color: var(--amber);
  background: var(--amber-bg);
}
.upload-icon {
  font-size: 40px;
  margin-bottom: 12px;
  display: block;
}
.upload-title {
  font-family: var(--font-display);
  font-size: 18px;
  font-weight: 600;
  color: var(--navy);
  margin-bottom: 6px;
}
.upload-sub {
  font-size: 13px;
  color: var(--text-muted);
}

/* ── Status badges ───────────────────────────────────────────────────── */
.badge {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  padding: 3px 10px;
  border-radius: 20px;
  font-size: 12px;
  font-weight: 600;
}
.badge-green { background: var(--green-bg); color: #065F46; }
.badge-red   { background: var(--red-bg);   color: #991B1B; }
.badge-amber { background: var(--amber-bg); color: #92400E; }

/* ── Info banners ─────────────────────────────────────────────────────── */
.banner {
  padding: 12px 16px;
  border-radius: var(--radius-sm);
  font-size: 13.5px;
  font-weight: 500;
  margin-bottom: 16px;
  display: flex;
  align-items: flex-start;
  gap: 10px;
  line-height: 1.5;
}
.banner-info  { background: #EFF6FF; color: #1D4ED8; border-left: 3px solid #3B82F6; }
.banner-ok    { background: var(--green-bg); color: #065F46; border-left: 3px solid var(--green); }
.banner-warn  { background: var(--amber-bg); color: #92400E; border-left: 3px solid var(--amber); }
.banner-error { background: var(--red-bg); color: #991B1B; border-left: 3px solid var(--red); }

/* ── Download button strip ───────────────────────────────────────────── */
.download-strip {
  background: var(--navy);
  border-radius: var(--radius-md);
  padding: 20px 24px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  margin-top: 20px;
  flex-wrap: wrap;
}
.download-strip-label {
  font-family: var(--font-display);
  font-size: 16px;
  font-weight: 600;
  color: var(--white);
}
.download-strip-sub {
  font-size: 12px;
  color: var(--steel-dark);
  margin-top: 2px;
}

/* ── Streamlit widget overrides ──────────────────────────────────────── */

/* Primary buttons → amber CTA */
.stButton > button[kind="primary"],
.stButton > button[data-testid*="primary"] {
  background: var(--amber) !important;
  color: var(--navy) !important;
  border: none !important;
  border-radius: var(--radius-sm) !important;
  font-family: var(--font-body) !important;
  font-weight: 600 !important;
  font-size: 14px !important;
  padding: 10px 22px !important;
  transition: var(--transition) !important;
  box-shadow: 0 2px 8px rgba(245,158,11,0.35) !important;
}
.stButton > button[kind="primary"]:hover {
  background: var(--amber-dark) !important;
  box-shadow: 0 4px 16px rgba(245,158,11,0.45) !important;
  transform: translateY(-1px) !important;
}

/* Secondary buttons */
.stButton > button:not([kind="primary"]) {
  background: var(--white) !important;
  color: var(--navy) !important;
  border: 1.5px solid var(--steel) !important;
  border-radius: var(--radius-sm) !important;
  font-weight: 500 !important;
  font-size: 13px !important;
  transition: var(--transition) !important;
}
.stButton > button:not([kind="primary"]):hover {
  border-color: var(--navy) !important;
  box-shadow: 0 2px 8px rgba(10,22,40,0.12) !important;
}

/* Download buttons */
.stDownloadButton > button {
  background: var(--white) !important;
  color: var(--navy) !important;
  border: 1.5px solid #E2E8F0 !important;
  border-radius: var(--radius-sm) !important;
  font-weight: 600 !important;
  font-size: 13px !important;
  padding: 10px 18px !important;
  transition: var(--transition) !important;
}
.stDownloadButton > button:hover {
  border-color: var(--amber) !important;
  color: var(--amber-dark) !important;
  box-shadow: 0 2px 12px rgba(245,158,11,0.2) !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
  background: var(--white) !important;
  border: 2px dashed var(--steel) !important;
  border-radius: var(--radius-lg) !important;
  padding: 8px !important;
  transition: var(--transition) !important;
}
[data-testid="stFileUploader"]:hover {
  border-color: var(--amber) !important;
}
[data-testid="stFileUploaderDropzone"] {
  background: transparent !important;
}

/* Text input */
.stTextInput > div > div > input {
  border: 1.5px solid #E2E8F0 !important;
  border-radius: var(--radius-sm) !important;
  font-family: var(--font-body) !important;
  font-size: 14px !important;
  padding: 10px 14px !important;
  transition: var(--transition) !important;
  background: var(--white) !important;
}
.stTextInput > div > div > input:focus {
  border-color: var(--amber) !important;
  box-shadow: 0 0 0 3px rgba(245,158,11,0.15) !important;
  outline: none !important;
}

/* Expanders */
.streamlit-expanderHeader {
  background: var(--white) !important;
  border: 1px solid #E2E8F0 !important;
  border-radius: var(--radius-sm) !important;
  font-family: var(--font-body) !important;
  font-weight: 600 !important;
  font-size: 14px !important;
  color: var(--navy) !important;
  padding: 12px 16px !important;
}
.streamlit-expanderHeader:hover {
  border-color: var(--amber) !important;
  background: var(--amber-bg) !important;
}
.streamlit-expanderContent {
  border: 1px solid #E2E8F0 !important;
  border-top: none !important;
  border-radius: 0 0 var(--radius-sm) var(--radius-sm) !important;
  background: var(--white) !important;
  padding: 16px !important;
}

/* Dataframe */
[data-testid="stDataFrame"] {
  border-radius: var(--radius-md) !important;
  overflow: hidden !important;
  border: 1px solid #E2E8F0 !important;
  box-shadow: var(--shadow-card) !important;
}

/* Metrics — override Streamlit's default */
[data-testid="stMetric"] {
  background: var(--white) !important;
  border: 1px solid #E2E8F0 !important;
  border-radius: var(--radius-md) !important;
  padding: 16px !important;
  box-shadow: var(--shadow-card) !important;
}
[data-testid="stMetricLabel"] {
  font-family: var(--font-body) !important;
  font-size: 12px !important;
  text-transform: uppercase !important;
  letter-spacing: 0.5px !important;
  color: var(--text-muted) !important;
}
[data-testid="stMetricValue"] {
  font-family: var(--font-display) !important;
  font-size: 30px !important;
  font-weight: 700 !important;
  color: var(--navy) !important;
}

/* Spinner */
.stSpinner > div {
  border-top-color: var(--amber) !important;
}

/* Success / error / info messages */
.stSuccess {
  background: var(--green-bg) !important;
  color: #065F46 !important;
  border-left: 3px solid var(--green) !important;
  border-radius: var(--radius-sm) !important;
  font-weight: 500 !important;
}
.stError {
  background: var(--red-bg) !important;
  color: #991B1B !important;
  border-left: 3px solid var(--red) !important;
  border-radius: var(--radius-sm) !important;
}
.stInfo {
  background: #EFF6FF !important;
  color: #1D4ED8 !important;
  border-left: 3px solid #3B82F6 !important;
  border-radius: var(--radius-sm) !important;
}
.stWarning {
  background: var(--amber-bg) !important;
  color: #92400E !important;
  border-left: 3px solid var(--amber) !important;
  border-radius: var(--radius-sm) !important;
}

/* Divider */
hr {
  border: none !important;
  border-top: 1px solid #E2E8F0 !important;
  margin: 24px 0 !important;
}

/* Caption / small text */
.stCaption, .stCaption p {
  color: var(--text-muted) !important;
  font-size: 12px !important;
}

/* ── Sidebar logo block ───────────────────────────────────────────────── */
.sidebar-logo {
  padding: 24px 16px 16px;
  border-bottom: 1px solid rgba(255,255,255,0.08);
  margin-bottom: 8px;
}
.sidebar-logo-mark {
  width: 44px; height: 44px;
  background: var(--amber);
  border-radius: 10px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 22px;
  margin-bottom: 12px;
}
.sidebar-org {
  font-family: var(--font-display);
  font-size: 18px;
  font-weight: 700;
  color: var(--white);
  line-height: 1.2;
  letter-spacing: 0.3px;
}
.sidebar-dept {
  font-size: 11px;
  color: var(--steel-dark);
  margin-top: 2px;
  font-weight: 400;
  letter-spacing: 0.4px;
}
.sidebar-nav-label {
  font-size: 10px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 1px;
  color: var(--steel-dark);
  padding: 12px 16px 6px;
}
.sidebar-footer {
  position: absolute;
  bottom: 0; left: 0; right: 0;
  padding: 14px 16px;
  border-top: 1px solid rgba(255,255,255,0.07);
}
.sidebar-footer-text {
  font-size: 11px;
  color: rgba(203,213,225,0.5);
  text-align: center;
}

/* ── Query chip buttons ──────────────────────────────────────────────── */
.chip {
  display: inline-block;
  background: #F1F5F9;
  border: 1px solid #E2E8F0;
  border-radius: 20px;
  padding: 4px 12px;
  font-size: 12px;
  font-weight: 500;
  color: var(--navy);
  margin: 3px 3px 3px 0;
  cursor: pointer;
  transition: var(--transition);
}
.chip:hover {
  background: var(--amber-bg);
  border-color: var(--amber);
  color: var(--amber-dark);
}

/* ── Table tag pills ─────────────────────────────────────────────────── */
.tag-tested {
  background: #D1FAE5;
  color: #065F46;
  padding: 2px 8px;
  border-radius: 12px;
  font-size: 11px;
  font-weight: 600;
}
.tag-review {
  background: #FEE2E2;
  color: #991B1B;
  padding: 2px 8px;
  border-radius: 12px;
  font-size: 11px;
  font-weight: 600;
}

/* ── Scrollbar ───────────────────────────────────────────────────────── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #F1F5F9; }
::-webkit-scrollbar-thumb { background: var(--steel); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: var(--navy); }

/* ── Responsive ──────────────────────────────────────────────────────── */
@media (max-width: 900px) {
  .metric-grid { grid-template-columns: repeat(2, 1fr); }
  .main-content { padding: 0 16px 32px; }
  .download-strip { flex-direction: column; }
}

/* ── Log table styling ───────────────────────────────────────────────── */
.log-stat {
  font-family: var(--font-display);
  font-size: 28px;
  font-weight: 700;
  color: var(--navy);
}
.log-stat-label {
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  color: var(--text-muted);
}

</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">
      <div class="sidebar-logo-mark">🚆</div>
      <div class="sidebar-org">SECR</div>
      <div class="sidebar-dept">Signal &amp; Telecommunication Dept.</div>
    </div>
    <div class="sidebar-nav-label">Modules</div>
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

    st.markdown("""
    <div style="height:24px;"></div>
    <div style="padding: 0 16px;">
      <div style="background:rgba(245,158,11,0.08);border:1px solid rgba(245,158,11,0.2);
                  border-radius:8px;padding:12px 14px;">
        <div style="font-size:11px;color:#94A3B8;text-transform:uppercase;
                    letter-spacing:0.6px;margin-bottom:6px;">Operational Window</div>
        <div style="font-family:'Rajdhani',sans-serif;font-size:20px;
                    font-weight:700;color:#F59E0B;">06:00 – 18:00</div>
        <div style="font-size:11px;color:#64748B;margin-top:2px;">
          Valid testing hours
        </div>
      </div>
    </div>
    <div class="sidebar-footer">
      <div class="sidebar-footer-text">
        South East Central Railway · Raipur Division<br>
        Point Testing Automation v2.0
      </div>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# HEADER — appears on every page
# ══════════════════════════════════════════════════════════════════════════════
PAGE_META = {
    "🗂️  Single File":     ("Single File Processing",   "Upload and process one Excel file at a time"),
    "📦  Batch Processing": ("Batch Processing",          "Process multiple files and merge results"),
    "📜  Processing Logs":  ("Processing Logs",           "Audit trail of all processing runs"),
    "🔍  AI Query Engine":  ("AI Query Engine",           "Query your results in plain English"),
}
title_text, sub_text = PAGE_META.get(mode, ("Point Testing Automation", ""))

st.markdown(f"""
<div class="page-header">
  <div class="page-header-inner">
    <div>
      <div class="page-header-title">{title_text}</div>
      <div class="page-header-sub">{sub_text}</div>
    </div>
    <div class="page-header-badge">SECR · S&amp;T Dept.</div>
  </div>
  <div class="track-animation"><div class="track-line"></div></div>
</div>
<div class="main-content">
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def render_metrics(result: dict):
    """Custom metric grid — replaces st.metric."""
    t = result["tested_count"]
    r = result["review_count"]
    total = t + r
    pct = f"{round(t/total*100)}%" if total else "—"

    st.markdown(f"""
    <div class="metric-grid">
      <div class="metric-card highlight">
        <div class="metric-value">{result['rows_processed']}</div>
        <div class="metric-label">Records Processed</div>
      </div>
      <div class="metric-card">
        <div class="metric-value">{result['stations_processed']}</div>
        <div class="metric-label">Stations</div>
      </div>
      <div class="metric-card">
        <div class="metric-value">{result['points_processed']}</div>
        <div class="metric-label">Points</div>
      </div>
      <div class="metric-card success">
        <div class="metric-value green">{t}</div>
        <div class="metric-label">Tested</div>
      </div>
      <div class="metric-card alert">
        <div class="metric-value red">{r}</div>
        <div class="metric-label">Manual Review</div>
      </div>
    </div>
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:20px;
                font-size:12px;color:var(--text-muted);">
      <span class="exec-dot"></span>
      Execution time <strong style="color:var(--navy);">{result['execution_time']}s</strong>
      &nbsp;·&nbsp; Generated <strong style="color:var(--navy);">{result['generated_on']}</strong>
      &nbsp;·&nbsp; Compliance rate <strong style="color:#10B981;">{pct}</strong>
    </div>
    """, unsafe_allow_html=True)


def render_section(icon_emoji: str, title: str, subtitle: str = ""):
    sub_html = f'<div class="section-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(f"""
    <div style="margin-bottom:6px;">
      <div class="section-title">
        <span class="section-title-icon">{icon_emoji}</span>
        {title}
      </div>
      {sub_html}
    </div>
    """, unsafe_allow_html=True)


def show_station_summary(result: dict):
    render_section("📊", "Station Summary", "Point testing status grouped by station")
    summary_df = build_station_summary(result["final_rows"])
    if summary_df.empty:
        st.markdown('<div class="banner banner-info">ℹ️ No station data available.</div>', unsafe_allow_html=True)
        return
    st.dataframe(summary_df, use_container_width=True, hide_index=True)


def show_downloads(result: dict, df_raw: pd.DataFrame):
    excel_buf = build_excel_package(result, df_raw)
    pdf_buf   = build_executive_pdf(result)

    st.markdown("""
    <div style="margin-bottom:8px;">
      <div class="section-title"><span class="section-title-icon">⬇️</span>Download Report Package</div>
      <div class="section-sub">All reports are generated and ready for download</div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="card card-accent">'
                    '<div style="font-weight:600;color:var(--navy);margin-bottom:4px;">📊 Full Report Package</div>'
                    '<div style="font-size:12px;color:var(--text-muted);margin-bottom:12px;">'
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
        st.markdown('<div class="card card-accent">'
                    '<div style="font-weight:600;color:var(--navy);margin-bottom:4px;">📄 Executive Summary</div>'
                    '<div style="font-size:12px;color:var(--text-muted);margin-bottom:12px;">'
                    'PDF report suitable for senior officers and documentation</div>',
                    unsafe_allow_html=True)
        st.download_button(
            "Download .pdf Summary",
            data=pdf_buf,
            file_name="SECR_Executive_Summary.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# MODE 1 — SINGLE FILE
# ══════════════════════════════════════════════════════════════════════════════
if mode == "🗂️  Single File":

    # Upload card
    st.markdown('<div class="card">', unsafe_allow_html=True)
    render_section("📁", "Upload File", "Accepts .xlsx files exported from the railway point testing system")
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
                f'<div class="banner banner-ok">✓ &nbsp;<strong>{uploaded_file.name}</strong> '
                f'loaded — {len(df)} rows detected</div>',
                unsafe_allow_html=True
            )

            with st.expander("Preview Input Data", expanded=False):
                st.dataframe(df.head(10), use_container_width=True, hide_index=True)

            st.markdown('<div style="height:4px;"></div>', unsafe_allow_html=True)

            if st.button("⚙️  Generate Report", type="primary", use_container_width=False):
                with st.spinner("Processing records..."):
                    result = process_dataframe(df, filename=uploaded_file.name)
                    write_log(result)

                st.markdown('<hr>', unsafe_allow_html=True)

                # Metrics
                render_section("⚡", "Execution Metrics")
                render_metrics(result)

                st.markdown('<hr>', unsafe_allow_html=True)

                # Final report
                render_section("📋", "Final Report", "All classified point-position combinations")
                final_df = pd.DataFrame(result["final_rows"])
                st.dataframe(
                    final_df.style.apply(
                        lambda row: [
                            "background-color: #C6EFCE" if row["Status"] == "Tested"
                            else "background-color: #FFC7CE"
                            for _ in row
                        ], axis=1,
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

                st.markdown('<hr>', unsafe_allow_html=True)

                # Collapsible sub-reports
                with st.expander(f"🔍  Audit Report — Verification Trail"):
                    audit_df = pd.DataFrame(result["audit_rows"])
                    st.dataframe(audit_df, use_container_width=True, hide_index=True)

                review_n = result['review_count']
                label_exc = (
                    f"⚠️  Exception Report — {review_n} record(s) need review"
                    if review_n > 0 else
                    "✅  Exception Report — No exceptions found"
                )
                with st.expander(label_exc, expanded=review_n > 0):
                    if result["exception_rows"]:
                        exc_df = pd.DataFrame(result["exception_rows"])
                        st.dataframe(exc_df, use_container_width=True, hide_index=True)
                    else:
                        st.markdown(
                            '<div class="banner banner-ok">✓ All records classified as Tested. No exceptions.</div>',
                            unsafe_allow_html=True
                        )

                with st.expander("🧪  Data Quality Report"):
                    dq_df = build_data_quality_report(df)
                    st.dataframe(dq_df, use_container_width=True, hide_index=True)

                st.markdown('<hr>', unsafe_allow_html=True)
                show_station_summary(result)

                st.markdown('<hr>', unsafe_allow_html=True)
                show_downloads(result, df)

                st.markdown(
                    '<div class="banner banner-ok" style="margin-top:16px;">'
                    '✓ &nbsp;Report generated successfully and ready for download.</div>',
                    unsafe_allow_html=True
                )

        except Exception as e:
            st.markdown(
                f'<div class="banner banner-error">✗ &nbsp;{str(e)}</div>',
                unsafe_allow_html=True
            )


# ══════════════════════════════════════════════════════════════════════════════
# MODE 2 — BATCH PROCESSING
# ══════════════════════════════════════════════════════════════════════════════
elif mode == "📦  Batch Processing":

    st.markdown('<div class="card">', unsafe_allow_html=True)
    render_section("📦", "Batch Upload", "Select multiple .xlsx files — results are merged with deduplication")
    st.markdown(
        '<div class="banner banner-info">ℹ️ Hold <kbd>Ctrl</kbd> (or <kbd>Cmd</kbd> on Mac) '
        'while clicking to select multiple files in the file picker.</div>',
        unsafe_allow_html=True
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
            f'<div style="font-weight:600;color:var(--navy);margin-bottom:8px;">'
            f'{len(uploaded_files)} file(s) queued</div>',
            unsafe_allow_html=True
        )
        for uf in uploaded_files:
            st.markdown(
                f'<div style="font-size:13px;color:var(--text-muted);padding:3px 0;">'
                f'📄 &nbsp;{uf.name}</div>',
                unsafe_allow_html=True
            )
        st.markdown('</div>', unsafe_allow_html=True)

        if st.button(f"⚙️  Process {len(uploaded_files)} File(s)", type="primary"):
            with st.spinner(f"Processing {len(uploaded_files)} file(s) and merging results..."):
                result = process_multiple_files(uploaded_files)
                for uf in uploaded_files:
                    write_log({**result, "filename": uf.name})

            if result:
                render_section("⚡", "Batch Execution Metrics")
                render_metrics(result)
                st.markdown('<hr>', unsafe_allow_html=True)

                render_section("📋", "Combined Final Report", "Merged and deduplicated across all uploaded files")
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
                    '<div class="banner banner-ok" style="margin-top:16px;">'
                    '✓ &nbsp;Batch report generated successfully.</div>',
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    '<div class="banner banner-error">'
                    '✗ &nbsp;No valid data could be processed from the uploaded files.</div>',
                    unsafe_allow_html=True
                )


# ══════════════════════════════════════════════════════════════════════════════
# MODE 3 — PROCESSING LOGS
# ══════════════════════════════════════════════════════════════════════════════
elif mode == "📜  Processing Logs":

    logs = read_logs()

    if not logs:
        st.markdown("""
        <div class="card" style="text-align:center;padding:48px;">
          <div style="font-size:40px;margin-bottom:12px;">📭</div>
          <div style="font-family:'Rajdhani',sans-serif;font-size:20px;font-weight:700;
                      color:var(--navy);margin-bottom:6px;">No logs yet</div>
          <div style="font-size:13px;color:var(--text-muted);">
            Generate a report from Single File or Batch Processing to start logging.
          </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        logs_df = pd.DataFrame(logs)
        total_tested = logs_df["Tested Count"].astype(int).sum()
        total_review = logs_df["Manual Review Count"].astype(int).sum()
        total_rows   = logs_df["Rows Processed"].astype(int).sum()

        # Summary strip
        st.markdown(f"""
        <div class="metric-grid" style="grid-template-columns:repeat(3,1fr);">
          <div class="metric-card highlight">
            <div class="metric-value">{len(logs_df)}</div>
            <div class="metric-label">Total Runs</div>
          </div>
          <div class="metric-card">
            <div class="metric-value">{total_rows}</div>
            <div class="metric-label">Total Rows Processed</div>
          </div>
          <div class="metric-card success">
            <div class="metric-value green">{total_tested}</div>
            <div class="metric-label">Total Tested</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        render_section("📜", "Processing History", "Every report generation run — auto-saved to logs/processing_log.csv")
        st.dataframe(logs_df, use_container_width=True, hide_index=True)

        log_csv = logs_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥  Download Log as CSV",
            data=log_csv,
            file_name="SECR_processing_log.csv",
            mime="text/csv",
        )


# ══════════════════════════════════════════════════════════════════════════════
# MODE 4 — AI QUERY ENGINE
# ══════════════════════════════════════════════════════════════════════════════
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
        '<div class="banner banner-info">ℹ️ Runs entirely on Pandas — no external AI APIs. '
        'Your data never leaves your machine.</div>',
        unsafe_allow_html=True
    )

    uploaded_file = st.file_uploader(
        "Upload Excel file to query",
        type=["xlsx"],
        key="query_uploader",
        label_visibility="collapsed",
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # Example query chips
    st.markdown('<div style="margin-bottom:12px;">', unsafe_allow_html=True)
    st.markdown(
        '<div style="font-size:12px;font-weight:600;color:var(--text-muted);'
        'text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">Example queries</div>'
        + "".join(f'<span class="chip">{q}</span>' for q in EXAMPLE_QUERIES),
        unsafe_allow_html=True
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file, header=4)
            result = process_dataframe(df_raw, filename=uploaded_file.name)
            final_df = pd.DataFrame(result["final_rows"])

            st.markdown(
                f'<div class="banner banner-ok">✓ &nbsp;<strong>{uploaded_file.name}</strong> '
                f'ready — {len(final_df)} classified records available for querying.</div>',
                unsafe_allow_html=True
            )

            query = st.text_input(
                "Your query",
                placeholder="e.g.  Show all manual review points for station BLP",
                label_visibility="collapsed",
            )

            if query:
                with st.spinner("Querying..."):
                    filtered, explanation = run_query(query, final_df)

                st.markdown(
                    f'<div class="banner banner-info">🔎 &nbsp;{explanation}</div>',
                    unsafe_allow_html=True
                )

                if filtered.empty:
                    st.markdown(
                        '<div class="banner banner-warn">No records matched your query. '
                        'Try rephrasing or check the example queries above.</div>',
                        unsafe_allow_html=True
                    )
                else:
                    st.markdown(
                        f'<div style="font-size:13px;font-weight:600;color:var(--green);'
                        f'margin-bottom:10px;">{len(filtered)} record(s) found</div>',
                        unsafe_allow_html=True
                    )
                    st.dataframe(filtered, use_container_width=True, hide_index=True)

        except Exception as e:
            st.markdown(
                f'<div class="banner banner-error">✗ &nbsp;{str(e)}</div>',
                unsafe_allow_html=True
            )

# Close main-content wrapper
st.markdown('</div>', unsafe_allow_html=True)
