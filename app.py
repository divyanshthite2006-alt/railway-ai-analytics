"""
app.py — Railway Point Testing Automation System
Enterprise Edition — All 12 upgrades integrated.

Run:  streamlit run app.py
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

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Railway Point Testing Automation",
    page_icon="🚆",
    layout="wide",
)

st.title("🚆 Railway Point Testing Automation")

# ── Sidebar navigation ────────────────────────────────────────────────────────
st.sidebar.title("Navigation")
mode = st.sidebar.radio(
    "Mode",
    ["Single File", "Batch Processing", "Processing Logs", "AI Query Engine"],
    index=0,
)

# =============================================================================
# HELPER — show execution metrics strip
# =============================================================================
def show_metrics(result: dict):
    st.subheader("⚡ Execution Metrics")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Records Processed",    result["rows_processed"])
    c2.metric("Stations Processed",   result["stations_processed"])
    c3.metric("Points Processed",     result["points_processed"])
    c4.metric("Tested",               result["tested_count"])
    c5.metric("Manual Review",        result["review_count"])
    st.caption(f"⏱ Execution time: **{result['execution_time']}s** | Generated: {result['generated_on']}")


# =============================================================================
# HELPER — show station summary table
# =============================================================================
def show_station_summary(result: dict):
    st.subheader("📊 Station Summary")
    summary_df = build_station_summary(result["final_rows"])
    if summary_df.empty:
        st.info("No station data available.")
        return
    st.dataframe(summary_df, use_container_width=True)


# =============================================================================
# HELPER — render download buttons for the full report package
# =============================================================================
def show_downloads(result: dict, df_raw: pd.DataFrame):
    st.subheader("📥 Download Report Package")

    excel_buf = build_excel_package(result, df_raw)
    pdf_buf   = build_executive_pdf(result)

    col1, col2 = st.columns(2)
    col1.download_button(
        "📊 Download Full Report Package (.xlsx)",
        data=excel_buf,
        file_name="Railway_Report_Package.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    col2.download_button(
        "📄 Download Executive Summary (.pdf)",
        data=pdf_buf,
        file_name="Railway_Executive_Summary.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
    st.caption("The Excel package contains: Final Report · Exception Report · Audit Report · Station Summary · Data Quality Report")


# =============================================================================
# MODE 1 — Single File Processing
# =============================================================================
if mode == "Single File":

    uploaded_file = st.file_uploader(
        "Upload Excel File",
        type=["xlsx"],
    )

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file, header=4)
            st.success("File Uploaded Successfully!")

            st.subheader("Input Preview")
            st.dataframe(df.head(), use_container_width=True)

            if st.button("Generate Report", type="primary"):
                with st.spinner("Processing..."):
                    result = process_dataframe(df, filename=uploaded_file.name)
                    write_log(result)

                # ── Metrics strip ─────────────────────────────────────────
                show_metrics(result)

                st.divider()

                # ── Final Report table ────────────────────────────────────
                st.subheader("📋 Final Report")
                final_df = pd.DataFrame(result["final_rows"])
                st.dataframe(
                    final_df.style.apply(
                        lambda row: [
                            "background-color: #C6EFCE" if row["Status"] == "Tested"
                            else "background-color: #FFC7CE"
                            for _ in row
                        ],
                        axis=1,
                    ),
                    use_container_width=True,
                )

                st.divider()

                # ── Audit table ───────────────────────────────────────────
                with st.expander("🔍 Audit Report — Verification Trail", expanded=False):
                    audit_df = pd.DataFrame(result["audit_rows"])
                    st.dataframe(audit_df, use_container_width=True)

                # ── Exception table ───────────────────────────────────────
                with st.expander(
                    f"⚠️ Exception Report — {result['review_count']} record(s) requiring review",
                    expanded=result["review_count"] > 0,
                ):
                    if result["exception_rows"]:
                        exc_df = pd.DataFrame(result["exception_rows"])
                        st.dataframe(exc_df, use_container_width=True)
                    else:
                        st.success("No exceptions — all records classified as Tested.")

                # ── Data Quality ──────────────────────────────────────────
                with st.expander("🧪 Data Quality Report", expanded=False):
                    dq_df = build_data_quality_report(df)
                    st.dataframe(dq_df, use_container_width=True)

                st.divider()

                # ── Station Summary ───────────────────────────────────────
                show_station_summary(result)

                st.divider()

                # ── Downloads ─────────────────────────────────────────────
                show_downloads(result, df)

                st.success("✅ Report Generated Successfully!")

        except Exception as e:
            st.error(f"Error: {str(e)}")


# =============================================================================
# MODE 2 — Batch Processing
# =============================================================================
elif mode == "Batch Processing":

    st.subheader("📦 Batch Processing — Multiple Files")
    st.info(
        "Upload multiple Excel files at once. Results will be merged into a single "
        "combined report package with deduplication across files."
    )

    uploaded_files = st.file_uploader(
        "Upload Excel Files (select multiple)",
        type=["xlsx"],
        accept_multiple_files=True,
    )

    if uploaded_files:
        st.write(f"**{len(uploaded_files)} file(s) selected:**")
        for uf in uploaded_files:
            st.caption(f"• {uf.name}")

        if st.button("Process All Files", type="primary"):
            with st.spinner(f"Processing {len(uploaded_files)} file(s)..."):
                result = process_multiple_files(uploaded_files)
                for uf in uploaded_files:
                    write_log({**result, "filename": uf.name})

            if result:
                show_metrics(result)
                st.divider()

                st.subheader("📋 Combined Final Report")
                final_df = pd.DataFrame(result["final_rows"])
                st.dataframe(final_df, use_container_width=True)

                st.divider()
                show_station_summary(result)
                st.divider()

                # Re-read one file for raw df (used in data quality pass)
                try:
                    uploaded_files[0].seek(0)
                    df_sample = pd.read_excel(uploaded_files[0], header=4)
                except Exception:
                    df_sample = pd.DataFrame()

                show_downloads(result, df_sample)
                st.success("✅ Batch Report Generated Successfully!")
            else:
                st.error("No valid data could be processed from the uploaded files.")


# =============================================================================
# MODE 3 — Processing Logs
# =============================================================================
elif mode == "Processing Logs":

    st.subheader("📜 Processing Logs")
    st.caption("Auto-saved log of every processing run performed in this session directory.")

    logs = read_logs()

    if not logs:
        st.info("No processing runs logged yet. Generate a report to see logs here.")
    else:
        logs_df = pd.DataFrame(logs)
        st.dataframe(logs_df, use_container_width=True)

        log_csv = logs_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥 Download Processing Log (.csv)",
            data=log_csv,
            file_name="processing_log.csv",
            mime="text/csv",
        )

        # Quick summary metrics from log
        st.divider()
        st.subheader("Log Summary")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Runs",           len(logs_df))
        c2.metric("Total Rows Processed", logs_df["Rows Processed"].astype(int).sum())
        c3.metric("Total Tested",         logs_df["Tested Count"].astype(int).sum())


# =============================================================================
# MODE 4 — AI Query Engine
# =============================================================================
elif mode == "AI Query Engine":

    st.subheader("🔍 AI Query Engine")
    st.info(
        "Upload your Excel file and query the results in plain English. "
        "Implemented entirely with Pandas — no external AI APIs required."
    )

    uploaded_file = st.file_uploader(
        "Upload Excel File to Query",
        type=["xlsx"],
        key="query_uploader",
    )

    EXAMPLE_QUERIES = [
        "Show all points tested on 15",
        "Which station has most reviews?",
        "Show all manual review points",
        "Show testing activity for station DPH",
        "Show all tested points",
        "Show all points",
    ]

    st.caption("**Example queries:** " + " · ".join(f'`{q}`' for q in EXAMPLE_QUERIES))

    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file, header=4)
            result = process_dataframe(df_raw, filename=uploaded_file.name)
            final_df = pd.DataFrame(result["final_rows"])

            query = st.text_input(
                "Enter your query",
                placeholder="e.g. Show all manual review points",
            )

            if query:
                with st.spinner("Querying..."):
                    filtered, explanation = run_query(query, final_df)

                st.caption(f"🔎 {explanation}")

                if filtered.empty:
                    st.warning("No records matched your query.")
                else:
                    st.success(f"{len(filtered)} record(s) found.")
                    st.dataframe(filtered, use_container_width=True)

        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
