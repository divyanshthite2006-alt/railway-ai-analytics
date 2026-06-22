"""
batch_processor.py — Combine results from multiple uploaded Excel files
into one unified result dict, suitable for building a merged report package.
"""

import pandas as pd
from modules.processor import process_dataframe


def merge_results(results: list[dict]) -> dict:
    """
    Merge a list of per-file result dicts into a single combined result.
    Used by the batch upload feature in the Streamlit app.
    """
    if not results:
        return {}

    combined_final    = []
    combined_audit    = []
    combined_exception= []
    total_rows        = 0
    total_tested      = 0
    total_review      = 0
    total_exec_time   = 0.0
    all_stations      = set()
    all_points        = set()
    filenames         = []

    # Keep track of (station, point, position) keys to deduplicate across files
    seen_keys: dict[tuple, int] = {}

    for r in results:
        filenames.append(r.get("filename", "—"))
        total_rows     += r["rows_processed"]
        total_exec_time+= r["execution_time"]

        for row in r["final_rows"]:
            key = (row["Station"], row["Point"], row["Position"])
            if key in seen_keys:
                # Merge: if either is Tested, mark Tested
                idx = seen_keys[key]
                existing = combined_final[idx]
                if row["Status"] == "Tested":
                    existing["Status"] = "Tested"
                    existing["Reason"] = "Valid timestamps found (merged from multiple files)"
                    # Merge dates
                    existing_dates = set(existing["Dates"].split(",")) if existing["Dates"] else set()
                    new_dates = set(row["Dates"].split(",")) if row["Dates"] else set()
                    merged_dates = sorted(existing_dates | new_dates - {""}, key=lambda x: int(x) if x.isdigit() else 0)
                    existing["Dates"] = ",".join(merged_dates)
            else:
                seen_keys[key] = len(combined_final)
                combined_final.append(dict(row))

            all_stations.add(row["Station"])
            all_points.add((row["Station"], row["Point"]))

        for row in r["audit_rows"]:
            combined_audit.append(dict(row))

        for row in r["exception_rows"]:
            combined_exception.append(dict(row))

    # Recount after dedup
    total_tested = sum(1 for r in combined_final if r["Status"] == "Tested")
    total_review = sum(1 for r in combined_final if r["Status"] == "Manual Review Required")

    # Rebuild exception list from final (since merges may promote some)
    combined_exception = [r for r in combined_final if r["Status"] == "Manual Review Required"]

    from datetime import datetime
    return {
        "final_rows":         combined_final,
        "audit_rows":         combined_audit,
        "exception_rows":     combined_exception,
        "generated_on":       datetime.now().strftime("%d-%b-%Y %H:%M"),
        "filename":           ", ".join(filenames),
        "rows_processed":     total_rows,
        "tested_count":       total_tested,
        "review_count":       total_review,
        "stations_processed": len(all_stations),
        "points_processed":   len(all_points),
        "execution_time":     round(total_exec_time, 3),
    }


def process_multiple_files(uploaded_files: list) -> dict:
    """
    Accept a list of Streamlit UploadedFile objects.
    Returns a merged result dict ready for report building.
    """
    results = []
    for uf in uploaded_files:
        try:
            df = pd.read_excel(uf, header=4)
            result = process_dataframe(df, filename=uf.name)
            results.append(result)
        except Exception as e:
            # Log and skip bad files; don't crash the whole batch
            print(f"Batch: skipping {uf.name} — {e}")
    return merge_results(results)
