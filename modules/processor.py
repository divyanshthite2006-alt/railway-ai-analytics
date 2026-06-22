"""
processor.py — Core data extraction and classification engine.
Handles all parsing, validation, and result assembly.
"""

import re
import time as time_module
from datetime import datetime, time
from typing import Optional


START_TIME = time(6, 0, 0)
END_TIME = time(18, 0, 0)

TIMESTAMP_PATTERN = re.compile(
    r'\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}:\d{3}'
)
POINT_PATTERN = re.compile(r'(\d+)')


def extract_point(message: str) -> tuple[Optional[str], Optional[str]]:
    """
    Returns (point_number, error_reason) — one will always be None.
    """
    match = POINT_PATTERN.search(message)
    if match:
        return match.group(1), None
    return None, "Point extraction failed"


def extract_position(message: str) -> tuple[Optional[str], Optional[str]]:
    """
    Returns (position, error_reason) — one will always be None.
    """
    upper = message.upper()
    if "NORMAL" in upper:
        return "Normal", None
    if "REVERSE" in upper:
        return "Reverse", None
    return None, "Position not found in message"


def parse_timestamps(time_text: str) -> tuple[list[datetime], list[str]]:
    """
    Returns (parsed_datetimes, raw_strings).
    Parses all timestamp strings found in time_text.
    """
    raw_matches = TIMESTAMP_PATTERN.findall(time_text)
    parsed = []
    for ts in raw_matches:
        try:
            clean = ts[:-4]  # strip milliseconds
            dt = datetime.strptime(clean, "%m/%d/%Y %H:%M:%S")
            parsed.append(dt)
        except ValueError:
            pass
    return parsed, raw_matches


def filter_valid_timestamps(timestamps: list[datetime]) -> list[datetime]:
    """Keep only timestamps falling within operational window (06:00–18:00)."""
    return [dt for dt in timestamps if START_TIME <= dt.time() <= END_TIME]


def process_row(row) -> dict:
    """
    Process a single DataFrame row.
    Returns a result dict with all fields needed for every report sheet.
    """
    result = {
        "station": None,
        "point": None,
        "position": None,
        "raw_timestamp_count": 0,
        "valid_timestamp_count": 0,
        "valid_dates": [],
        "status": "Manual Review Required",
        "reason": "",
        "parsing_errors": [],
    }

    # ── Station ──────────────────────────────────────────────
    station = str(row.get("STATION", "")).strip()
    if not station or station.lower() == "nan":
        result["reason"] = "Missing station"
        result["parsing_errors"].append("Missing station field")
        return result
    result["station"] = station

    # ── Message ───────────────────────────────────────────────
    message = str(row.get("FAULT MESSAGE", "")).strip()
    if not message or message.lower() == "nan":
        result["reason"] = "Missing fault message"
        result["parsing_errors"].append("Missing fault message")
        return result

    # ── Point ─────────────────────────────────────────────────
    point, point_err = extract_point(message)
    if point_err:
        result["reason"] = point_err
        result["parsing_errors"].append(point_err)
        return result
    result["point"] = point

    # ── Position ──────────────────────────────────────────────
    position, pos_err = extract_position(message)
    if pos_err:
        result["reason"] = pos_err
        result["parsing_errors"].append(pos_err)
        return result
    result["position"] = position

    # ── Timestamps ────────────────────────────────────────────
    time_text = str(row.get("TIMEDETAILS", "")).strip()
    if not time_text or time_text.lower() == "nan":
        result["reason"] = "Missing timestamp details"
        result["parsing_errors"].append("Missing TIMEDETAILS field")
        return result

    parsed, raw_matches = parse_timestamps(time_text)
    result["raw_timestamp_count"] = len(raw_matches)

    if len(raw_matches) == 0:
        result["reason"] = "No timestamps found in record"
        return result

    if len(parsed) == 0:
        result["reason"] = "Timestamp parsing failed"
        result["parsing_errors"].append("All timestamps failed to parse")
        return result

    valid = filter_valid_timestamps(parsed)
    result["valid_timestamp_count"] = len(valid)

    if not valid:
        result["reason"] = "No valid timestamp found (outside 06:00–18:00 window)"
        return result

    # ── Success ───────────────────────────────────────────────
    dates = sorted({dt.strftime("%d") for dt in valid}, key=int)
    result["valid_dates"] = dates
    result["status"] = "Tested"
    result["reason"] = "Valid timestamps found"
    return result


def process_dataframe(df, filename: str = "") -> dict:
    """
    Process entire DataFrame. Returns a dict of all result structures
    needed to build every report.
    """
    t_start = time_module.time()
    generated_on = datetime.now().strftime("%d-%b-%Y %H:%M")

    # Accumulate per-key (station, point, position) results
    keyed: dict[tuple, dict] = {}

    for _, row in df.iterrows():
        r = process_row(row)

        station = r["station"]
        point = r["point"]
        position = r["position"]

        # Rows that failed before key extraction go to a special bucket
        if station is None or point is None or position is None:
            key = (
                station or "UNKNOWN",
                point or "UNKNOWN",
                position or "UNKNOWN",
            )
            if key not in keyed:
                keyed[key] = {
                    "station": station or "UNKNOWN",
                    "point": point or "UNKNOWN",
                    "position": position or "UNKNOWN",
                    "raw_ts_total": 0,
                    "valid_ts_total": 0,
                    "valid_dates": set(),
                    "status": "Manual Review Required",
                    "reason": r["reason"],
                    "parsing_errors": r["parsing_errors"][:],
                }
            else:
                existing = keyed[key]
                existing["parsing_errors"].extend(r["parsing_errors"])
            continue

        key = (station, point, position)
        if key not in keyed:
            keyed[key] = {
                "station": station,
                "point": point,
                "position": position,
                "raw_ts_total": 0,
                "valid_ts_total": 0,
                "valid_dates": set(),
                "status": "Manual Review Required",
                "reason": r["reason"],
                "parsing_errors": r["parsing_errors"][:],
            }

        entry = keyed[key]
        entry["raw_ts_total"] += r["raw_timestamp_count"]
        entry["valid_ts_total"] += r["valid_timestamp_count"]
        entry["valid_dates"].update(r["valid_dates"])

        # Once any row for this key is Tested, promote status
        if r["status"] == "Tested":
            entry["status"] = "Tested"
            entry["reason"] = "Valid timestamps found"

    # Build flat rows
    final_rows = []
    audit_rows = []
    exception_rows = []

    for entry in keyed.values():
        dates_sorted = sorted(entry["valid_dates"], key=int) if entry["valid_dates"] else []
        dates_str = ",".join(dates_sorted)

        final_rows.append({
            "Station": entry["station"],
            "Point": entry["point"],
            "Position": entry["position"],
            "Dates": dates_str,
            "Status": entry["status"],
            "Reason": entry["reason"],
            "Generated On": generated_on,
        })

        audit_rows.append({
            "Station": entry["station"],
            "Point": entry["point"],
            "Position": entry["position"],
            "Raw Timestamp Count": entry["raw_ts_total"],
            "Valid Timestamp Count": entry["valid_ts_total"],
            "Valid Dates": dates_str,
            "Status": entry["status"],
            "Reason": entry["reason"],
        })

        if entry["status"] == "Manual Review Required":
            exception_rows.append({
                "Station": entry["station"],
                "Point": entry["point"],
                "Position": entry["position"],
                "Dates": dates_str,
                "Status": entry["status"],
                "Reason": entry["reason"],
                "Generated On": generated_on,
            })

    exec_time = round(time_module.time() - t_start, 3)
    tested_count = sum(1 for r in final_rows if r["Status"] == "Tested")
    review_count = sum(1 for r in final_rows if r["Status"] == "Manual Review Required")
    stations_set = {r["Station"] for r in final_rows if r["Station"] != "UNKNOWN"}
    points_set = {(r["Station"], r["Point"]) for r in final_rows}

    return {
        "final_rows": final_rows,
        "audit_rows": audit_rows,
        "exception_rows": exception_rows,
        "generated_on": generated_on,
        "filename": filename,
        "rows_processed": len(df),
        "tested_count": tested_count,
        "review_count": review_count,
        "stations_processed": len(stations_set),
        "points_processed": len(points_set),
        "execution_time": exec_time,
    }

