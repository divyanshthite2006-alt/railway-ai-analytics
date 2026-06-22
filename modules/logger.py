"""
logger.py — Append-only processing log stored as CSV in the logs/ folder.
Each run writes one row. Log file persists across sessions.
"""

import csv
import os
from datetime import datetime
from pathlib import Path

LOG_DIR  = Path("logs")
LOG_FILE = LOG_DIR / "processing_log.csv"

COLUMNS = [
    "Timestamp",
    "File Name",
    "Rows Processed",
    "Tested Count",
    "Manual Review Count",
    "Stations Processed",
    "Points Processed",
    "Execution Time (s)",
]


def _ensure_log_file():
    LOG_DIR.mkdir(exist_ok=True)
    if not LOG_FILE.exists():
        with open(LOG_FILE, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=COLUMNS)
            writer.writeheader()


def write_log(result: dict):
    """Append one processing record to the log CSV."""
    _ensure_log_file()
    row = {
        "Timestamp":              datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "File Name":              result.get("filename", "—"),
        "Rows Processed":         result.get("rows_processed", 0),
        "Tested Count":           result.get("tested_count", 0),
        "Manual Review Count":    result.get("review_count", 0),
        "Stations Processed":     result.get("stations_processed", 0),
        "Points Processed":       result.get("points_processed", 0),
        "Execution Time (s)":     result.get("execution_time", 0),
    }
    with open(LOG_FILE, "a", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writerow(row)


def read_logs() -> list[dict]:
    """Return all log rows as list of dicts. Returns empty list if no log exists."""
    _ensure_log_file()
    rows = []
    with open(LOG_FILE, "r", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(dict(row))
    return rows
