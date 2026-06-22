"""
folder_watcher.py — Standalone folder monitoring service.

Run independently:  python modules/folder_watcher.py

Watches input/ folder. On new .xlsx file:
  1. Reads and processes it
  2. Saves Excel report package to output/
  3. Saves Executive Summary PDF to output/
  4. Moves processed file to input/processed/
  5. Writes to processing log

Dependencies: watchdog  →  pip install watchdog
"""

import os
import sys
import shutil
import time
from pathlib import Path
from datetime import datetime

# Allow imports from parent directory
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    WATCHDOG_OK = True
except ImportError:
    WATCHDOG_OK = False

import pandas as pd
from modules.processor import process_dataframe
from modules.report_builder import build_excel_package, build_executive_pdf
from modules.logger import write_log

INPUT_DIR     = Path("input")
OUTPUT_DIR    = Path("output")
PROCESSED_DIR = INPUT_DIR / "processed"


def _ensure_dirs():
    for d in (INPUT_DIR, OUTPUT_DIR, PROCESSED_DIR):
        d.mkdir(parents=True, exist_ok=True)


def process_file(filepath: Path):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Detected: {filepath.name}")

    try:
        df = pd.read_excel(filepath, header=4)
    except Exception as e:
        print(f"  ✗ Failed to read file: {e}")
        return

    result = process_dataframe(df, filename=filepath.name)
    write_log(result)

    stem      = filepath.stem
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Excel package
    excel_buf = build_excel_package(result, df)
    excel_out = OUTPUT_DIR / f"{stem}_Report_{timestamp}.xlsx"
    excel_out.write_bytes(excel_buf.read())
    print(f"  ✓ Excel report saved: {excel_out.name}")

    # PDF executive summary
    pdf_buf = build_executive_pdf(result)
    pdf_out = OUTPUT_DIR / f"{stem}_Executive_Summary_{timestamp}.pdf"
    pdf_out.write_bytes(pdf_buf.read())
    print(f"  ✓ PDF summary saved: {pdf_out.name}")

    # Move processed file
    dest = PROCESSED_DIR / filepath.name
    shutil.move(str(filepath), str(dest))
    print(f"  ✓ Source moved to processed/")

    print(f"  Tested: {result['tested_count']} | Review: {result['review_count']} | "
          f"Time: {result['execution_time']}s\n")


class ExcelHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        path = Path(event.src_path)
        if path.suffix.lower() == ".xlsx":
            # Small delay to ensure file is fully written before reading
            time.sleep(1.5)
            process_file(path)


def run_watcher():
    _ensure_dirs()
    print("=" * 55)
    print("  Railway Automation — Folder Watcher")
    print(f"  Monitoring: {INPUT_DIR.resolve()}")
    print(f"  Output to:  {OUTPUT_DIR.resolve()}")
    print("  Drop .xlsx files into input/ to auto-process.")
    print("  Press Ctrl+C to stop.")
    print("=" * 55)

    event_handler = ExcelHandler()
    observer = Observer()
    observer.schedule(event_handler, str(INPUT_DIR), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(2)
    except KeyboardInterrupt:
        observer.stop()
        print("\nWatcher stopped.")
    observer.join()


if __name__ == "__main__":
    if not WATCHDOG_OK:
        print("watchdog not installed. Run:  pip install watchdog")
        sys.exit(1)
    run_watcher()
