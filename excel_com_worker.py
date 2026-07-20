"""
excel_com_worker.py
--------------------
Windows-only COM automation worker: opens a large Excel workbook via the
Excel application object, reads specific data, processes it, and writes
a summary report (new sheet + optional .xlsx / .csv export).

Requirements:
    pip install pywin32

Why COM instead of openpyxl/pandas?
    - You need Excel's own calculation engine (formulas, pivot tables, macros).
    - The file has features openpyxl can't read/write (certain chart types,
      VBA projects, external links that must be resolved live).
    - You need to drive an already-open, human-in-the-loop workbook.

If you don't need live Excel (no formulas/macros to evaluate), prefer
pandas/openpyxl instead — it's far faster and doesn't need Excel installed.
"""

import argparse
import contextlib
import csv
import logging
import os
import sys
import time
from dataclasses import dataclass, field
from typing import Any, Iterator, Optional

try:
    import win32com.client as win32
    import pywintypes
    from win32com.client import constants as xlconst
except ImportError:
    print(
        "This script requires pywin32. Install it with:\n"
        "    pip install pywin32\n"
        "and run on Windows with Excel installed.",
        file=sys.stderr,
    )
    sys.exit(1)


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("excel_com_worker")

# Excel constants (avoid depending on gen_py cache being built)
xlUp = -4162
xlCellTypeLastCell = 11
xlWorkbookDefault = 51
xlCalculationManual = -4135
xlCalculationAutomatic = -4105


# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------

@dataclass
class WorkerConfig:
    input_path: str
    sheet_name: Optional[str] = None      # None = active/first sheet
    header_row: int = 1
    output_xlsx: Optional[str] = None     # if set, save a copy with report sheet
    output_csv: Optional[str] = None      # if set, dump the report as CSV
    visible: bool = False                 # show the Excel UI (slower, useful for debugging)
    read_only: bool = True                # open source file read-only
    columns: list = field(default_factory=list)  # specific column headers to extract; [] = all


# --------------------------------------------------------------------------
# Excel application lifecycle (context manager keeps this safe)
# --------------------------------------------------------------------------

@contextlib.contextmanager
def excel_application(visible: bool = False) -> Iterator[Any]:
    """
    Start Excel, disable alerts/screen updating for speed, and guarantee
    cleanup (quitting Excel, releasing COM refs) even if something throws.
    """
    log.info("Starting Excel application...")
    app = win32.gencache.EnsureDispatch("Excel.Application")
    app.Visible = visible
    app.DisplayAlerts = False
    app.ScreenUpdating = visible  # only bother updating the screen if it's shown
    prev_calc_mode = app.Calculation
    app.Calculation = xlCalculationManual  # big speed win on large files

    try:
        yield app
    finally:
        log.info("Cleaning up Excel application...")
        try:
            app.Calculation = prev_calc_mode
            app.DisplayAlerts = False
            app.Quit()
        except pywintypes.com_error:
            log.warning("Excel already closed or unresponsive during cleanup.")
        del app


@contextlib.contextmanager
def open_workbook(app: Any, path: str, read_only: bool = True) -> Iterator[Any]:
    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Workbook not found: {path}")

    log.info("Opening workbook: %s", path)
    wb = app.Workbooks.Open(
        path,
        ReadOnly=read_only,
        UpdateLinks=0,       # don't prompt/refresh external links
        Notify=False,
    )
    try:
        yield wb
    finally:
        try:
            wb.Close(SaveChanges=False)
        except pywintypes.com_error:
            log.warning("Workbook already closed.")


# --------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------

def get_used_range_bounds(sheet: Any) -> tuple[int, int]:
    """Return (last_row, last_col) of the sheet's used range."""
    used = sheet.UsedRange
    last_row = used.Row + used.Rows.Count - 1
    last_col = used.Column + used.Columns.Count - 1
    return last_row, last_col


def read_sheet_as_records(
    wb: Any, sheet_name: Optional[str], header_row: int, wanted_columns: list
) -> list[dict]:
    """
    Bulk-read a sheet into a list of dicts using a single COM call
    (Range.Value2) rather than cell-by-cell reads, which is orders of
    magnitude faster for large files.
    """
    sheet = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
    log.info("Reading sheet: %s", sheet.Name)

    last_row, last_col = get_used_range_bounds(sheet)
    log.info("Used range: %d rows x %d cols", last_row, last_col)

    # One round-trip: grab everything as a tuple-of-tuples.
    full_range = sheet.Range(
        sheet.Cells(header_row, 1), sheet.Cells(last_row, last_col)
    )
    data = full_range.Value2  # tuple of tuples, fast bulk transfer

    if data is None:
        log.warning("Sheet appears empty.")
        return []

    headers = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(data[0])]

    if wanted_columns:
        keep_idx = [i for i, h in enumerate(headers) if h in wanted_columns]
        if not keep_idx:
            raise ValueError(f"None of {wanted_columns} found in headers: {headers}")
    else:
        keep_idx = list(range(len(headers)))

    records = []
    for row in data[1:]:
        if row is None or all(v is None for v in row):
            continue  # skip fully blank rows
        record = {headers[i]: row[i] for i in keep_idx}
        records.append(record)

    log.info("Read %d data rows.", len(records))
    return records


# --------------------------------------------------------------------------
# Processing (customize this for your actual business logic)
# --------------------------------------------------------------------------

def process_records(records: list[dict]) -> dict:
    """
    Example processing: numeric column totals/averages + row count.
    Replace this with whatever aggregation/validation logic you need.
    """
    if not records:
        return {"row_count": 0, "numeric_summary": {}}

    numeric_summary: dict[str, dict] = {}
    for key in records[0].keys():
        values = [r[key] for r in records if isinstance(r.get(key), (int, float))]
        if values:
            numeric_summary[key] = {
                "sum": sum(values),
                "avg": sum(values) / len(values),
                "min": min(values),
                "max": max(values),
                "count": len(values),
            }

    return {
        "row_count": len(records),
        "numeric_summary": numeric_summary,
    }


# --------------------------------------------------------------------------
# Report generation
# --------------------------------------------------------------------------

def write_report_sheet(wb: Any, summary: dict, source_sheet_name: str) -> None:
    """Add a 'Report' sheet to the (in-memory) workbook object with the summary."""
    for s in wb.Sheets:
        if s.Name == "Report":
            s.Delete()
            break

    report = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
    report.Name = "Report"

    report.Cells(1, 1).Value = f"Report for sheet: {source_sheet_name}"
    report.Cells(1, 1).Font.Bold = True
    report.Cells(2, 1).Value = f"Generated: {time.strftime('%Y-%m-%d %H:%M:%S')}"
    report.Cells(3, 1).Value = f"Row count: {summary['row_count']}"

    row = 5
    report.Cells(row, 1).Value = "Column"
    report.Cells(row, 2).Value = "Sum"
    report.Cells(row, 3).Value = "Avg"
    report.Cells(row, 4).Value = "Min"
    report.Cells(row, 5).Value = "Max"
    report.Cells(row, 6).Value = "Count"
    for c in range(1, 7):
        report.Cells(row, c).Font.Bold = True
    row += 1

    for col_name, stats in summary["numeric_summary"].items():
        report.Cells(row, 1).Value = col_name
        report.Cells(row, 2).Value = stats["sum"]
        report.Cells(row, 3).Value = stats["avg"]
        report.Cells(row, 4).Value = stats["min"]
        report.Cells(row, 5).Value = stats["max"]
        report.Cells(row, 6).Value = stats["count"]
        row += 1

    report.Columns.AutoFit()


def write_csv_report(summary: dict, path: str) -> None:
    path = os.path.abspath(path)
    log.info("Writing CSV report: %s", path)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["row_count", summary["row_count"]])
        writer.writerow([])
        writer.writerow(["column", "sum", "avg", "min", "max", "count"])
        for col_name, stats in summary["numeric_summary"].items():
            writer.writerow(
                [col_name, stats["sum"], stats["avg"], stats["min"], stats["max"], stats["count"]]
            )


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------

def run(config: WorkerConfig) -> dict:
    with excel_application(visible=config.visible) as app:
        with open_workbook(app, config.input_path, read_only=config.read_only) as wb:
            records = read_sheet_as_records(
                wb, config.sheet_name, config.header_row, config.columns
            )
            summary = process_records(records)

            source_name = config.sheet_name or wb.ActiveSheet.Name
            write_report_sheet(wb, summary, source_name)

            if config.output_xlsx:
                out_path = os.path.abspath(config.output_xlsx)
                log.info("Saving report workbook: %s", out_path)
                wb.SaveAs(out_path, FileFormat=xlWorkbookDefault)

            if config.output_csv:
                write_csv_report(summary, config.output_csv)

    return summary


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def parse_args() -> WorkerConfig:
    p = argparse.ArgumentParser(description="Excel COM automation worker")
    p.add_argument("input_path", help="Path to the source .xlsx/.xlsm file")
    p.add_argument("--sheet", dest="sheet_name", default=None, help="Sheet name to read (default: active sheet)")
    p.add_argument("--header-row", type=int, default=1, help="1-based row number containing headers")
    p.add_argument("--columns", nargs="*", default=[], help="Specific column headers to extract (default: all)")
    p.add_argument("--output-xlsx", default=None, help="Path to save the workbook with the added Report sheet")
    p.add_argument("--output-csv", default=None, help="Path to write a CSV summary")
    p.add_argument("--visible", action="store_true", help="Show the Excel UI while running")
    p.add_argument("--writable", action="store_true", help="Open source file writable instead of read-only")
    args = p.parse_args()

    return WorkerConfig(
        input_path=args.input_path,
        sheet_name=args.sheet_name,
        header_row=args.header_row,
        output_xlsx=args.output_xlsx,
        output_csv=args.output_csv,
        visible=args.visible,
        read_only=not args.writable,
        columns=args.columns,
    )


def main() -> None:
    config = parse_args()
    start = time.time()
    try:
        summary = run(config)
    except Exception:
        log.exception("Worker failed.")
        sys.exit(1)

    elapsed = time.time() - start
    log.info("Done in %.1fs. Row count: %d", elapsed, summary["row_count"])
    for col, stats in summary["numeric_summary"].items():
        log.info("  %s: sum=%.2f avg=%.2f min=%.2f max=%.2f (n=%d)",
                  col, stats["sum"], stats["avg"], stats["min"], stats["max"], stats["count"])


if __name__ == "__main__":
    main()