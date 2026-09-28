import sys
from pathlib import Path
import time
import pandas as pd
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
csv_file = os.path.join(script_dir, "attendanceoutputs.csv")


# ---------- 1. SETTINGS (change these if your needs change) ----------
RAW_COLUMNS_SORT = ["Date", "Time"]   # raw tab is ordered by date, then time
WEEK_ENDS_ON = "SUN"                  # Monday-Sunday weeks end on Sunday


def load_raw(csv_path):
    """Read the CSV and add a combined date+time column used for sorting."""
    # dtype=str keeps everything as text, so employee codes like "00212"
    # never lose leading zeros.
    df = pd.read_csv(csv_path, dtype=str)

    # Remove accidental spaces around values.
    df = df.apply(lambda col: col.str.strip())

    # Join Date and Time into one real timestamp. errors="coerce" turns bad
    # rows into NaT (blank) instead of crashing the whole run.
    df["Timestamp"] = pd.to_datetime(
    df["Date"] + " " + df["Time"],
    format="%d-%m-%Y %H:%M:%S",
    errors="coerce",
)

    # Drop rows whose date/time could not be understood (and say so).
    bad = df["Timestamp"].isna().sum()
    if bad:
        print(f"Warning: skipped {bad} row(s) with an unreadable Date/Time")
    df = df.dropna(subset=["Timestamp"])

    # Sort chronologically once so every later step is in order.
    return df.sort_values("Timestamp").reset_index(drop=True)


def first_and_last(day_df):
    """For one day: keep each employee's first and latest sighting only."""
    grouped = day_df.groupby("EmployeeCode", sort=False)

    first = grouped.head(1).copy()      # earliest row per employee
    last = grouped.tail(1).copy()       # latest row per employee
    first["Record"] = "First In"
    last["Record"] = "Last Out"

    # If an employee was seen only once, first and last are the SAME row.
    # Keep it once (as "First In") so we do not invent a fake exit.
    last = last[~last.index.isin(first.index)]

    out = pd.concat([first, last]).sort_values(["EmployeeCode", "Timestamp"])

    cols = ["EmployeeCode", "Record", "Date", "Time", "Camera", "Event", "Similarity"]
    return out[cols]


def style_sheet(ws):
    """Make the tab readable: bold header, frozen top row, filters, widths."""
    from openpyxl.styles import Font, PatternFill, Alignment

    header_fill = PatternFill("solid", fgColor="1F3864")          # dark blue
    for cell in ws[1]:                                             # row 1 = header
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col in ws.columns:
        longest = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = max(12, longest + 3)


def build_workbooks(csv_path, out_dir):
    df = load_raw(csv_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Monday-Sunday weeks: pandas names a week by the day it ENDS on.
    df["Week"] = df["Timestamp"].dt.to_period(f"W-{WEEK_ENDS_ON}")

    for week, week_df in df.groupby("Week"):
        start = week.start_time.strftime("%Y-%m-%d")
        end = week.end_time.strftime("%Y-%m-%d")
        week_no = week.start_time.isocalendar().week
        file_path = out_dir / f"Week {week_no:02d}.xlsx"

        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            for date, day_df in week_df.groupby("Date", sort=False):
                # Tab 1: raw data for the day
                raw_cols = ["EmployeeCode", "Date", "Time", "Camera", "Event", "Similarity"]
                raw = day_df[raw_cols].copy()
                raw["Similarity"] = raw["Similarity"].astype(float)
                raw.to_excel(writer, sheet_name=f"{date} Raw", index=False)

                # Tab 2: first in / last out per employee
                io = first_and_last(day_df)
                io["Similarity"] = io["Similarity"].astype(float)
                io.to_excel(writer, sheet_name=f"{date} In-Out", index=False)

            for ws in writer.book.worksheets:
                style_sheet(ws)

        print(f"Created {file_path.name}  ({week_df['Date'].nunique()} days)")

INTERVAL_SECONDS = 300   # 5 minutes

if __name__ == "__main__":
    csv_arg = sys.argv[1] if len(sys.argv) > 1 else csv_file
    output_folder = sys.argv[2] if len(sys.argv) > 2 else os.path.join(script_dir, "weekly_attendance")

    while True:
        print(f"--- Run started {pd.Timestamp.now():%d-%m-%Y %H:%M:%S} ---")
        try:
            build_workbooks(csv_arg, output_folder)
        except Exception as e:
            # One failed run must not stop the loop. It retries in 5 minutes.
            print(f"Run failed: {e}")
        time.sleep(INTERVAL_SECONDS)
